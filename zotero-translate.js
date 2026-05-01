/* global Zotero, Services, Components */

ZoteroTranslatePDF = {
  ADDON_ID: "zotero-translate@example.local",
  PREF_BRANCH: "extensions.zotero-translate.",
  MENU_ID: "zotero-translate-item-menu",
  TOOLS_MENU_ID: "zotero-translate-tools-menu",
  SETTINGS_MENU_ID: "zotero-translate-settings-menu",
  SETTINGS_PANEL_ID: "zotero-translate-settings-panel",
  RETRY_TIMER_ID: "zotero-translate-retry",

  id: null,
  version: null,
  rootURI: null,
  addedElementIDs: [],

  init({ id, version, rootURI }) {
    this.id = id;
    this.version = version;
    this.rootURI = rootURI;
  },

  log(message) {
    Zotero.debug(`Zotero Translate PDF: ${message}`);
  },

  addToAllWindows() {
    const windows = Zotero.getMainWindows ? Zotero.getMainWindows() : [];
    for (const win of windows) {
      if (win.ZoteroPane) {
        this.addToWindow(win);
      }
    }
  },

  retryAddToAllWindows() {
    Zotero.Promise.delay(500).then(() => this.addToAllWindows());
    Zotero.Promise.delay(1500).then(() => this.addToAllWindows());
    Zotero.Promise.delay(3000).then(() => this.addToAllWindows());
  },

  addToWindow(window) {
    const doc = window.document;

    const itemMenu = doc.getElementById("zotero-itemmenu")
      || doc.getElementById("zotero-itemmenu-popup")
      || doc.querySelector("#zotero-itemmenu, #zotero-itemmenu-popup");

    if (itemMenu && !doc.getElementById(this.MENU_ID)) {
      const menuItem = doc.createXULElement("menuitem");
      menuItem.id = this.MENU_ID;
      menuItem.setAttribute("label", "Translate PDF with PDFMathTranslate");
      menuItem.addEventListener("command", () => this.translateSelectedPDFs(window));
      itemMenu.appendChild(menuItem);
      this.storeAddedElement(menuItem);
    }

    const toolsPopup = doc.getElementById("menu_ToolsPopup")
      || doc.getElementById("menu-tools-popup")
      || doc.querySelector("#menu_ToolsPopup, #menu-tools-popup");

    if (toolsPopup && !doc.getElementById(this.TOOLS_MENU_ID)) {
      const separator = doc.createXULElement("menuseparator");
      separator.id = `${this.TOOLS_MENU_ID}-separator`;

      const translateItem = doc.createXULElement("menuitem");
      translateItem.id = this.TOOLS_MENU_ID;
      translateItem.setAttribute("label", "Translate Selected Zotero PDFs");
      translateItem.addEventListener("command", () => this.translateSelectedPDFs(window));

      const settingsItem = doc.createXULElement("menuitem");
      settingsItem.id = this.SETTINGS_MENU_ID;
      settingsItem.setAttribute("label", "Configure PDF Translation Backend");
      settingsItem.addEventListener("command", () => this.configureBackend(window));

      toolsPopup.appendChild(separator);
      toolsPopup.appendChild(translateItem);
      toolsPopup.appendChild(settingsItem);
      this.storeAddedElement(separator);
      this.storeAddedElement(translateItem);
      this.storeAddedElement(settingsItem);
    }
  },

  storeAddedElement(element) {
    if (!this.addedElementIDs.includes(element.id)) {
      this.addedElementIDs.push(element.id);
    }
  },

  removeFromWindow(window) {
    const doc = window.document;
    for (const id of this.addedElementIDs) {
      doc.getElementById(id)?.remove();
    }
  },

  removeFromAllWindows() {
    const windows = Zotero.getMainWindows ? Zotero.getMainWindows() : [];
    for (const win of windows) {
      if (win.ZoteroPane) {
        this.removeFromWindow(win);
      }
    }
  },

  async translateSelectedPDFs(window) {
    try {
      const items = this.getSelectedItems(window);
      const pdfAttachments = [];

      for (const item of items) {
        if (item.isAttachment && item.isAttachment()) {
          if (this.isPDFAttachment(item)) {
            pdfAttachments.push(item);
          }
          continue;
        }

        if (item.isRegularItem && item.isRegularItem()) {
          const attachments = await item.getAttachments();
          for (const attachmentID of attachments) {
            const attachment = Zotero.Items.get(attachmentID);
            if (attachment && this.isPDFAttachment(attachment)) {
              pdfAttachments.push(attachment);
            }
          }
        }
      }

      if (!pdfAttachments.length) {
        this.alert(window, "No PDF attachments selected", "Select one or more PDF attachments, or regular items that contain PDF attachments.");
        return;
      }

      const config = this.getConfig();
      const progress = this.createProgress("Translating PDFs", `0/${pdfAttachments.length} finished`);
      const failures = [];

      for (let i = 0; i < pdfAttachments.length; i++) {
        const attachment = pdfAttachments[i];
        const sourcePath = await attachment.getFilePathAsync();
        if (!sourcePath) {
          failures.push(`${attachment.getField("title") || attachment.id}: local PDF file not found`);
          continue;
        }

        progress.changeHeadline?.("Translating PDFs");
        progress.addDescription?.(`${i + 1}/${pdfAttachments.length}: ${attachment.getField("title") || sourcePath}`);

        try {
          const outputDir = await this.getOutputDirectory(attachment.id);
          const result = await this.runPdfMathTranslate(config, sourcePath, outputDir);
          const translatedPath = await this.resolveTranslatedPDFPath(sourcePath, outputDir, config.outputVariant, result);
          const finalPath = this.renameTranslatedPDFForImport(translatedPath, sourcePath, config);
          const imported = await this.importTranslatedPDF(attachment, finalPath, config.outputVariant);
          this.log(`imported translated attachment ${imported.id}`);
        }
        catch (err) {
          Zotero.logError(err);
          failures.push(`${attachment.getField("title") || attachment.id}: ${err.message || err}`);
        }
      }

      progress.close?.();

      if (failures.length) {
        this.alert(window, "PDF translation finished with errors", failures.join("\n"));
      }
      else {
        this.alert(window, "PDF translation finished", `Translated ${pdfAttachments.length} PDF attachment(s).`);
      }
    }
    catch (err) {
      Zotero.logError(err);
      this.alert(window, "PDF translation failed", err.message || String(err));
    }
  },

  getSelectedItems(window) {
    if (window.ZoteroPane?.getSelectedItems) {
      return window.ZoteroPane.getSelectedItems();
    }
    return [];
  },

  isPDFAttachment(item) {
    return item.attachmentContentType === "application/pdf"
      || /\.pdf$/i.test(item.getField?.("title") || "")
      || /\.pdf$/i.test(item.attachmentFilename || "");
  },

  getConfig() {
    return {
      executablePath: Zotero.Prefs.get(`${this.PREF_BRANCH}executablePath`, true) || "pdf2zh",
      service: Zotero.Prefs.get(`${this.PREF_BRANCH}service`, true) || "openai",
      sourceLang: Zotero.Prefs.get(`${this.PREF_BRANCH}sourceLang`, true) || "en",
      targetLang: Zotero.Prefs.get(`${this.PREF_BRANCH}targetLang`, true) || "zh-CN",
      outputVariant: Zotero.Prefs.get(`${this.PREF_BRANCH}outputVariant`, true) || "dual",
      extraArgs: Zotero.Prefs.get(`${this.PREF_BRANCH}extraArgs`, true) || "",
      openaiBaseURL: Zotero.Prefs.get(`${this.PREF_BRANCH}openaiBaseURL`, true) || "https://api.deepseek.com",
      openaiAPIKey: Zotero.Prefs.get(`${this.PREF_BRANCH}openaiAPIKey`, true) || "sk-b14040917d2041efb361eaf23e927517",
      openaiModel: Zotero.Prefs.get(`${this.PREF_BRANCH}openaiModel`, true) || "deepseek-chat",
      pageMode: Zotero.Prefs.get(`${this.PREF_BRANCH}pageMode`, true) || "all",
      customPages: Zotero.Prefs.get(`${this.PREF_BRANCH}customPages`, true) || "",
      threads: Zotero.Prefs.get(`${this.PREF_BRANCH}threads`, true) || "4",
      skipFontSubsetting: !!Zotero.Prefs.get(`${this.PREF_BRANCH}skipFontSubsetting`, true),
      ignoreCache: !!Zotero.Prefs.get(`${this.PREF_BRANCH}ignoreCache`, true),
      formulaFontRegex: Zotero.Prefs.get(`${this.PREF_BRANCH}formulaFontRegex`, true) || "",
      customPrompt: Zotero.Prefs.get(`${this.PREF_BRANCH}customPrompt`, true) || "",
      useBabelDOC: !!Zotero.Prefs.get(`${this.PREF_BRANCH}useBabelDOC`, true)
    };
  },

  setConfig(config) {
    for (const [key, value] of Object.entries(config)) {
      Zotero.Prefs.set(`${this.PREF_BRANCH}${key}`, value, true);
    }
  },

  configureBackend(window) {
    this.showSettingsPanel(window);
  },

  showSettingsPanel(window) {
    const doc = window.document;
    doc.getElementById(this.SETTINGS_PANEL_ID)?.remove();

    const config = this.getConfig();
    const overlay = this.html(doc, "div", {
      id: this.SETTINGS_PANEL_ID,
      style: [
        "position: fixed",
        "inset: 0",
        "z-index: 2147483647",
        "background: rgba(32, 35, 39, 0.32)",
        "display: flex",
        "align-items: center",
        "justify-content: center",
        "font: menu",
        "padding: 28px"
      ].join(";")
    });

    const panel = this.html(doc, "div", {
      style: [
        "width: min(820px, calc(100vw - 56px))",
        "max-height: calc(100vh - 56px)",
        "overflow: auto",
        "box-sizing: border-box",
        "background: #fff",
        "border: 1px solid #d0d7de",
        "border-radius: 8px",
        "box-shadow: 0 22px 70px rgba(0, 0, 0, 0.26)",
        "color: #1f2328"
      ].join(";")
    });

    const header = this.html(doc, "div", {
      style: [
        "padding: 18px 20px 14px",
        "border-bottom: 1px solid #d8dee4",
        "display: flex",
        "align-items: flex-start",
        "justify-content: space-between",
        "gap: 16px"
      ].join(";")
    });

    const headingWrap = this.html(doc, "div", {});
    headingWrap.append(
      this.html(doc, "div", {
        textContent: "PDFMathTranslate Backend",
        style: "font-size: 20px; font-weight: 650; margin-bottom: 4px;"
      }),
      this.html(doc, "div", {
        textContent: "Configure the local translator used by Zotero PDF attachments.",
        style: "color: #57606a; line-height: 1.35;"
      })
    );

    header.append(headingWrap);

    const body = this.html(doc, "div", {
      style: "padding: 18px 20px 4px;"
    });

    const inputs = {};
    const general = this.createSettingsSection(doc, "Backend", "Choose the local PDFMathTranslate executable and default languages.");
    this.addTextField(doc, general, inputs, "executablePath", "Executable", config.executablePath, {
      buttonText: "Browse",
      buttonHandler: () => {
        this.browseExecutable(window, inputs.executablePath);
      }
    });
    this.addFieldNote(doc, general, "Use the full path to pdf2zh.exe for the Windows packaged build.");
    this.addTextField(doc, general, inputs, "sourceLang", "Source language", config.sourceLang);
    this.addTextField(doc, general, inputs, "targetLang", "Target language", config.targetLang);
    this.addSelectField(doc, general, inputs, "outputVariant", "Imported PDF", config.outputVariant, [
      ["dual", "dual bilingual PDF"],
      ["mono", "mono translated PDF"]
    ]);
    this.addSelectField(doc, general, inputs, "pageMode", "Pages", config.pageMode, [
      ["all", "All"],
      ["first", "First"],
      ["first5", "First 5 pages"],
      ["custom", "Others"]
    ]);
    const customPagesRow = this.addTextField(doc, general, inputs, "customPages", "Custom pages", config.customPages, {
      placeholder: "for example: 1,3,5-8"
    });

    const service = this.createSettingsSection(doc, "Translation Service", "Only OpenAI compatible endpoints are exposed. DeepSeek can be used through OPENAI_BASE_URL.");
    this.addSelectField(doc, service, inputs, "service", this.normalizeService(config.service), [
      ["openai", "OpenAI compatible"]
    ]);
    this.addTextField(doc, service, inputs, "openaiBaseURL", "OPENAI_BASE_URL", config.openaiBaseURL, { placeholder: "https://api.deepseek.com" });
    this.addTextField(doc, service, inputs, "openaiAPIKey", "OPENAI_API_KEY", config.openaiAPIKey, { type: "password" });
    this.addTextField(doc, service, inputs, "openaiModel", "OPENAI_MODEL", config.openaiModel, { placeholder: "deepseek-chat" });
    this.addFieldNote(doc, service, "DeepSeek is OpenAI compatible: use Service = OpenAI compatible, OPENAI_BASE_URL = https://api.deepseek.com, OPENAI_MODEL = deepseek-chat.");

    const experimental = this.createSettingsSection(doc, "Experimental", "Options matching the PDFMathTranslate web UI.");
    this.addTextField(doc, experimental, inputs, "threads", "number of threads", config.threads, {
      type: "number",
      placeholder: "4"
    });
    this.addCheckboxField(doc, experimental, inputs, "skipFontSubsetting", "Skip font subsetting", config.skipFontSubsetting);
    this.addCheckboxField(doc, experimental, inputs, "ignoreCache", "Ignore cache", config.ignoreCache);
    this.addTextField(doc, experimental, inputs, "formulaFontRegex", "Custom formula font regex (vfont)", config.formulaFontRegex, {
      placeholder: "for example: (CM.*|MS.*)"
    });
    this.addTextareaField(doc, experimental, inputs, "customPrompt", "Custom Prompt for llm", config.customPrompt, {
      placeholder: "Optional prompt text. It will be written to a temporary --prompt file."
    });
    this.addCheckboxField(doc, experimental, inputs, "useBabelDOC", "Use BabelDOC", config.useBabelDOC);

    const advanced = this.createSettingsSection(doc, "Advanced", "Optional flags appended to the pdf2zh command.");
    this.addTextField(doc, advanced, inputs, "extraArgs", "Extra arguments", config.extraArgs, {
      placeholder: "--thread 2 --ignore-cache"
    });

    body.append(general.section, service.section, experimental.section, advanced.section);

    const actions = this.html(doc, "div", {
      style: [
        "display: flex",
        "justify-content: space-between",
        "align-items: center",
        "gap: 12px",
        "padding: 14px 20px 18px",
        "border-top: 1px solid #d8dee4",
        "background: #f6f8fa",
        "border-radius: 0 0 8px 8px"
      ].join(";")
    });

    const status = this.html(doc, "div", {
      textContent: "Settings are saved locally in Zotero preferences.",
      style: "color: #57606a; font-size: 12px;"
    });

    const buttonGroup = this.html(doc, "div", { style: "display: flex; gap: 8px;" });
    const cancelButton = this.html(doc, "button", {
      textContent: "Cancel",
      type: "button",
      style: this.buttonStyle("secondary")
    });
    cancelButton.addEventListener("click", () => overlay.remove());

    const saveButton = this.html(doc, "button", {
      textContent: "Save",
      type: "button",
      style: this.buttonStyle("primary")
    });
    saveButton.addEventListener("click", () => {
      this.setConfig({
        executablePath: inputs.executablePath.value.trim() || "pdf2zh",
        service: "openai",
        sourceLang: inputs.sourceLang.value.trim() || "en",
        targetLang: inputs.targetLang.value.trim() || "zh-CN",
        outputVariant: inputs.outputVariant.value || "dual",
        extraArgs: inputs.extraArgs.value.trim(),
        openaiBaseURL: inputs.openaiBaseURL.value.trim(),
        openaiAPIKey: inputs.openaiAPIKey.value.trim(),
        openaiModel: inputs.openaiModel.value.trim(),
        pageMode: inputs.pageMode.value || "all",
        customPages: inputs.customPages.value.trim(),
        threads: inputs.threads.value.trim() || "4",
        skipFontSubsetting: inputs.skipFontSubsetting.checked,
        ignoreCache: inputs.ignoreCache.checked,
        formulaFontRegex: inputs.formulaFontRegex.value.trim(),
        customPrompt: inputs.customPrompt.value.trim(),
        useBabelDOC: inputs.useBabelDOC.checked
      });
      overlay.remove();
      this.alert(window, "Backend configured", "PDF translation settings have been saved.");
    });

    buttonGroup.append(cancelButton, saveButton);
    actions.append(status, buttonGroup);

    const updateVisibility = () => {
      this.setRowsVisible([customPagesRow], inputs.pageMode.value === "custom");
    };
    inputs.pageMode.addEventListener("change", updateVisibility);
    updateVisibility();

    overlay.addEventListener("click", (event) => {
      if (event.target === overlay) {
        overlay.remove();
      }
    });

    panel.append(header, body, actions);
    overlay.append(panel);
    doc.documentElement.appendChild(overlay);
  },

  createSettingsSection(doc, title, description) {
    const section = this.html(doc, "section", {
      style: [
        "border: 1px solid #d8dee4",
        "border-radius: 8px",
        "margin-bottom: 14px",
        "background: #fff"
      ].join(";")
    });
    const header = this.html(doc, "div", {
      style: "padding: 13px 14px 10px; border-bottom: 1px solid #edf0f2;"
    });
    header.append(
      this.html(doc, "div", {
        textContent: title,
        style: "font-weight: 650; margin-bottom: 3px;"
      }),
      this.html(doc, "div", {
        textContent: description,
        style: "color: #57606a; font-size: 12px; line-height: 1.35;"
      })
    );
    const grid = this.html(doc, "div", {
      style: [
        "display: grid",
        "grid-template-columns: 150px minmax(260px, 1fr) auto",
        "gap: 10px",
        "align-items: center",
        "padding: 14px"
      ].join(";")
    });
    section.append(header, grid);
    return { section, grid };
  },

  addTextField(doc, target, inputs, key, labelText, value, options = {}) {
    const grid = target.grid || target;
    const label = this.html(doc, "label", { textContent: labelText, style: "color: #24292f;" });
    const input = this.html(doc, "input", {
      type: options.type || "text",
      value: value || "",
      placeholder: options.placeholder || "",
      style: [
        "box-sizing: border-box",
        "width: 100%",
        "min-height: 32px",
        "padding: 5px 8px",
        "border: 1px solid #d0d7de",
        "border-radius: 6px",
        "background: #fff",
        "font: menu"
      ].join(";")
    });
    const button = options.buttonText
      ? this.html(doc, "button", { type: "button", textContent: options.buttonText, style: this.buttonStyle("secondary") })
      : this.html(doc, "span", {});

    if (options.buttonHandler) {
      button.addEventListener("click", options.buttonHandler);
    }

    inputs[key] = input;
    grid.append(label, input, button);
    return [label, input, button];
  },

  addSelectField(doc, target, inputs, key, labelTextOrValue, valueOrOptions, maybeOptions) {
    const grid = target.grid || target;
    const compactSignature = Array.isArray(valueOrOptions);
    const labelText = compactSignature ? "Service" : labelTextOrValue;
    const value = compactSignature ? labelTextOrValue : valueOrOptions;
    const options = compactSignature ? valueOrOptions : maybeOptions;
    const label = this.html(doc, "label", { textContent: labelText, style: "color: #24292f;" });
    const select = this.html(doc, "select", {
      style: [
        "box-sizing: border-box",
        "width: 100%",
        "min-height: 32px",
        "padding: 5px 8px",
        "border: 1px solid #d0d7de",
        "border-radius: 6px",
        "background: #fff",
        "font: menu"
      ].join(";")
    });
    for (const [optionValue, optionLabel] of options) {
      const option = this.html(doc, "option", { value: optionValue, textContent: optionLabel });
      if (optionValue === value) {
        option.selected = true;
      }
      select.appendChild(option);
    }
    inputs[key] = select;
    const spacer = this.html(doc, "span", {});
    grid.append(label, select, spacer);
    return [label, select, spacer];
  },

  addFieldNote(doc, target, text) {
    const grid = target.grid || target;
    grid.append(
      this.html(doc, "span", {}),
      this.html(doc, "div", {
        textContent: text,
        style: "grid-column: span 2; color: #57606a; font-size: 12px; margin-top: -6px; margin-bottom: 2px; line-height: 1.35;"
      })
    );
  },

  addCheckboxField(doc, target, inputs, key, labelText, checked) {
    const grid = target.grid || target;
    const label = this.html(doc, "label", { textContent: labelText, style: "color: #24292f;" });
    const input = this.html(doc, "input", {
      type: "checkbox",
      checked: !!checked,
      style: "width: 16px; height: 16px;"
    });
    const spacer = this.html(doc, "span", {});
    inputs[key] = input;
    grid.append(label, input, spacer);
    return [label, input, spacer];
  },

  addTextareaField(doc, target, inputs, key, labelText, value, options = {}) {
    const grid = target.grid || target;
    const label = this.html(doc, "label", { textContent: labelText, style: "color: #24292f; align-self: start; padding-top: 7px;" });
    const textarea = this.html(doc, "textarea", {
      value: value || "",
      placeholder: options.placeholder || "",
      rows: options.rows || 4,
      style: [
        "box-sizing: border-box",
        "width: 100%",
        "min-height: 82px",
        "padding: 5px 8px",
        "border: 1px solid #d0d7de",
        "border-radius: 6px",
        "background: #fff",
        "font: menu",
        "resize: vertical"
      ].join(";")
    });
    const spacer = this.html(doc, "span", {});
    inputs[key] = textarea;
    grid.append(label, textarea, spacer);
    return [label, textarea, spacer];
  },

  setRowsVisible(rows, visible) {
    for (const row of rows) {
      for (const element of row) {
        element.style.display = visible ? "" : "none";
      }
    }
  },

  buttonStyle(kind) {
    const base = [
      "min-height: 32px",
      "padding: 5px 12px",
      "border-radius: 6px",
      "font: menu",
      "cursor: pointer"
    ];
    if (kind === "primary") {
      return [...base, "border: 1px solid #1f883d", "background: #1f883d", "color: #fff"].join(";");
    }
    return [...base, "border: 1px solid #d0d7de", "background: #f6f8fa", "color: #24292f"].join(";");
  },

  browseExecutable(window, input) {
    const picker = Components.classes["@mozilla.org/filepicker;1"]
      .createInstance(Components.interfaces.nsIFilePicker);

    picker.init(window, "Select PDFMathTranslate executable", Components.interfaces.nsIFilePicker.modeOpen);
    picker.appendFilters(Components.interfaces.nsIFilePicker.filterApps);
    picker.appendFilters(Components.interfaces.nsIFilePicker.filterAll);

    picker.open((result) => {
      if (result === Components.interfaces.nsIFilePicker.returnOK && picker.file) {
        input.value = picker.file.path;
      }
    });
  },

  html(doc, tagName, properties = {}) {
    const element = doc.createElementNS("http://www.w3.org/1999/xhtml", tagName);
    for (const [key, value] of Object.entries(properties)) {
      if (key in element) {
        element[key] = value;
      }
      else {
        element.setAttribute(key, value);
      }
    }
    return element;
  },

  async runPdfMathTranslate(config, sourcePath, outputDir) {
    const logPath = `${outputDir}\\pdf2zh.log`;
    const configPath = `${outputDir}\\pdf2zh-config.json`;
    this.writeTextFile(configPath, "{}\n");
    const service = this.normalizeService(config.service);
    const args = [
      this.cleanPath(sourcePath),
      "-li", config.sourceLang,
      "-lo", config.targetLang,
      "-s", service,
      "-o", outputDir,
      "--config", this.cleanPath(configPath)
    ];

    const pageArgument = this.getPageArgument(config);
    if (pageArgument) {
      args.push("-p", pageArgument);
    }

    const threads = parseInt(config.threads, 10);
    if (Number.isFinite(threads) && threads > 0) {
      args.push("-t", String(threads));
    }

    if (config.skipFontSubsetting) {
      args.push("--skip-subset-fonts");
    }

    if (config.ignoreCache) {
      args.push("--ignore-cache");
    }

    if (config.formulaFontRegex) {
      args.push("-f", config.formulaFontRegex);
    }

    if (config.customPrompt) {
      const promptPath = `${outputDir}\\pdf2zh-prompt.txt`;
      this.writeTextFile(promptPath, config.customPrompt);
      args.push("--prompt", this.cleanPath(promptPath));
    }

    if (config.useBabelDOC) {
      args.push("--babeldoc");
    }

    for (const arg of this.splitArgs(config.extraArgs)) {
      args.push(arg);
    }

    return this.runProcess(config, args, logPath);
  },

  getPageArgument(config) {
    switch (config.pageMode) {
      case "first":
        return "1";
      case "first5":
        return "1-5";
      case "custom":
        return String(config.customPages || "").trim();
      case "all":
      default:
        return "";
    }
  },

  normalizeService(service) {
    const value = String(service || "").trim();
    const aliases = {
      OpenAI: "openai",
      DeepSeek: "openai",
      Deepseek: "openai",
      Google: "openai",
      Bing: "openai",
      DeepL: "openai",
      DeepLX: "openai"
    };
    return aliases[value] || "openai";
  },

  runProcess(config, args, logPath) {
    return new Promise((resolve, reject) => {
      const process = Components.classes["@mozilla.org/process/util;1"].createInstance(Components.interfaces.nsIProcess);
      const envSnapshot = this.applyProcessEnvironment(config);
      const launch = this.buildLaunch(config, args, logPath);

      process.init(launch.executable);
      process.runAsync(launch.args, launch.args.length, {
        observe: async () => {
          this.restoreProcessEnvironment(envSnapshot);
          if (process.exitValue === 0) {
            resolve({ exitValue: process.exitValue, logPath });
          }
          else {
            await Zotero.Promise.delay(800);
            const logText = await this.readTextFile(logPath);
            const tail = this.tail(logText, 3000);
            const message = [
              `PDFMathTranslate exited with code ${process.exitValue}.`,
              `Command: ${launch.displayCommand}`,
              tail ? `Log:\n${tail}` : `Log file was empty or unavailable: ${logPath}`
            ].join("\n\n");
            reject(new Error(message));
          }
        }
      }, false);
    });
  },

  buildLaunch(config, args, logPath) {
    if (Services.appinfo.OS === "WINNT") {
      return {
        executable: this.localFile(config.executablePath),
        args,
        displayCommand: `${this.windowsQuote(config.executablePath)} ${args.map(this.windowsQuote).join(" ")}`
      };
    }

    const commandLine = this.buildCommandLine(config, args, logPath);
    return {
      executable: this.getShellExecutable(),
      args: this.getShellCommandArgs(commandLine),
      displayCommand: commandLine
    };
  },

  getShellExecutable() {
    if (Services.appinfo.OS === "WINNT") {
      return this.localFile(`${Services.dirsvc.get("WinD", Components.interfaces.nsIFile).path}\\System32\\cmd.exe`);
    }
    return this.localFile("/bin/sh");
  },

  getShellCommandArgs(commandLine) {
    if (Services.appinfo.OS === "WINNT") {
      return ["/d", "/s", "/c", commandLine];
    }
    return ["-lc", commandLine];
  },

  buildCommandLine(config, args, logPath) {
    const workingDir = this.getExecutableDirectory(config.executablePath);
    if (Services.appinfo.OS === "WINNT") {
      const env = this.buildWindowsEnvPrefix(config);
      const parts = [this.windowsQuote(config.executablePath), ...args.map(this.windowsQuote)];
      const cd = workingDir ? `cd /d ${this.windowsQuote(workingDir)} && ` : "";
      return `${cd}${env}${parts.join(" ")} > ${this.windowsQuote(logPath)} 2>&1`;
    }
    const env = this.buildUnixEnvPrefix(config);
    const parts = [this.shellQuote(config.executablePath), ...args.map(this.shellQuote)];
    const cd = workingDir ? `cd ${this.shellQuote(workingDir)} && ` : "";
    return `${cd}${env}${parts.join(" ")} > ${this.shellQuote(logPath)} 2>&1`;
  },

  getExecutableDirectory(executablePath) {
    if (!/[\\/]/.test(executablePath)) {
      return "";
    }
    return String(executablePath).replace(/[\\/][^\\/]+$/, "");
  },

  buildWindowsEnvPrefix(config) {
    const vars = this.getBackendEnv(config);
    return vars.map(([key, value]) => `set "${key}=${String(value).replace(/"/g, '""')}" && `).join("");
  },

  buildUnixEnvPrefix(config) {
    const vars = this.getBackendEnv(config);
    return vars.map(([key, value]) => `${key}=${this.shellQuote(value)} `).join("");
  },

  getBackendEnv(config) {
    const vars = [];
    if (config.openaiBaseURL) {
      vars.push(["OPENAI_BASE_URL", config.openaiBaseURL]);
    }
    if (config.openaiAPIKey) {
      vars.push(["OPENAI_API_KEY", config.openaiAPIKey]);
    }
    if (config.openaiModel) {
      vars.push(["OPENAI_MODEL", config.openaiModel]);
    }
    return vars;
  },

  applyProcessEnvironment(config) {
    const environment = Components.classes["@mozilla.org/process/environment;1"]
      .getService(Components.interfaces.nsIEnvironment);
    const snapshot = [];
    for (const [key, value] of this.getBackendEnv(config)) {
      snapshot.push([key, environment.exists(key), environment.exists(key) ? environment.get(key) : ""]);
      environment.set(key, value);
    }
    return snapshot;
  },

  restoreProcessEnvironment(snapshot) {
    const environment = Components.classes["@mozilla.org/process/environment;1"]
      .getService(Components.interfaces.nsIEnvironment);
    for (const [key, existed, value] of snapshot) {
      if (existed) {
        environment.set(key, value);
      }
      else {
        environment.set(key, "");
      }
    }
  },

  windowsQuote(value) {
    return `"${String(value).replace(/"/g, '""')}"`;
  },

  shellQuote(value) {
    return `'${String(value).replace(/'/g, "'\\''")}'`;
  },

  async readTextFile(path) {
    try {
      const file = this.localFile(path);
      if (!file.exists()) {
        return "";
      }

      const fileInputStream = Components.classes["@mozilla.org/network/file-input-stream;1"]
        .createInstance(Components.interfaces.nsIFileInputStream);
      const converterInputStream = Components.classes["@mozilla.org/intl/converter-input-stream;1"]
        .createInstance(Components.interfaces.nsIConverterInputStream);

      fileInputStream.init(file, 0x01, 0o444, 0);
      converterInputStream.init(fileInputStream, "UTF-8", 0, 0);

      const chunks = [];
      const out = {};
      while (converterInputStream.readString(4096, out) !== 0) {
        chunks.push(out.value);
      }
      converterInputStream.close();
      fileInputStream.close();
      return chunks.join("");
    }
    catch (err) {
      Zotero.logError(err);
      return "";
    }
  },

  tail(text, length) {
    if (!text) {
      return "";
    }
    return text.length > length ? text.slice(text.length - length) : text;
  },

  writeTextFile(path, text) {
    const file = this.localFile(path);
    const stream = Components.classes["@mozilla.org/network/file-output-stream;1"]
      .createInstance(Components.interfaces.nsIFileOutputStream);
    stream.init(file, 0x02 | 0x08 | 0x20, 0o644, 0);
    stream.write(text, text.length);
    stream.close();
  },

  splitArgs(value) {
    if (!value || !value.trim()) {
      return [];
    }
    const matches = value.match(/"([^"]*)"|'([^']*)'|[^\s]+/g) || [];
    return matches.map((part) => {
      if ((part.startsWith('"') && part.endsWith('"')) || (part.startsWith("'") && part.endsWith("'"))) {
        return part.slice(1, -1);
      }
      return part;
    });
  },

  async getOutputDirectory(attachmentID) {
    const root = Services.dirsvc.get("TmpD", Components.interfaces.nsIFile);
    root.append("zotero-translate-pdfmathtranslate");
    this.ensureDirectory(root);

    const dir = root.clone();
    dir.append(`${attachmentID}-${Date.now()}`);
    this.ensureDirectory(dir);
    return dir.path;
  },

  ensureDirectory(dir) {
    if (dir.exists()) {
      return;
    }
    const parent = dir.parent;
    if (parent && !parent.exists()) {
      this.ensureDirectory(parent);
    }
    dir.create(Components.interfaces.nsIFile.DIRECTORY_TYPE, 0o755);
  },

  async resolveTranslatedPDFPath(sourcePath, outputDir, variant) {
    const basename = sourcePath.replace(/^.*[\\/]/, "").replace(/\.pdf$/i, "");
    const suffix = variant === "mono" ? "mono" : "dual";
    const candidates = [
      `${outputDir}\\${basename}-${suffix}.pdf`,
      `${outputDir}/${basename}-${suffix}.pdf`,
      `${outputDir}\\${basename}.${suffix}.pdf`,
      `${outputDir}/${basename}.${suffix}.pdf`
    ];

    for (const candidate of candidates) {
      const file = this.localFile(candidate);
      if (file.exists()) {
        return file.path;
      }
    }

    const found = this.findNewestPDF(outputDir);
    if (found) {
      return found.path;
    }

    throw new Error(`Translated ${suffix} PDF was not found in ${outputDir}`);
  },

  findNewestPDF(outputDir) {
    const dir = this.localFile(outputDir);
    const entries = dir.directoryEntries;
    let newest = null;
    while (entries.hasMoreElements()) {
      const file = entries.getNext().QueryInterface(Components.interfaces.nsIFile);
      if (file.isFile() && /\.pdf$/i.test(file.leafName)) {
        if (!newest || file.lastModifiedTime > newest.lastModifiedTime) {
          newest = file;
        }
      }
    }
    return newest;
  },

  renameTranslatedPDFForImport(translatedPath, sourcePath, config) {
    const sourceFile = this.localFile(translatedPath);
    const parent = sourceFile.parent;
    const targetName = this.getTranslatedOutputFilename(sourcePath, config);

    if (sourceFile.leafName === targetName) {
      return sourceFile.path;
    }

    const targetFile = parent.clone();
    targetFile.append(targetName);
    if (targetFile.exists()) {
      targetFile.remove(false);
    }

    sourceFile.moveTo(parent, targetName);
    return sourceFile.path;
  },

  getTranslatedOutputFilename(sourcePath, config) {
    const inputName = this.getLeafName(sourcePath).replace(/\.pdf$/i, "") || "translated";
    const variantLabel = config.outputVariant === "mono" ? "mono translation" : "dual translation";
    const pageLabel = this.getPageFilenameLabel(config);
    return `(${variantLabel})${pageLabel}${this.sanitizeFilename(inputName)}.pdf`;
  },

  getPageFilenameLabel(config) {
    const pageArgument = this.getPageArgument(config);
    if (!pageArgument) {
      return "";
    }
    return `(${this.sanitizeFilename(pageArgument)})`;
  },

  getLeafName(path) {
    return String(path || "").replace(/^.*[\\/]/, "");
  },

  sanitizeFilename(value) {
    return String(value || "")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/[\x00-\x1f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  },

  async importTranslatedPDF(sourceAttachment, translatedPath, variant) {
    const file = this.localFile(translatedPath);
    const title = file.leafName.replace(/\.pdf$/i, "");
    const options = {
      file,
      title: title || `${sourceAttachment.getField("title") || "PDF"} (${variant} translation)`
    };

    if (sourceAttachment.parentItemID) {
      options.parentItemID = sourceAttachment.parentItemID;
    }
    else {
      options.collections = sourceAttachment.getCollections();
    }

    const imported = await Zotero.Attachments.importFromFile(options);
    await imported.saveTx();
    return imported;
  },

  localFile(path) {
    const file = Components.classes["@mozilla.org/file/local;1"].createInstance(Components.interfaces.nsIFile);
    file.initWithPath(this.cleanPath(path));
    return file;
  },

  cleanPath(path) {
    const cleaned = String(path || "")
      .trim()
      .replace(/^["']+|["']+$/g, "");
    if (Services.appinfo.OS === "WINNT") {
      return cleaned.replace(/\//g, "\\");
    }
    return cleaned;
  },

  createProgress(headline, description) {
    const progress = new Zotero.ProgressWindow();
    progress.changeHeadline(headline);
    progress.addDescription(description);
    progress.show();
    return progress;
  },

  alert(window, title, message) {
    Services.prompt.alert(window, title, message);
  }
};
