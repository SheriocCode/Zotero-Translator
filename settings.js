/* global Components */

var ZoteroTranslateSettings = {
  data: null,

  load() {
    this.data = window.arguments?.[0] || {};
    const config = this.data.config || {};

    for (const key of [
      "executablePath", "service", "sourceLang", "targetLang", "outputVariant", "extraArgs",
      "openaiBaseURL", "openaiAPIKey", "openaiModel", "pageMode", "customPages", "threads",
      "formulaFontRegex", "customPrompt"
    ]) {
      const element = document.getElementById(key);
      if (element) {
        element.value = config[key] || "";
      }
    }

    for (const key of ["skipFontSubsetting", "ignoreCache", "useBabelDOC"]) {
      const element = document.getElementById(key);
      if (element) {
        element.checked = !!config[key];
      }
    }
  },

  browseExecutable() {
    const picker = Components.classes["@mozilla.org/filepicker;1"]
      .createInstance(Components.interfaces.nsIFilePicker);

    picker.init(window, "Select PDFMathTranslate executable", Components.interfaces.nsIFilePicker.modeOpen);
    picker.appendFilters(Components.interfaces.nsIFilePicker.filterApps);
    picker.appendFilters(Components.interfaces.nsIFilePicker.filterAll);

    picker.open((result) => {
      if (result === Components.interfaces.nsIFilePicker.returnOK && picker.file) {
        document.getElementById("executablePath").value = picker.file.path;
      }
    });
  },

  save() {
    const config = {
      executablePath: document.getElementById("executablePath").value.trim() || "pdf2zh",
      service: "openai",
      sourceLang: document.getElementById("sourceLang").value.trim() || "en",
      targetLang: document.getElementById("targetLang").value.trim() || "zh-CN",
      outputVariant: document.getElementById("outputVariant").value || document.getElementById("outputVariant").selectedItem?.value || "dual",
      extraArgs: document.getElementById("extraArgs").value.trim(),
      openaiBaseURL: document.getElementById("openaiBaseURL").value.trim(),
      openaiAPIKey: document.getElementById("openaiAPIKey").value.trim(),
      openaiModel: document.getElementById("openaiModel").value.trim(),
      pageMode: document.getElementById("pageMode").value || document.getElementById("pageMode").selectedItem?.value || "all",
      customPages: document.getElementById("customPages").value.trim(),
      threads: document.getElementById("threads").value.trim() || "4",
      skipFontSubsetting: document.getElementById("skipFontSubsetting").checked,
      ignoreCache: document.getElementById("ignoreCache").checked,
      formulaFontRegex: document.getElementById("formulaFontRegex").value.trim(),
      customPrompt: document.getElementById("customPrompt").value.trim(),
      useBabelDOC: document.getElementById("useBabelDOC").checked
    };

    this.data.save?.(config);
    window.close();
  },

  cancel() {
    window.close();
  }
};
