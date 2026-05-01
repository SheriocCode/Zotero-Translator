/* global Zotero, Services */

var ZoteroTranslatePDF;

function log(message) {
  Zotero.debug(`Zotero Translate PDF: ${message}`);
}

function install() {
  log("installed");
}

async function startup({ id, version, rootURI }) {
  log(`starting ${version}`);

  Services.scriptloader.loadSubScript(rootURI + "zotero-translate.js");
  ZoteroTranslatePDF.init({ id, version, rootURI });
  ZoteroTranslatePDF.addToAllWindows();
  ZoteroTranslatePDF.retryAddToAllWindows();

  try {
    Zotero.PreferencePanes.register({
      pluginID: id,
      src: rootURI + "preferences.xhtml",
      label: "PDF Translate"
    });
  }
  catch (err) {
    Zotero.logError(err);
  }
}

function onMainWindowLoad({ window }) {
  ZoteroTranslatePDF.addToWindow(window);
}

function onMainWindowUnload({ window }) {
  ZoteroTranslatePDF.removeFromWindow(window);
}

function shutdown() {
  log("shutting down");
  ZoteroTranslatePDF?.removeFromAllWindows();
  ZoteroTranslatePDF = undefined;
}

function uninstall() {
  log("uninstalled");
}
