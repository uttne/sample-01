import "../@types/jszip.t.js";

export class PowerApps {
  constructor() {
    /** @type {JSZip|null} */
    this._zip = null;
    /** @type {JSZip|null} */
    this._msapp = null;
    /** @type {File} */
    this._file = null;
  }

  /**
   *
   * @param {File} file
   */
  async loadAsync(file) {
    /** @type {JSZip} */
    const zip = new JSZip();
    try {
      await zip.loadAsync(file);
    } catch (e) {
      console.log(e);
      throw e;
    }

    const keys = Object.keys(zip.files);

    // TODO manifest ファイルの中身を確認して PowerApps の zip かを判定する

    const msappKey = keys.filter((key) => key.endsWith(".msapp"))[0];

    /** @type {JSZip} */
    const msapp = new JSZip();

    try {
      const blob = await zip.files[msappKey].async("blob");
      await msapp.loadAsync(blob);
    } catch (e) {
      console.log(e);
      throw e;
    }

    this._file = file;
    this._zip = zip;
    this._msapp = msapp;
  }
}
