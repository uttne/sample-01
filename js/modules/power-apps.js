import "../@types/jszip.t.js";

/**
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/AppCheckerResult.sarif
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/checksum.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/Controls/1.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/Controls/4.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/Header.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/Properties.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/References/DataSources.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/References/ModernThemes.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/References/Resources.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/References/Templates.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/References/Themes.json
 * N51005195-dd8a-4c56-aedc-4f068951aadd-document/Resources/PublishInfo.json
 */

/**
 * @typedef {Object} ManifestDetails
 * @property {string} displayName
 * @property {string} description
 * @property {string} createdTime
 * @property {string} packageTelemetryId
 * @property {string} creator
 * @property {string} sourceEnvironment
 */

/**
 * @typedef {Object} Manifest
 * @property {string} schema
 * @property {ManifestDetails} details
 */

/**
 * @typedef {Object} DataSource
 * @property {string} ApiId
 * @property {string} Name
 * @property {string} Type
 */

/**
 * @typedef {Object} DataSourcesJson
 * @property {DataSource[]} DataSources
 */

export class PowerApps {
  constructor() {
    /** @type {JSZip|null} */
    this._zip = null;
    /** @type {JSZip|null} */
    this._msapp = null;
    /** @type {File} */
    this._file = null;
    /** @type {Manifest} */
    this.manifest = null;
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

    const manifestKey = keys.filter((key) => key.endsWith("/manifest.json"))[0];
    if (!manifestKey) {
      throw new Error(
        "This is not PowerApps export file. : manifest.json is not found."
      );
    }

    const manifestJson = await zip.files[manifestKey].async("string");
    /** @type {Manifest} */
    const manifest = JSON.parse(manifestJson);

    if (!["1.0"].includes(manifest.schema)) {
      throw new Error(`Schema '${manifest.schema}' is not support.`);
    }

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
    this.manifest = manifest;
  }

  /**
   *
   * @returns {Promise<DataSource[]>}
   */
  async getDataSourcesAsync() {
    const msapp = this._msapp;
    if (!msapp) throw Error("This is not loaded.");

    const keys = Object.keys(msapp.files);
    const dataSourcesKey = keys.filter((key) =>
      key.endsWith("/References/DataSources.json")
    )[0];

    if (!dataSourcesKey) {
      throw Error("'DataSourses.json' is not found.");
    }

    const dataSourcesJson = await msapp.files[dataSourcesKey].async("string");
    /** @type {DataSourcesJson} */
    const dataSources = JSON.parse(dataSourcesJson);

    return dataSources.DataSources;
  }
}
