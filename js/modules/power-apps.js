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
 * @typedef {Object} DataSourceBase
 * @property {string} ApiId
 * @property {string} Name
 * @property {string} Type
 */

/**
 * @typedef {Object} WadlMetadata
 * @property {string} WadlXml
 */

/**
 * @typedef {Object} FlowDataSource
 * @property {string} ApiId
 * @property {string} Name
 * @property {string} Type
 *
 * @property {string} FlowNameId
 * @property {string} ServiceKind
 * @property {WadlMetadata} WadlMetadata
 * @property {function(): Object} getWadlMetadata
 */

/**
 * @typedef {Object} SPOListDataSource
 * @property {string} ApiId
 * @property {string} Name
 * @property {string} Type
 *
 * @property {boolean} IsWritable
 * @property {string} DatasetName
 * @property {string} TableName
 * @property {Object.<string,string>} DataEntityMetadataJson - key : TableName
 * @property {Object.<string,string} ConnectedDataSourceInfoNameMapping - key : SPO Column Name, Value : DataSource Column Name
 * @property {function(): Object} getDataEntityMetadata - DataEntityMetadataJson の中身を取得する
 */

/**
 * @typedef {Object} DataSourcesJson
 * @property {DataSourceBase[]} DataSources
 */

const DATASOURCE_API_ID__FLOW =
  "/providers/microsoft.powerapps/apis/shared_logicflows";
const DATASOURCE_API_ID__SPO_LIST =
  "/providers/microsoft.powerapps/apis/shared_sharepointonline";

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

    const manifestKey = keys
      .filter((key) => ("/" + key).endsWith("/manifest.json"))
      .sort((a, b) => {
        return a.length - b.length;
      })[0];
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
   * @returns {Promise<DataSourceBase[]>}
   */
  async getDataSourcesAsync() {
    const msapp = this._msapp;
    if (!msapp) throw Error("This is not loaded.");

    const keys = Object.keys(msapp.files);
    const dataSourcesKey = keys.filter((key) =>
      ("\\" + key).endsWith("\\References\\DataSources.json")
    )[0];

    if (!dataSourcesKey) {
      throw Error("'DataSourses.json' is not found.");
    }

    const dataSourcesJson = await msapp.files[dataSourcesKey].async("string");
    /** @type {DataSourcesJson} */
    const dataSources = JSON.parse(dataSourcesJson);

    return dataSources.DataSources;
  }

  /**
   *
   * @returns {Promise<SPOListDataSource[]>}
   */
  async getSPOListDataSourcesAsync() {
    const dataSources = await this.getDataSourcesAsync();

    /** @type {SPOListDataSource[]} */
    const res = dataSources
      .filter((d) => d.ApiId === DATASOURCE_API_ID__SPO_LIST)
      .map((d) => {
        /** @type {SPOListDataSource} */
        const thisItem = d;
        thisItem.getDataEntityMetadata = () => {
          const j = thisItem.DataEntityMetadataJson[thisItem.TableName];
          return JSON.parse(j);
        };

        return thisItem;
      });
    return res;
  }

  /**
   *
   * @returns {Promise<FlowDataSource[]>}
   */
  async getFlowDataSourcesAsync() {
    const dataSources = await this.getDataSourcesAsync();

    /** @type {FlowDataSource[]} */
    const res = dataSources
      .filter((d) => d.ApiId === DATASOURCE_API_ID__FLOW)
      .map((d) => {
        /** @type {FlowDataSource} */
        const thisItem = d;
        thisItem.getWadlMetadata = () => {
          const xml = thisItem.WadlMetadata.WadlXml;
          const parser = new DOMParser();
          return parser.parseFromString(xml, "text/xml");
        };

        return thisItem;
      });
    return res;
  }

  /**
   *
   * @returns {Promise<DataSourceBase[]>}
   */
  async getOtherDataSourcesAsync() {
    const dataSources = await this.getDataSourcesAsync();

    const res = dataSources.filter(
      (d) =>
        d.ApiId !== DATASOURCE_API_ID__FLOW &&
        d.ApiId !== DATASOURCE_API_ID__SPO_LIST
    );

    return res;
  }

  /**
   *
   * @returns {Blob}
   */
  async createBlobAsync() {
    const msapp = this._msapp;
    if (!msapp) throw Error("This is not loaded.");

    const compBlob = await this._zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });

    return compBlob;
  }
}
