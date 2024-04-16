/**
 * @typedef {Object} SettingData
 * @property {string} tenantId - テナントID
 */

export class Settings {
  constructor(settingsUrl) {
    const url = settingsUrl ? settingsUrl : "./assets/settings.json";

    /** @type {string} */
    this._url = url;
    /** @type {SettingData} */
    this.data = null;
  }

  /**
   *
   * @returns {Promise<SettingData>}
   */
  async loadAsync() {
    const data = await fetch(this._url, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
      },
    }).then((r) => r.json());
    this.data = data;
    return data;
  }
}
