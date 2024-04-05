import "./@types/jszip.t.js";
import { PowerApps } from "./modules/power-apps.js";

class SelectElement {
  constructor() {
    this._elm = document.getElementById("comboBox");
    /** @type {Object.<string, string>} */
    this._values = {};

    const thisObj = this;

    this._elm.addEventListener("change", (e) => {
      const value = e.target.value;

      document.getElementById("textData").textContent = thisObj._values[value];
    });
  }

  /**
   *
   * @param {string} name
   * @param {string} content
   */
  addOption(name, content) {
    const option = document.createElement("option");
    option.value = name;
    option.innerText = name;

    this._elm.appendChild(option);
    this._values[name] = content;
  }

  /**
   *
   * @param {string} name
   * @returns {string}
   */
  getValue(name) {
    return this._values[name];
  }
}

window.selectElm = new SelectElement();
window.powerApps = new PowerApps();

function main() {
  const zipInput = document.getElementById("zipInput");
  zipInput.addEventListener("change", (e) => {
    /** @type {File} */
    const file = e.target.files[0];

    if (!file) return;

    window.powerApps.loadAsync(file);

    /** @type {JSZip} */
    const jszip = new JSZip();
    jszip.loadAsync(file).then(
      (zip) => {
        window.__loaded_cache = zip;
        console.log(window.__loaded_cache);

        Object.keys(zip.files).forEach(async (file) => {
          const zipObj = zip.files[file];
          const name = zipObj.name;
          const dir = zipObj.dir;

          if (dir) return;

          const content = await zipObj.async("string");
          window.selectElm.addOption(name, content);
        });
      },
      (err) => {
        console.log(err);
      }
    );
  });
}

main();
