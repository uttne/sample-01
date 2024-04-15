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

/**
 * @typedef {Object} DataSourceRow
 * @property {number} no
 * @property {string} dataType
 * @property {string} name
 * @property {string} link
 */

class DataSourceTable {
  constructor() {
    this._elm = document.getElementById("dataSourceTable");
    /** @type {DataSourceRow[]} */
    this._rows = [
      {
        no: 1,
        dataType: "sample",
        name: "sample",
        link: "https://example.com",
      },
    ];
  }

  /**
   *
   * @param {string} dataType
   * @param {string} name
   * @param {string} link
   */
  add(dataType, name, link) {
    const lastRow = this._rows.slice(-1)[0];
    const no = lastRow ? lastRow.no + 1 : 1;

    this._rows.push({
      no: no,
      dataType: dataType,
      name: name,
      link: link,
    });
  }

  clear() {
    this._rows = [];
  }

  refresh() {
    const tbody = this._elm.querySelector("tbody");
    tbody.innerHTML = "";

    this._rows.forEach((row) => {
      const tr = document.createElement("tr");

      const thNo = document.createElement("th");
      thNo.setAttribute("scope", "row");
      thNo.innerText = row.no.toString();
      tr.appendChild(thNo);

      const tdType = document.createElement("td");
      tdType.innerText = row.dataType;
      tr.appendChild(tdType);

      const tdName = document.createElement("td");
      tdName.innerText = row.name;
      tr.appendChild(tdName);

      const tdLink = document.createElement("td");
      const a = document.createElement("a");
      a.setAttribute("href", row.link);
      a.setAttribute("target", "_blank");
      a.innerText = row.link;

      tdLink.appendChild(a);
      tr.appendChild(tdLink);

      tbody.appendChild(tr);
    });
  }
}

window.selectElm = new SelectElement();
window.tableElm = new DataSourceTable();
window.powerApps = new PowerApps();

function main() {
  const zipInput = document.getElementById("zipInput");
  zipInput.addEventListener("change", (e) => {
    /** @type {File} */
    const file = e.target.files[0];

    if (!file) return;

    /** @type {PowerApps} */
    const powerApps = window.powerApps;
    powerApps.loadAsync(file).then(async () => {
      /** @type {DataSourceTable} */
      const tableElm = window.tableElm;
      tableElm.clear();

      const spoListDataSources = await powerApps.getSPOListDataSourcesAsync();
      const flowDataSources = await powerApps.getFlowDataSourcesAsync();
      const otherDataSources = await powerApps.getOtherDataSourcesAsync();

      spoListDataSources.forEach((ds) => {
        tableElm.add(
          "SharePoint Online リスト",
          ds.Name,
          // とりあえず設定ページ
          `${ds.DatasetName}/_layouts/15/listedit.aspx?List=${ds.TableName}`
        );
      });

      flowDataSources.forEach((ds) => {
        tableElm.add(
          "Power Automate",
          ds.Name,
          // FlowNameId を使ってURLを構築
          `${ds.FlowNameId}`
        );
      });

      otherDataSources.forEach((ds) => {
        tableElm.add("その他", ds.Name, "");
      });

      tableElm.refresh();
    });
  });
}

main();
