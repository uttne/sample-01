import "./@types/jszip.t.js";
import { PowerApps } from "./modules/power-apps.js";
import { Settings } from "./modules/settings.js";
import { SpoClient } from "./modules/spo.js";

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

class DownloadButton {
  constructor() {
    this._elm = document.getElementById("download-button");

    this._elm.addEventListener("click", async (e) => {
      /** @type {PowerApps} */
      const powerApps = window.powerApps;

      // サンプル: 名称の変更
      const name = await powerApps.getNameAsync();
      console.log(`app name : ${name}`);
      await powerApps.setNameAsync(name + "_編集版");

      // Blob オブジェクトを作成
      const data = await powerApps.createBlobAsync();

      // Blob URL を生成
      const url = window.URL.createObjectURL(data);

      // a エレメントを動的に作成
      const link = document.createElement("a");
      link.href = url;
      link.download = "horizontally.zip"; // ダウンロードするファイルの名前

      // リンクをクリックする
      link.click();

      // リソースを解放
      window.URL.revokeObjectURL(url);
    });
  }
}

class SpoElements {
  constructor() {
    /** @type {HTMLButtonElement} */
    this._buttonElm = document.getElementById("new-table-button");
    /** @type {HTMLInputElement} */
    this._siteUrlElm = document.getElementById("site-url-input");
    /** @type {HTMLInputElement} */
    this._tableTitleElm = document.getElementById("new-table-title-input");

    this._buttonElm.addEventListener("click", async (e) => {
      const siteUrl = this._siteUrlElm.value;
      const tableTitle = this._tableTitleElm.value;

      if (!siteUrl) {
        alert("サイトのURLを入力してください");
        return;
      }

      if (!tableTitle) {
        alert("作成するテーブルのタイトルを入力してください");
        return;
      }

      const spoClient = new SpoClient(siteUrl);

      try {
        spoClient.createListAsync(tableTitle, "");
        alert(`${tableTitle} を作成しました`);
      } catch (e) {
        console.log(e);
        alert(`${tableTitle} の作成に失敗しました : ${e.toString()}`);
      }
    });
  }
}

window.selectElm = new SelectElement();
window.tableElm = new DataSourceTable();
window.downloadButton = new DownloadButton();
window.powerApps = new PowerApps();
window.spoElms = new SpoElements();
window.settings = new Settings();

function main() {
  /** @type {Settings} */
  const settings = window.settings;
  settings.loadAsync().then((s) => console.log(s));

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

      const settingData = settings.data;

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
          `https://make.powerautomate.com/environments/Default-${settingData.tenantId}/flows/${ds.FlowNameId}/details`
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
