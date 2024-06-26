﻿<!DOCTYPE html>
<html lang="ja">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>sample</title>
    <link href="./css/libs/bootstrap.min.css" rel="stylesheet" />
  </head>
  <body>
    <div class="container text-center">
      <div class="row">
        <div class="col">
          <div class="input-group mb-3">
            <label class="input-group-text" for="zipInput"
              >Zipファイルを選択してください</label
            >
            <input
              type="file"
              class="form-control"
              id="zipInput"
              accept=".zip"
            />
          </div>
        </div>
      </div>
      <div class="row">
        <div class="col">
          <select class="custom-select" id="comboBox">
            <option selected>選択してください...</option>
          </select>
        </div>
      </div>
      <div class="row">
        <div class="col">
          <div class="mt-3" id="textData"></div>
        </div>
      </div>
      <div class="row">
        <div class="col">
          <table class="table" id="dataSourceTable">
            <thead>
              <tr>
                <th scope="col">#</th>
                <th scope="col">データソースタイプ</th>
                <th scope="col">名称</th>
                <th scope="col">リンク</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <th scope="row">1</th>
                <td>SharePoint Online リスト</td>
                <td>リストの名称</td>
                <td><a href="https://example.com">https://example.com</a></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
      <div class="row">
        <div class="col">
          <button type="button" class="btn btn-primary" id="download-button">
            編集したZIPパッケージをダウンロード
          </button>
        </div>
      </div>

      <div class="row">
        <div class="col">
          <div class="input-group mb-3">
            <span class="input-group-text" id="site-url">サイトURL</span>
            <input
              id="site-url-input"
              type="text"
              class="form-control"
              placeholder="SiteUrl"
              aria-label="SiteUrl"
              aria-describedby="site-url"
            />
          </div>
          <div class="input-group mb-3">
            <span class="input-group-text" id="new-table-title"
              >新しいリストの名前</span
            >
            <input
              id="new-table-title-input"
              type="text"
              class="form-control"
              placeholder="NewTableTitle"
              aria-label="NewTableTitle"
              aria-describedby="new-table-title"
            />
          </div>
          <div class="input-group mb-3">
            <button type="button" class="btn btn-primary" id="new-table-button">
              新しいリストを生成
            </button>
          </div>
        </div>
      </div>
    </div>
    <script src="./js/libs/bootstrap.bundle.min.js"></script>
    <script src="./js/libs/jszip.js"></script>
    <script type="module" src="./js/index.js"></script>
  </body>
</html>
