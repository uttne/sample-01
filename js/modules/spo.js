export class SpoClient {
  /**
   *
   * @param {string} siteUrl
   */
  constructor(siteUrl) {
    this._siteUrl = siteUrl.replace(/\/+$/g, "");
  }

  async getFdvAsync() {
    const url = this._siteUrl + "/_api/contextinfo";
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
      },
    });

    if (!res.ok) {
      throw new Error("存在しないサイトです");
    }
    const contextinfo = await res.json();
    let fdv = contextinfo.FormDigestValue;
    return fdv;
  }

  /**
   *
   * @param {string} title
   * @param {string} description
   */
  async createListAsync(title, description) {
    const fdv = await this.getFdvAsync();

    const url = this._siteUrl + "/_api/web/lists";
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": fdv,
      },
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        AllowContentTypes: true,
        BaseTemplate: 100,
        ContentTypesEnabled: false, // GUI で作成するデフォルトは false のよう
        Description: description,
        Title: title,
      }),
    });

    if (!res.ok) {
      throw new Error("リストの作成失敗");
    }
  }
}
