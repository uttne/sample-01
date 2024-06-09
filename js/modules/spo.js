const EntityTypeNames = Object.freeze({
  Shared_x0020_Documents: "Shared_x0020_Documents",
  OData__x005f_catalogs_x002f_appdata: "OData__x005f_catalogs_x002f_appdata",
  OData__x005f_catalogs_x002f_appfiles: "OData__x005f_catalogs_x002f_appfiles",
  TaxonomyHiddenListList: "TaxonomyHiddenListList",
  OData__x005f_catalogs_x002f_wte: "OData__x005f_catalogs_x002f_wte",
  OData__x005f_catalogs_x002f_wp: "OData__x005f_catalogs_x002f_wp",
  SitePages: "SitePages",
  SiteAssets: "SiteAssets",
  Style_x0020_Library: "Style_x0020_Library",
  OData__x005f_catalogs_x002f_solutions:
    "OData__x005f_catalogs_x002f_solutions",
  OData__x005f_catalogs_x002f_theme: "OData__x005f_catalogs_x002f_theme",
  ListList: "ListList",
  Shared_x0020_Documents: "Shared_x0020_Documents",
  FormServerTemplates: "FormServerTemplates",
  OData__x005f_catalogs_x002f_masterpage:
    "OData__x005f_catalogs_x002f_masterpage",
  UserInfo: "UserInfo",
  OData__x005f_catalogs_x002f_lt: "OData__x005f_catalogs_x002f_lt",
  OData__x005f_catalogs_x002f_design: "OData__x005f_catalogs_x002f_design",
  IWConvertedForms: "IWConvertedForms",
});

/** フィールドの型一覧 */
export const FieldTypeKinds = Object.freeze({
  Text: 2, // テキストフィールド
  Memo: 3, // メモフィールド（複数行テキスト）
  Number: 4, // 数値フィールド
  DateTime: 6, // 日付と時刻フィールド
  Lookup: 7, // 参照フィールド（ルックアップ）
  Boolean: 8, // ブール値フィールド
  Select: 9, // 選択フィールド（選択肢）
});

/**
 * @typedef {Object} Deferred
 * @property {string} uri
 */

/**
 * @typedef {Object} RefField
 * @property {Deferred} __deferred
 */

/**
 * @typedef {Object} Metadata
 * @property {string} id
 * @property {string} type
 * @property {string} uri
 */

/**
 * @typedef {Object} ListObject リストのオブジェクト
 * @property {number} BaseTemplate
 * @property {number} BaseType
 * @property {string} Created
 * @property {string} EntityTypeName
 * @property {RefField} Fields
 * @property {boolean} Hidden
 * @property {string} Id
 * @property {RefField} Items
 * @property {string} LastItemDeletedDate
 * @property {string} LastItemModifiedDate
 * @property {string} LastItemUserModifiedDate
 * @property {RefField} RoleAssignments
 * @property {string} Title
 * @property {RefField} Views
 * @property {Metadata} __metadata
 */

/**
 * @typedef {Object} ListFieldObject リストのフィールドのオブジェクト
 * @property {boolean} AutoIndexed
 * @property {boolean} CanBeDeleted
 * @property {string} Description
 * @property {string} EntityPropertyName
 * @property {number} FieldTypeKind
 * @property {string} TypeAsString
 * @property {number} Filterable
 * @property {boolean} Hidden
 * @property {string} Id
 * @property {boolean} ReadOnlyField
 * @property {string} SchemaXml
 * @property {string} StaticName
 * @property {string} Title
 * @property {string} InternalName
 * @property {boolean} Required
 * @property {boolean} EnforceUniqueValues
 * @property {Metadata} __metadata
 */

/**
 * @typedef {Object} ListViewObject リストのビューのオブジェクト
 * @property {string} Id
 * @property {string} Title
 * @property {RefField} ViewFields
 * @property {Metadata} __metadata
 */

/** 404 の例外が発生したときのエラー */
export class NotFoundError extends Error {
  constructor(message) {
    super(message);
  }
}

/** すでに存在するエラー */
export class ExistedError extends Error {
  constructor(message) {
    super(message);
  }
}

/** SharePoint のリストを操作するためのクラス
 */
export class SpoClient {
  /**
   *
   * @param {string} siteUrl サイトのURL
   */
  constructor(siteUrl) {
    this._siteUrl = siteUrl.replace(/\/+$/g, "");
    this._fqdn = /^(http|https):\/\/(.+?)(\/|$)/.exec(this._siteUrl)[2];

    /** @type {string|undefined} */
    this._fdv = undefined;
  }

  /**
   * 視覚情報の取得
   * @param {boolean} force
   * @returns {string}
   */
  async getFdvAsync(force = undefined) {
    if (!force) {
      if (!this._fdv) {
        return this._fdv;
      }
    }
    const url = `${this._siteUrl}/_api/contextinfo`;
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
    this._fdv = fdv;
    return fdv;
  }

  /**
   * リストを取得する
   * @param {string} title
   * @returns {ListObject}
   */
  async getListByTitle(title) {
    const url = `${this._siteUrl}/_api/web/lists/GetByTitle('${title}')`;
    const fdv = await this.getFdvAsync();
    const headers = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": fdv,
    };

    const res = await fetch(url, { method: "GET", headers: headers });
    const j = await res.json();
    if (!res.ok) {
      console.log(j);
      if (res.status === 404) {
        throw new NotFoundError(`リストは存在しませんでした : ${title}`);
      }
      throw new Error(`リストの取得を失敗しました : ${title}`);
    }

    return j.d;
  }

  /**
   * リストを取得する
   * @param {string} id
   * @returns {ListObject}
   */
  async getListById(id) {
    const url = `${this._siteUrl}/_api/web/lists(guid'${id}')`;
    const fdv = await this.getFdvAsync();
    const headers = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": fdv,
    };

    const res = await fetch(url, { method: "GET", headers: headers });
    const j = await res.json();

    if (!res.ok) {
      console.log(j);
      if (res.status === 404) {
        throw new NotFoundError(`リストは存在しませんでした : ${id}`);
      }
      throw new Error(`リストの取得を失敗しました : ${id}`);
    }

    return j.d;
  }

  /**
   * リストのフィールドを取得
   * @param {ListObject} listObj
   * @returns {ListFieldObject[]}
   */
  async getListFields(listObj) {
    const url = listObj.Fields.__deferred.uri;
    const fdv = await this.getFdvAsync();
    const headers = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": fdv,
    };

    const res = await fetch(url, { method: "GET", headers: headers });
    const j = await res.json();

    if (!res.ok) {
      console.log(j);
      if (res.status === 404) {
        throw new NotFoundError(`リストのフィールドは存在しませんでした`);
      }
      throw new Error(`リストのフィールドの取得を失敗しました`);
    }

    return j.d.results;
  }

  /**
   * リストのビューを取得
   * @param {ListObject} listObj
   * @returns {ListViewObject[]}
   */
  async getListViews(listObj) {
    const url = listObj.Views.__deferred.uri;
    const fdv = await this.getFdvAsync();
    const headers = {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": fdv,
    };

    const res = await fetch(url, { method: "GET", headers: headers });
    const j = await res.json();

    if (!res.ok) {
      console.log(j);
      if (res.status === 404) {
        throw new NotFoundError(`リストのビューは存在しませんでした`);
      }
      throw new Error(`リストのビューの取得を失敗しました`);
    }

    return j.d.results;
  }

  /**
   * リストを作成
   * @param {string} title
   * @param {string} description
   * @returns {ListObject}
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
        // 100: カスタムリスト
        // 101: ドキュメントライブラリ
        // 102: リンクリスト
        // 103: カレンダーリスト
        // 104: タスクリスト
        // 105: 連絡先リスト
        BaseTemplate: 100, // GUI で作るリスト
        ContentTypesEnabled: false, // GUI で作成するデフォルトは false のよう
        Description: description,
        Title: title,
      }),
    });
    const j = await res.json();

    if (!res.ok) {
      console.log(j);
      if (res.error.code === "-2130575342, Microsoft.SharePoint.SPException") {
        throw new ExistedError(j.error.message);
      }
      throw new Error("リストの作成失敗");
    }

    return j.d;
  }

  /**
   * リストのフィールドを追加する
   * fieldObj.Title, FieldTypeKind は必須
   * ※ 同じ InternalName を指定した場合、連番が付与されて新しいフィールドが作成される
   * @param {ListObject} listObj
   * @param {ListFieldObject} fieldObj
   * @returns {ListFieldObject}
   */
  async addListFieldAsync(listObj, fieldObj) {
    const fdv = await this.getFdvAsync();

    const url = listObj.Fields.__deferred.uri;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": fdv,
      },
      body: JSON.stringify(
        Object.assign({}, { __metadata: { type: "SP.Field" } }, fieldObj)
      ),
    });
    const j = await res.json();

    if (!res.ok) {
      console.log(j);
      throw new Error("フィールドの追加失敗");
    }

    return j.d;
  }

  /**
   * リストのフィールドを変更する
   * @param {ListFieldObject} targetFieldObj
   * @param {ListFieldObject} newFieldObj
   */
  async patchListFieldAsync(targetFieldObj, newFieldObj) {
    const fdv = await this.getFdvAsync();

    const url = targetFieldObj.__metadata.uri;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": fdv,
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify(
        Object.assign({}, { __metadata: { type: "SP.Field" } }, newFieldObj)
      ),
    });

    if (!res.ok) {
      const j = await res.json();
      console.log(j);
      throw new Error("フィールドの更新失敗");
    }
  }

  /**
   * リストのフィールドを変更する
   * @param {ListViewObject} viewObj
   * @param {ListFieldObject} fieldObj
   */
  async addFieldToViewAsync(viewObj, fieldObj) {
    const fdv = await this.getFdvAsync();

    const url =
      viewObj.ViewFields.__deferred.uri +
      `/AddViewField('${fieldObj.EntityPropertyName}')`;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": fdv,
      },
      body: JSON.stringify(
        Object.assign({}, { __metadata: { type: "SP.Field" } }, newFieldObj)
      ),
    });

    if (!res.ok) {
      const j = await res.json();
      console.log(j);
      if (j.error.code === "-2147024809, System.ArgumentException") {
        throw new NotFoundError(j.error.message);
      }
      throw new Error("フィールドの更新失敗");
    }
  }
}
