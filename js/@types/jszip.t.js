
/**
 * @typedef {Object} JSZip
 * @property {function(File|Blob): Promise<JSZip>} loadAsync - 読み込み
 * @property {Object.<string, ZipObject>} files
 */

/**
 * @typedef {Object} ZipObject
 * @property {boolean} dir - dir か否か
 * @property {string} name - パス
 * @property {function(string): Promise<string>} async - 内容の読み込み async("string").then(content=>console.log(content)) のように記述する : "string" | "blob"
 */
