/**
 * @typedef {Object} JSZip
 * @property {function(File|Blob): Promise<JSZip>} loadAsync - 読み込み
 * @property {Object.<string, ZipObject>} files
 * @property {function(GenerateAsyncOptions): Promise<Blob>} generateAsync
 */

/**
 * @typedef {Object} compressionOptions
 * @property {number} level - 9 が最高品質
 */

/**
 * @typedef {Object} GenerateAsyncOptions - https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html
 * @property {string} type - blob : blob 以外を指定することも可能だが考慮が大変なので blob に決め打つ
 * @property {string} compression - DEFLATE
 * @property {compressionOptions} compressionOptions
 */

/**
 * @typedef {Object} ZipObject
 * @property {boolean} dir - dir か否か
 * @property {string} name - パス
 * @property {function(string): Promise<string>} async - 内容の読み込み async("string").then(content=>console.log(content)) のように記述する : "string" | "blob"
 */
