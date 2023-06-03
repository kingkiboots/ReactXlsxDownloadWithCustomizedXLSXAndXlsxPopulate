/**
 *
 * @param {string} byte byte로 변환할 문자열
 * @param {boolean} koreanAs2Byte true: 한글을 2바이트로 반환, default(false, undefined): 한글을 3바이트로 반환
 * @returns 문자열의 바이트사이즈
 */
const byteSize = (str, koreanAs2Byte) => (koreanAs2Byte ? byteSizeOfEUCKR(str) : new Blob([str]).size);
const byteSizeOfEUCKR = function (s, b, i, c) {
  for (b = i = 0; (c = String(s).charCodeAt(i++)); b += c >> 11 ? 2 : c >> 7 ? 2 : 1);
  return b;
};

export { byteSize };
