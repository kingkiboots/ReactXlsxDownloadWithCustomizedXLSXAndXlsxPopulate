const numberFormat = (value = '') => {
  if (value === '') return '';
  return value.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') || value;
};

const numberFormatWithSuffixYear = (value = '') => {
  if (value === '') return value;
  return `${value}년`;
};
/**
 * 숫자 콤마 변환 + '명' ex) 123456789 => 123,456,789원
 * @param value 변환할 값
 * @returns 변환된 값 + '원'
 */
const numberFormatWithSuffixPeople = (value = '') => {
  if (value === '') return '';
  return numberFormat(value) + '명';
};

export { numberFormat, numberFormatWithSuffixYear, numberFormatWithSuffixPeople };
