import { no } from 'util/ComponentUtil';
import { handleExport, json_to_book, table_to_book } from './CustomizedXlsx'; //eslint-disable-line no-unused-vars

const PageOrWhole = {
  page: 'page',
  whole: 'whole'
};
/**
 * 현재 보이는 테이블 정보를(현재 페이지) 엑셀파일로 추출하는 메서드
 * @param tableRef xlsx파일로 추출할 테이블의 useRef() 값
 * @param layoutHeaderName 파일명 => TableLayout의 layoutHeaderName 프롭에 주시는 값을 주시면 됩니다.
 *
 * {1} : 테이블ref를 넘겨서 그 안에 있는 cell Data들을 sheetObjt 로 반환
 * {2} : 엑셀시트 오브젝트를 스타일 먹이고 파일로 다운로드
 */
const table2XLSX = async (tableRef, layoutHeaderName) => {
  /* {1} */
  const sheetObj = table_to_book(layoutHeaderName, tableRef.current);
  /* {2} */
  executeHandleExport(sheetObj, PageOrWhole.page, layoutHeaderName);
};

/**
 * 서버에 전체 데이터를 요청하여 그 응답을 해당 테이블 형식에 맞게 해서 엑셀 파일로 추출하는 메서드
 * @param rowDef 컬럼의 속성 정보가 담긴 rowDef
 * @param list 데이터가 담긴 리스트
 * @param layoutHeaderName 파일명 => TableLayout의 layoutHeaderName 프롭에 주시는 값을 주시면 됩니다.
 * @param noTitleHeader 엑셀시트의 1,2 번 로우에 제목이 적히는 지 않적히는 지 여부
 *
 * {1} : 만약 헤더가 두 줄로 되어 있는 경우 두개의 row로 만들기 위한 함수
 * {2} : rowdef의 순서에 맞게 json데이터의 값들도 정렬
 *          => 전체가 배열, 한 row 그 전체 배열의 배열, 내부 배열의 값들은 엑셀 셀에 들어갈 데이터들
 * {3} : 엑셀파일의 헤더가 될 라벨네임과 colspan, rowspan을 추출
 * {4} : 헤더가 맨 윗 줄이 되고, 그 밑에 헤더 name에 의해서 정렬된({2} 참고) 배열도 추가
 * {5} : 준비된 데이터를 이제 엑셀파일 워크북으로 변환
 * {6} : 데이터의 길이가 동일한 컬럼의 인덱스들, 데이터가 1,000% 같이 숫자인 컬럼의 인덱스들을 구합니다.
 * {7} : 헤더 부분만 특별히 스타일링 해주기 위해서 헤더의 row 길이를 여기서 구했습니다.
 * {8} : 스타일링을 위해서 sheetObj의 headerRowLen에 {6}번 값 삽입
 * {9} : 엑셀시트 오브젝트를 스타일 먹이고 파일로 다운로드 하기
 */
const json2XLSX = (rowDef, list, layoutHeaderName, noTitleHeader) => {
  /* {1} */
  const headerVal = seperateRowDef(rowDef);
  console.log(`headerVal => ${JSON.stringify(headerVal)}`);
  /* {2} */
  const tdDataValues = extractOnlyValues(list, rowDef);
  console.log(`tdDataValues => ${JSON.stringify(tdDataValues)}`);
  /* {3} */
  const headerLabels = extractLabelNameFromHeaderVal(headerVal);
  console.log(`headerLabels => ${JSON.stringify(headerLabels)}`);
  /* {4} */
  const wholeData = headerLabels.concat(tdDataValues);
  console.log(`wholeData => ${JSON.stringify(wholeData)}`);
  /* {5} */
  const sheetObj = json_to_book(layoutHeaderName, wholeData, noTitleHeader);
  console.log('sheetObj', sheetObj);
  console.log('sheetObj.ws', sheetObj.ws);
  /* {6} */
  const { equalLengthColumnIndexList, numberFormatIndexList } = getColumnWidthInfo(tdDataValues);
  /* {7} */
  const headerRowLen = headerVal.length;
  /* {8} */
  sheetObj['headerRowLen'] = headerRowLen;
  sheetObj['equalLengthColumnIndexList'] = equalLengthColumnIndexList;
  sheetObj['numberFormatIndexList'] = numberFormatIndexList;

  /* {9} */
  executeHandleExport(sheetObj, PageOrWhole.whole, layoutHeaderName, noTitleHeader);
};

/**
 * 헤더가 두 줄 이상이 될 수도 있다는 점을 고려하여 두 줄 이상일 시 한개의 배열을 여러개로 찢는 메서드
 * @param rowDef 헤더 정보
 * @returns 두 줄 이상일 시 한개의 배열을 여러개로 찢는 메서드
 */
const seperateRowDef = (rowDef) => {
  const headerVals = [];
  let currentArray = rowDef;

  while (currentArray.length > 0) {
    // 맨위의 레벨 headerVals에 다 집어 넣기 => 그 다음에 업데이트 된 다음 세대 애들을 headerVals에 다 집어 넣기
    headerVals.push(currentArray.map((item) => item ?? ''));

    // nextLevelArray에 이것의 자식들을 다 집어넣어 currentArray를 업데이트 시킴
    const nextLevelArray = [];
    currentArray.forEach((item) => {
      if (item.children) {
        nextLevelArray.push(...item.children);
      }
    });
    // 없다면 안 집어넣음
    currentArray = nextLevelArray;
  }
  return headerVals;
};

/**
 * 화면에 보여진 값들만 추출된 객체를 headerName의 순서에 맞게 정렬시킨 후 값들만 배열로 반환한다.
 * @param list 헤더에 등록된 것만 구별된 json 데이터
 * @param rowDef 헤더 정보
 * @returns 순서에 맞게 정렬된 값들만 추출 후 배열로 반환
 */
const extractOnlyValues = (list, rowDef) => {
  const fittedData = arrangeValues(list, rowDef);
  const tdDataValues = extractValuesToArray(fittedData);
  return tdDataValues;
};

/**
 * object 내에서 rowDef와는 상관 없이 뒤죽박죽이 되어있을 key: value 들을 rowDef에 정의된 순서대로 정렬하고 포맷 등 rowDef의 다양한 옵션대로 value작성한다.
 * @param list 헤더에 등록된 것만 구별된 json 데이터
 * @param rowDef 헤더 정보
 * @returns
 *
 * 결과적으로
 * [{undefined: 1, menuGrpNm: '해피페이머니', menuId: 'Money', menuNm: '해피페이머니', menuUrl: '/', …},
 * {undefined: 2, menuGrpNm: '해피페이머니', menuId: 'MemberInfoMgmt', menuNm: '회원정보관리', menuUrl: '/memberInfoMgmt', …}]
 * 이런 형태로 정렬됨
 */
const arrangeValues = (list, rowDef) => {
  let i = 1;
  const fittedData = list.map((row) => {
    const { newRow, newI } = getNewRowsAndNewI(rowDef, row, i);
    i = newI;
    return newRow;
  });
  return fittedData;
};

/**
 * rowDef에는 type: no일 시 rownum을 작성하는 것, value의 포매팅(format, dynamicFormat) 혹은
 * 다른 어떤 값에 따라 value가 어떻게 되는 것(valueCondition)등에 대한 옵션이 있는데 이 옵션들을 적용하여 엑셀시트에 작성될 데이터들을 작성한다.
 * @param {Array} rowDef 헤더 정보
 * @param {object} row 여러줄의 데이터(list) 중 하나의 row 즉 하나의 Json string
 * @param {number} i rows의 인덱스 번호
 * @returns
 */
const getNewRowsAndNewI = (rowDef, row, i) => {
  const newRow = {};
  rowDef.forEach((header) => {
    const headerNm = header.name;
    if (headerNm !== undefined) {
      // if (header.format) newRow[headerNm] = header.format(row[headerNm] ?? '');
      // else if (header.valueCondition) newRow[headerNm] = header.valueCondition(row);
      // else newRow[headerNm] = row[headerNm];
      let targetValue;
      if (header.valueCondition) targetValue = header.valueCondition(row);
      else targetValue = row[headerNm] ?? '';
      if (header.format) targetValue = header.format(targetValue ?? '');
      else if (header.dynamicFormat) targetValue = header.dynamicFormat(targetValue ?? '');
      newRow[headerNm] = targetValue;
    } else if (header.type == no) {
      // 헤더의 타입이 No.라면 직접 번호 값을 주는 것으로
      const isNoHide = header.noHideCondition ? header.noHideCondition(row) : false;
      if (isNoHide) {
        newRow[headerNm] = '';
      } else newRow[headerNm] = i++;
    }
    if (header.children) {
      // const rowOfChild = getNewRowsAndNewI(header.children, row, i);
      const { newRow: rowOfChild } = getNewRowsAndNewI(header.children, row, i);
      Object.assign(newRow, rowOfChild);
    }
  });
  return { newRow: newRow, newI: i };
};

/**
 * 순서에 맞게 정렬 및 작성된 객체에서 그 값들만 추출 후 배열로 반환
 * @param fittedData 순서에 맞게 정렬된 객체 배열
 * @returns 순서에 맞게 정렬된 객체에서 그 값들만 추출 된 배열
 */
const extractValuesToArray = (fittedData) => {
  const tdDataValues = fittedData.map((item) => {
    return Object.entries(item).map((val) => {
      return val[1] ?? '';
    });
  });
  return tdDataValues;
};

/**
 * 엑셀파일의 헤더가 될 라벨네임과 colspan, rowspan을 추출
 * @param headerVal
 * @returns 라벨네임과 colspan, rowspan을 추출된 객체 배열
 */
const extractLabelNameFromHeaderVal = (headerVal) => {
  const headerLabels = headerVal.map((item) => {
    const eachLabelArr = item.map((innerItem) => {
      const labelName = innerItem.labelName;
      if (innerItem.rowSpan || innerItem.colSpan) {
        let cellInfo = {};
        if (innerItem.rowSpan && innerItem.rowSpan > 1) {
          cellInfo['rowspan'] = innerItem.rowSpan;
        }
        if (innerItem.colSpan && innerItem.colSpan > 1) {
          cellInfo['colspan'] = innerItem.colSpan;
        }
        cellInfo['value'] = labelName;
        return cellInfo;
      }
      return labelName;
    });
    return eachLabelArr;
  });
  return headerLabels;
};

/**
 * 엑셀시트에 스타일을 적용할 시 숫자 포맷인 컬럼, 문자열인데 모든 데이터의 길이가 동일한 컬럼은 스타일링을 별도로 해주기 위해 해당 컬럼들의 인덱스를 저장 반환한다.
 * @param {바디 정보} tdDataValues
 * @returns {equalLengthColumnIndexList, numberFormatIndexList} 데이터의 길이가 동일한 컬럼의 인덱스들, 데이터가 1,000% 같이 숫자인 컬럼의 인덱스들
 */
const getColumnWidthInfo = (tdDataValues) => {
  const arrLength = tdDataValues[0].length;
  let isDatasLengthOfEachColumnNotEqualList = new Array(arrLength);
  let columnDataLengthList = new Array(arrLength);
  let isNotNumberFormatList = new Array(arrLength);
  tdDataValues.forEach((row, rowIdx) => {
    row.forEach((item, idx) => {
      const itemLength = item.length;
      if (!['합계', '소계'].includes(item) && itemLength > 0) {
        if (!isDatasLengthOfEachColumnNotEqualList[idx]) {
          if (columnDataLengthList[idx] && columnDataLengthList[idx] !== itemLength && rowIdx > 0) {
            isDatasLengthOfEachColumnNotEqualList[idx] = true;
          } else {
            isDatasLengthOfEachColumnNotEqualList[idx] = false;
          }
          columnDataLengthList[idx] = itemLength;
        }
        if (!isNotNumberFormatList[idx] && rowIdx > 0) {
          const regex1 = /^(\D)*-?(\d{1,3},)*(\d{1,3})(\.\d*[1-9]+)?$/g; // 단위가 앞에 있는 경우
          const regex2 = /^-?(\d{1,3},)*(\d{1,3})(\.\d*[1-9]+)?(\D)*$/g; // 단위가 뒤에 있는 경우
          if (regex1.test(item) || regex2.test(item)) {
            isNotNumberFormatList[idx] = false;
          } else {
            isNotNumberFormatList[idx] = true;
          }
        }
      }
    });
  });
  const equalLengthColumnIndexList = isDatasLengthOfEachColumnNotEqualList.map((e, idx) => (e === false && e !== undefined ? idx : undefined)).filter((e) => e !== undefined);
  const numberFormatIndexList = isNotNumberFormatList.map((e, idx) => (e === false && e !== undefined ? idx : undefined)).filter((e) => e !== undefined);

  return { equalLengthColumnIndexList, numberFormatIndexList };
};

/**
 * 엑셀시트 오브젝트를 스타일 먹이고 파일로 다운로드 하는 메서드
 * @param sheetObj
 * @param pageOrWhole 현재 페이지인지 아니면 전체 데이터인지
 * @param layoutHeaderName 파일 제목
 */
const executeHandleExport = (sheetObj, pageOrWhole, layoutHeaderName, noTitleHeader) => {
  const fileNm = setFileNameByItsPurpose(pageOrWhole, layoutHeaderName);
  const { ws: workBook } = sheetObj;
  delete sheetObj.ws;
  handleExport(workBook, sheetObj, noTitleHeader).then((url) => {
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute('href', url);
    downloadAnchorNode.setAttribute('download', `${fileNm}.xlsx`);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  });
};

/**
 * 현재 페이지인지 아니면 전체 데이터인지 구분하여 _전체 라는 것을 추가해준다.
 * @param pageOrWhole 현재 페이지인지 아니면 전체 데이터인지
 * @param layoutHeaderName 파일 제목
 * @returns 파일 이름 + 년도날짜시간 합하여 파일명을 반환
 */
const setFileNameByItsPurpose = (pageOrWhole, layoutHeaderName) => {
  const today = getFormatDate(new Date());
  switch (pageOrWhole) {
    case PageOrWhole.page:
      return `${layoutHeaderName}_${today}.xlsx`;
    case PageOrWhole.whole:
      return `${layoutHeaderName}.xlsx`;
    // case PageOrWhole.whole:
    //   return `${layoutHeaderName}_전체_${today}.xlsx`;
    default:
      return;
  }
};

/**
 * 지금 날짜, 시간 을 계산하여 반환하는 메서드
 * @param date new Date()
 * @returns 20230309+시간
 */
const getFormatDate = (date) => {
  var year = date.getFullYear(); //yyyy
  var month = 1 + date.getMonth(); //M
  month = month >= 10 ? month : '0' + month; //month 두자리로 저장
  var day = date.getDate(); //d
  day = day >= 10 ? day : '0' + day; //day 두자리로 저장
  var hour = date.getHours(); //hh
  hour = hour >= 10 ? hour : '0' + hour; //hh 두자리로 저장
  var minites = date.getMinutes(); //mm
  minites = minites >= 10 ? minites : '0' + minites; //minute 두자리로 저장
  var seconds = date.getSeconds(); //ss
  seconds = seconds >= 10 ? seconds : '0' + seconds; //second 두자리로 저장
  return year + month + day + hour + minites + seconds; //'-' 추가하여 yyyy-mm-dd 형태 생성 가능
};

export { table2XLSX, json2XLSX };
