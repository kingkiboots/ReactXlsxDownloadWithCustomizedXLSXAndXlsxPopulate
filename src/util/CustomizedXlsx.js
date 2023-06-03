/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*exported XLSX */

/* SheetJS 라이브러리의 소스코드를 수정 보완한 것입니다. */
/* 아직 못한 것들 : td가 데이터일 경우에는 엑셀에 텍스트가 안나오게 하는 것, 두개의 비슷한 메서드를 하나로 합치는 리팩토링 */

import XLSX from 'sheetjs-style';
import { byteSize } from 'util/DataUtil';
import * as XlsxPopulate from 'xlsx-populate/browser/xlsx-populate';

var DENSE = null;

/**
 * 현재 보이는 테이블 정보를(현재 페이지) 엑셀파일 워크북 객체로 반환하는 메서드
 * @param title layoutHeaderName
 * @param table tableRef.current
 * @param opts XlsxPopulate 모듈을 사용하여 시트에 스타일을 먹이기 위해서는 배열이 아닌 객체에 정보를 저장해야하므로 여기서는 쓸일이 없긴 하지만 소스코드에는 있어서 놔두고 있었습니다.
 * @returns 데이터들이 담긴 워크북 객체
 */
function table_to_book(title /*:?string*/, table /*:HTMLElement*/, opts /*:?any*/) /*:Workbook*/ {
  //eslint-disable-line no-unused-vars
  let sheetObj = parse_dom_table(title, table, opts);
  return returnWorkbook(sheetObj, opts);
}

/**
 * 테이블의 전체 데이터(JSON형식)를 (전체 페이지) 엑셀파일 워크북 객체로 반환하는 메서드
 * @param title layoutHeaderName
 * @param wholeData
 * @param opts 맨위의 설명과 같습니다.
 * @param noTitleHeader 시트 내에 맨 위의 헤더타이틀이 안뜨게 하고 싶으실 경우 true
 * @returns
 */
function json_to_book(title /*:?string*/, wholeData /*:HTMLElement*/, noTitleHeader /*:? boolean */, opts /*:?any*/) /*:Workbook*/ {
  //eslint-disable-line no-unused-vars
  let sheetObj = parse_json(title, wholeData, noTitleHeader, opts);
  return returnWorkbook(sheetObj, opts);
}

/**
 * 워크싯 데이터를 워크북으로 반환하는 메서드
 * @param {*} sheetObj
 * @param {*} opts 맨위의 설명과 같습니다.
 * @returns 워크북 객체
 */
function returnWorkbook(sheetObj, opts) {
  sheetObj['ws'] = sheet_to_workbook(sheetObj.ws, opts);
  return sheetObj;
}

/**
 * 엑셀 시트 이름과 엑셀시트에 들어갈 데이터를 집어넣는 메서드
 * @param sheet sheetObj
 * @param opts 맨위의 설명과 같습니다.
 * @returns 엑셀 시트 이름과 엑셀시트에 들어갈 데이터
 */
function sheet_to_workbook(sheet /*:Worksheet*/, opts) /*:Workbook*/ {
  var n = opts && opts.sheet ? opts.sheet : 'Sheet1';
  var sheets = {};
  sheets[n] = sheet;
  return { SheetNames: [n], Sheets: sheets };
}

/**
 * html 돔 데이터를 엑셀시트에 옮겨 적는 메서드
 * @param title layoutHeaderName
 * @param table tableRef.current
 * @param _opts 맨위의 opt 설명과 같습니다.
 * @returns {ws: 데이터가 로우, 컬럼에 맞게 저장된 워크싯, colLen: 최대 컬럼 길이, headerRowLen: 헤더의 로우 길이, totalRowLen: 헤더 포함 총 로우 길이} html 돔 데이터를 엑셀시트에 옮겨 적은 객체
 */
function parse_dom_table(title /*:?string*/, table /*:HTMLElement*/, _opts /*:?any*/) /*:Worksheet*/ {
  var opts = _opts || {};
  // dense가 없으니 dictionary에. 이거는 sheet의 워크싯이된다.
  var ws /*:Worksheet*/ = opts.dense ? ([] /*:any*/) : ({} /*:any*/);
  return sheet_add_dom(title, ws, table, _opts);
}

/**
 * 테이블의 전체 데이터(JSON형식)를 (전체 페이지) 엑셀시트에 옮겨 적는 메서드
 * @param title layoutHeaderName
 * @param table tableRef.current
 * @param noTitleHeader 시트 내에 맨 위의 헤더타이틀이 안뜨게 하고 싶으실 경우 true
 * @param _opts 맨위의 opt 설명과 같습니다.
 * @returns {ws: 데이터가 로우, 컬럼에 맞게 저장된 워크싯, colLen: 최대 컬럼 길이, totalRowLen: 헤더 포함 총 로우 길이} html 돔 데이터를 엑셀시트에 옮겨 적은 객체
 */
function parse_json(title /*:?string*/, wholeData /*:HeaderArray*/, noTitleHeader /*:? boolean */, _opts /*:?any*/) /*:Worksheet*/ {
  // opts: sheet 정보.
  var opts = _opts || {};
  // dense가 없으니 dictionary에. 이거는 sheet의 워크싯이된다.
  var ws /*:Worksheet*/ = opts.dense ? ([] /*:any*/) : ({} /*:any*/);
  return sheet_add_json(title, ws, wholeData, noTitleHeader, _opts);
}

/**
 * table에 담긴 데이터와 태그 종류, 속성(colspan, rowspan 등)을 사용하여 데이터를 그대로 옮기는 메서드
 * @param title layoutHeaderName
 * @param ws 이제 워크싯 정보를 담을 빈 객체
 * @param table tableRef.current
 * @param _opts 맨위의 opt 설명과 같습니다.
 * @returns {ws: 데이터가 로우, 컬럼에 맞게 저장된 워크싯, colLen: 최대 컬럼 길이, headerRowLen: 헤더의 로우 길이, totalRowLen: 헤더 포함 총 로우 길이} html 돔 데이터를 엑셀시트에 옮겨 적은 객체
 *
 * {1} : or_R은 엑셀 시트에서 데이터 시작 로우 인덱스이며 or_C는 데이터 시작 컬럼 인덱스입니다.
 *      1-2번 로우는 제목이 들어가므로 4번째 줄(3번 인덱스)부터 데이터를 시작하도록 하였습니다.
 * {2} : 테이블에서 tr들을 배열에 담습니다.
 * {3} : Math.min() 함수는 주어진 숫자들 중 가장 작은 값을 반환합니다. 여기서 sheetRows는 그냥 rows.length 가 됩니다.
 * {4} : 밑으로 가시면서 많이 보시게 될 구조 입니다. range는 모든 데이터를 담았을 때의 전체 범위 입니다. 최대 로우, 최대 컬럼이 담길 것입니다.
 *      => { s(start): { r(row index): 0, c(col index): 0 }, e(end): { r(row index): or_R, c(col index): or_C } };
 * {5} : merges는 데이터들의 rowspan, colspan 정보를 담는 객체배열입니다.
 * {6} : 이또한 밑으로 가시면서 많이 보시게 될 구조 입니다.
 *      => { t(type)): 's(string)', v(value): title }
 *          titleO는 자료형은 string, 값을 title이라는 객체입니다.
 * {7} : encode_cell은 데이터를 엑셀시트에 적는 행위라고 보시면 됩니다.
 *      => ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = titleO; 특정 좌표에 자료형과 값으로 이루어진 객체를({5} 참고) 집어넣었습니다.
 *          즉, 0번 컬럼, 0번 로우에 제목을 적겠다는 의미입니다.
 * {8} : 모든 tr들을 돌면서 그 안의 정보들을 활용합니다.
 *  {8-1} : 한 row 즉, tr태그
 *  {8-2} : 테이블의 데이터 중 hidden 처리 된 것이 있다면 감추려는 목적입니다.
 *  {8-3} : elements 즉 tr태그의 자식 태그들
 *  {8-4} : tr태그의 자식 태그들을 돌면서 값, 속성 등의 데이터를 추출하여 ~에 담습니다. => td 혹은 th
 *      {8-4-1} : v => 값 사용자에 보이는 한 td의 text value 혹은 labelName
 *      {8-4-2} : 이전에 merge라는 개체에 들어 간 colspan이나 rowspan이 있다면 그것을 계산 하여 그 다음 셀에 데이터를 넣어주는 곳입니다.
 *  {8-5} : colSpan 속성을 가져옵니다. 없다면 그냥 1
 *  {8-6} : rowspan 속성을 가져옵니다. 없다면 그냥 1
 *          그리고 cs나 rs가 1보다 크다면 셀을 합쳐야 한다는 뜻이니 merges에 정보를 입력하여 셀 병합 데이터를 입력합니다.
 *  {8-7} : 숫자면 숫자, 날짜면 날짜 포매팅 해주는 곳이지만 이미 관리자 화면에 포매팅이 되어있으므로 주석처리 하였습니다. => rowDef의 format 함수 사용하여 포매팅 처리하였습니다.
 *  {8-8} : merge 되었을 때 merge 된 컬럼인덱스까지로 한 셀을 병합시켜줍니다.
 *  {8-9} : 현재 셀의 colspan을 반영하여 다음 컬럼으로 이동합니다.
 * {9} : merge 된 것이 있다면 워크시트에서 병합하는것을 담당하는 ws['!merges']에 그 동안 축적했던 merges 데이터를 합칩니다.
 * {10} : unshift() 메서드는 새로운 요소를 배열의 맨 앞쪽에 추가하고, 새로운 길이를 반환합니다.
 *          이 부분은 이미 셀 병합한 것이 있을 경우 제목의 범위를 위에 2줄 + 최대 컬럼 수 까지 키워주는 곳입니다.
 * {11} : 이 부분은 이미 셀 병합한 것이 없을 경우 제목의 범위를 위에 2줄 + 최대 컬럼 수 까지 키워주는 곳입니다.
 * {12} : ws['!ref']는 엑셀 시트에 데이터를 쓸 범위를 의미합니다.
 * {13} : 전체 데이터를 집어 넣는 곳 마지막에 한번만 합니다.
 * {14} : 추후 헤더는 별도의 스타일링이 들어가므로 이처럼 몇 row 인지 계산 합니다.
 */
function sheet_add_dom(title /*:?string*/, ws /*:Worksheet*/, table /*:HTMLElement*/, _opts /*:?any*/) /*:Worksheet*/ {
  //eslint-disable-line no-unused-vars
  var opts = _opts || {};
  if (DENSE != null) opts.dense = DENSE;
  /* {1} */
  var or_R = 3,
    or_C = 0;
  // opts는 그저 빈 객체이므로 PASS
  if (opts.origin != null) {
    if (typeof opts.origin == 'number') or_R = opts.origin;
    else {
      var _origin /*:CellAddress*/ = typeof opts.origin == 'string' ? XLSX.utils.decode_cell(opts.origin) : opts.origin;
      console.log(`_origin => ${_origin}`);
      or_R = _origin.r;
      or_C = _origin.c;
    }
  }
  // console.log(`constructror 즉, 해당 태그의 종류 => ${table.constructor.name}`);
  // console.log(`HTMLTableElement => ${table instanceof HTMLTableElement}`);
  /* {2} */
  var rows /*:HTMLCollection<HTMLTableRowElement>*/ = table.getElementsByTagName('tr'); // sheet_add_json과 다른 부분
  /* {3} */
  var sheetRows = Math.min(opts.sheetRows || 10000000, rows.length);
  /* {4} */
  var range /*:Range*/ = { s: { r: 0, c: 0 }, e: { r: or_R, c: or_C } };
  // ws는 현재 그저 빈 객체이므로 PASS
  if (ws['!ref']) {
    var _range /*:Range*/ = XLSX.utils.decode_range(ws['!ref']);
    range.s.r = Math.min(range.s.r, _range.s.r);
    range.s.c = Math.min(range.s.c, _range.s.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    range.e.c = Math.max(range.e.c, _range.e.c);
    if (or_R == -1) range.e.r = or_R = _range.e.r + 1;
  }
  /* {5} */
  var merges /*:Array<Range>*/ = [],
    midx = 0;
  var rowinfo /*:Array<RowInfo>*/ = ws['!rows'] || (ws['!rows'] = []); // sheet_add_json과 다른 부분
  var _R = 0,
    R = 0,
    _C = 0,
    C = 0,
    RS = 0,
    CS = 0;
  /* {6} */
  var titleO /*:Cell*/ = { t: 's', v: title };
  // ws는 현재 그저 빈 객체이므로 PASS
  if (opts.dense) {
    if (!ws[0]) ws[0] = [];
    ws[0][0] = titleO;
  } else {
    /* {7} */
    ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = titleO;
  }
  // ws[XLSX.utils.encode_cell({c:C + or_C, r:R + or_R})] = o;
  if (!ws['!cols']) ws['!cols'] = [];
  /* {8} */
  for (; _R < rows.length && R < sheetRows; ++_R) {
    /* {8-1} */
    var row /*:HTMLTableRowElement*/ = rows[_R];
    console.log(`row => ${row}`); // sheet_add_json과 다른 부분
    /* {8-2} */
    if (is_dom_element_hidden(row)) {
      // sheet_add_json과 다른 부분
      if (opts.display) continue;
      rowinfo[R] = { hidden: true };
    }
    /* {8-3} */
    var elts /*:HTMLCollection<HTMLTableCellElement>*/ = (row.children /*:any*/); // sheet_add_json과 다른 부분
    console.log(`elts => ${elts}`);
    /* {8-4} */
    for (_C = C = 0; _C < elts.length; ++_C) {
      var elt /*:HTMLTableCellElement*/ = elts[_C];
      // opts가 비어있으므로 PASS
      if (opts.display && is_dom_element_hidden(elt)) continue; // sheet_add_json과 다른 부분
      /* {8-4-1} */
      var v /*:?string*/ = elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmldecode(elt.innerHTML); // 다름
      console.log(`v(textContent) => ${v}`);
      // z는 뭘해도 null 이 나오는 애입니다.
      var z /*:?string*/ = elt.getAttribute('data-z') || elt.getAttribute('z'); // sheet_add_json과 다른 부분
      console.log(`z => ${z}`);
      /* {8-4-2} */
      for (midx = 0; midx < merges.length; ++midx) {
        var m /*:Range*/ = merges[midx];
        console.log(`m1 => ${JSON.stringify(m)}`);
        if (m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) {
          C = m.e.c + 1 - or_C;
          midx = -1;
        }
        console.log(`m2 => ${JSON.stringify(m)}`);
      }
      /* TODO: figure out how to extract nonstandard mso- style */
      /* {8-5} */
      CS = +elt.getAttribute('colspan') || 1; // sheet_add_json과 다른 부분
      console.log(`CS => ${CS}`);
      /* {8-6} */
      if ((RS = +elt.getAttribute('rowspan') || 1) > 1 || CS > 1) merges.push({ s: { r: R + or_R, c: C + or_C }, e: { r: R + or_R + (RS || 1) - 1, c: C + or_C + (CS || 1) - 1 } }); // sheet_add_json과 다른 부분
      console.log(`RS => ${RS}`);
      // {6} 참고
      var o /*:Cell*/ = { t: 's', v: v };
      // _t, 얘도 뭘해도 그냥 빈칸이 나옵니다.
      var _t /*:string*/ = elt.getAttribute('data-t') || elt.getAttribute('t') || ''; // sheet_add_json과 다른 부분
      console.log(`t => ${_t}`);
      /* {8-7} */
      if (v != null) {
        if (v.length == 0) o.t = _t || 'z';
        else if (v === 'TRUE') o = { t: 'b', v: true };
        else if (v === 'FALSE') o = { t: 'b', v: false };
      }
      console.log(`o2 => ${JSON.stringify(o)}`);
      if (o.z === undefined && z != null) o.z = z;
      /* The first link is used.  Links are assumed to be fully specified.
       * TODO: The right way to process relative links is to make a new <a> */
      // a 태그 다루는 곳
      var l = '',
        Aelts = elt.getElementsByTagName('A');
      if (Aelts && Aelts.length)
        for (var Aelti = 0; Aelti < Aelts.length; ++Aelti)
          if (Aelts[Aelti].hasAttribute('href')) {
            // sheet_add_json과 다른 부분
            l = Aelts[Aelti].getAttribute('href');
            if (l.charAt(0) != '#') break;
          }
      if (l && l.charAt(0) != '#') o.l = { Target: l }; // sheet_add_json과 다른 부분
      // opts는 그저 빈 객체이므로 PASS
      if (opts.dense) {
        if (!ws[R + or_R]) ws[R + or_R] = [];
        ws[R + or_R][C + or_C] = o;
      } else ws[XLSX.utils.encode_cell({ c: C + or_C, r: R + or_R })] = o;
      console.log(`range.e.c : C + or_C => ${range.e.c} : ${C + or_C}`);
      /* {8-8} */
      if (range.e.c < C + or_C) range.e.c = C + or_C;
      /* {8-9} */
      C += CS;
    }
    ++R;
  }
  console.log(`merges => ${JSON.stringify(merges)}`);
  /* {9} */
  if (merges.length) ws['!merges'] = (ws['!merges'] || []).concat(merges);
  range.e.r = Math.max(range.e.r, R - 1 + or_R);
  // const titleRange = {"s":{"r":0,"c":0},"e":{"r":1,"c":range.e.c}};
  /* {10} */
  if (ws['!merges']) ws['!merges'].unshift({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
  /* {11} */ else ws['!merges'] = (ws['!merges'] || []).concat({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
  console.log(`range => ${JSON.stringify(range)}`);
  /* {12} */
  ws['!ref'] = XLSX.utils.encode_range(range);
  // We can count the real number of rows to parse but we don't to improve the performance
  /* {13} */
  if (R >= sheetRows) ws['!fullref'] = XLSX.utils.encode_range(((range.e.r = rows.length - _R + R - 1 + or_R), range));
  /* {14} */
  const thead = table.tHead; // sheet_add_json과 다른 부분
  const headerRowLen = thead.rows.length; // sheet_add_json과 다른 부분
  const res = {
    // sheet_add_json과 다른 부분
    // worksheet
    ws: ws,
    // 한 줄 최대 컬럼 길이
    colLen: range.e.c,
    // 헤더 줄 수
    headerRowLen: headerRowLen,
    // 헤더 포함 모든 로우 수
    totalRowLen: sheetRows
  };
  return res;
}
/**
 *
 * @param title layoutHeaderName
 * @param ws 이제 워크싯 정보를 담을 빈 객체
 * @param {*} wholeData
 * @param noTitleHeader 시트 내에 맨 위의 헤더타이틀이 안뜨게 하고 싶으실 경우 true
 * @param _opts 맨위의 opt 설명과 같습니다.
 * @returns {ws: 데이터가 로우, 컬럼에 맞게 저장된 워크싯, colLen: 최대 컬럼 길이, totalRowLen: 헤더 포함 총 로우 길이} html 돔 데이터를 엑셀시트에 옮겨 적은 객체
 *
 * {1} : or_R은 엑셀 시트에서 데이터 시작 로우 인덱스이며 or_C는 데이터 시작 컬럼 인덱스입니다.
 *      1-2번 로우는 제목이 들어가므로 4번째 줄(3번 인덱스)부터 데이터를 시작하도록 하였습니다.
 * {2} : 헤더와 바디를 합친 배열입니다.
 * {3} : Math.min() 함수는 주어진 숫자들 중 가장 작은 값을 반환합니다. 여기서 sheetRows는 그냥 rows.length 가 됩니다.
 * {4} : 밑으로 가시면서 많이 보시게 될 구조 입니다. range는 모든 데이터를 담았을 때의 전체 범위 입니다. 최대 로우, 최대 컬럼이 담길 것입니다.
 *      => { s(start): { r(row index): 0, c(col index): 0 }, e(end): { r(row index): or_R, c(col index): or_C } };
 * {5} : merges는 데이터들의 rowspan, colspan 정보를 담는 객체배열입니다.
 * {6} : 이또한 밑으로 가시면서 많이 보시게 될 구조 입니다.
 *      => { t(type)): 's(string)', v(value): title }
 *          titleO는 자료형은 string, 값을 title이라는 객체입니다.
 * {7} : encode_cell은 데이터를 엑셀시트에 적는 행위라고 보시면 됩니다.
 *      => ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = titleO; 특정 좌표에 자료형과 값으로 이루어진 객체를({5} 참고) 집어넣었습니다.
 *          즉, 0번 컬럼, 0번 로우에 제목을 적겠다는 의미입니다.
 * {8} : 모든 tr들을 돌면서 그 안의 정보들을 활용합니다.
 *  {8-1} : 한 row 즉, wholeData[idx]
 *  {8-2} : 한 row의 값들을 돌변서 labelName 및 속성들을 활용합니다.
 *      {8-2-1} : elements 즉 한 row의 각 data들
 *      {8-2-2} : v => 값 사용자에 보이는 한 td의 text value 혹은 labelName
 *      {8-2-3} : 이전에 merge라는 개체에 들어 간 colspan이나 rowspan이 있다면 그것을 계산 하여 그 다음 셀에 데이터를 넣어주는 곳입니다.
 *  {8-3} : colSpan 속성을 가져옵니다. 없다면 그냥 1
 *  {8-4} : rowspan 속성을 가져옵니다. 없다면 그냥 1
 *  {8-5} : 숫자면 숫자, 날짜면 날짜 포매팅 해주는 곳  => rowDef의 format 함수 사용하여 포매팅 처리하였습니다.
 *  {8-6} : merge 되었을 때 merge 된 컬럼인덱스까지로 한 셀을 병합시켜줍니다.
 *  {8-7} : 현재 셀의 colspan을 반영하여 다음 컬럼으로 이동합니다.
 * {9} : merge 된 것이 있다면 워크시트에서 병합하는것을 담당하는 ws['!merges']에 그 동안 축적했던 merges 데이터를 합칩니다.
 * {10} : unshift() 메서드는 새로운 요소를 배열의 맨 앞쪽에 추가하고, 새로운 길이를 반환합니다.
 *          이 부분은 이미 셀 병합한 것이 있을 경우 제목의 범위를 위에 2줄 + 최대 컬럼 수 까지 키워주는 곳입니다.
 *          헤더타이틀을 안 넣을 옵션을 줄 시는 실행하지 않습니다.
 * {11} : 이 부분은 이미 셀 병합한 것이 없을 경우 제목의 범위를 위에 2줄 + 최대 컬럼 수 까지 키워주는 곳입니다.
 * {12} : ws['!ref']는 엑셀 시트에 데이터를 쓸 범위를 의미합니다.
 * {13} : 전체 데이터를 집어 넣는 곳 마지막에 한번만 합니다.
 * {14} : 헤더 로우 수에 대한 정보는 wholeData에 헤더와 바디의 구분이 없으니 이 함수를 호출한 곳에서 계산해줍니다.
 *
 * {15} : 헤더 로우 수에 대한 정보는 wholeData에 헤더와 바디의 구분이 없으니 이 함수를 호출한 곳에서 계산해줍니다.
 *    {15-1} : 각 컬럼의 width를 담는 배열
 *    {15-2} : 문자열의 바이트 사이즈를 구하고 이에 2를 더한다. 2를 더하는 이유는 컬럼의 너비에 조금 여유를 주기 위함이다.
 *      {15-2-1} : 내용이 긴 열은 40,50,50에서 너비 조정
 *    {15-3} : 문자열 길이의 최대값을 지정함.
 */
function sheet_add_json(title /*:?string*/, ws /*:Worksheet*/, wholeData /*: wholeData*/, noTitleHeader /*:? boolean */, _opts /*:?any*/) /*:Worksheet*/ {
  //eslint-disable-line no-unused-vars
  var opts = _opts || {};
  if (DENSE != null) opts.dense = DENSE;
  /* {1} */
  var or_R = noTitleHeader ? 0 : 3,
    or_C = 0;
  // opts는 그저 빈 객체이므로 PASS
  if (opts.origin != null) {
    if (typeof opts.origin == 'number') or_R = opts.origin;
    else {
      var _origin /*:CellAddress*/ = typeof opts.origin == 'string' ? XLSX.utils.decode_cell(opts.origin) : opts.origin;
      console.log(`_origin => ${_origin}`);
      or_R = _origin.r;
      or_C = _origin.c;
    }
  }
  /* {2} */
  var rows /*:wholeData*/ = wholeData; // sheet_add_dom과 다른 부분
  // console.log("rows => ", rows) // 주석을 해제해보세요.
  /**
   * ex)
   * [[{"rowpan":2,"value":"No."},{"colSpan":2,"value":"A"},{"colSpan":2,"value":"B"}],
   * ["카드번호","상품명","유효기간","카드발급신청일자"]]
   */
  /* {3} */
  var sheetRows = Math.min(opts.sheetRows || 10000000, rows.length);
  /* {4} */
  var range /*:Range*/ = { s: { r: 0, c: 0 }, e: { r: or_R, c: or_C } };
  // ws는 현재 그저 빈 객체이므로 PASS
  if (ws['!ref']) {
    var _range /*:Range*/ = XLSX.utils.decode_range(ws['!ref']);
    range.s.r = Math.min(range.s.r, _range.s.r);
    range.s.c = Math.min(range.s.c, _range.s.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    range.e.c = Math.max(range.e.c, _range.e.c);
    if (or_R == -1) range.e.r = or_R = _range.e.r + 1;
  }
  /* {5} */
  var merges /*:Array<Range>*/ = [],
    midx = 0;
  var _R = 0,
    R = 0,
    _C = 0,
    C = 0,
    RS = 0,
    CS = 0;
  /* {6} */
  var titleO /*:Cell*/ = { t: 's', v: title };
  /* {15} */
  let colLengthArr;
  if (opts.dense) {
    if (!ws[0]) ws[0] = [];
    ws[0][0] = titleO;
  } else if (!noTitleHeader) {
    /* {7} */
    ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = titleO;
  }
  if (!ws['!cols']) ws['!cols'] = [];
  /* {15-1} */
  colLengthArr = new Array(rows[rows.length - 1].length);
  /* {8} */
  for (; _R < rows.length && R < sheetRows; ++_R) {
    /* {8-1} */
    var row /*:WholeData*/ = rows[_R];
    // 여기서는 wholeData의 각 element들 value가 있으면 하고 없으면 넘어가고 해야겠다.
    /* {8-2} */
    for (_C = C = 0; _C < row.length; ++_C) {
      /* {8-2-1} */
      var elt = row[_C];
      // console.log(`elt => ${JSON.stringify(elt)}`);
      /* {8-2-2} */
      var v /*:?string*/ = elt.value ? elt.value : elt; // sheet_add_dom과 다른 부분
      // console.log(`v(textContent) => ${v}`);
      // z는 뭘해도 null 이 나오는 애입니다.
      var z /*:?string*/ = null; // sheet_add_dom과 다른 부분
      /* {8-2-3} */
      for (midx = 0; midx < merges.length; ++midx) {
        var m /*:Range*/ = merges[midx];
        // console.log(`m => ${JSON.stringify(m)}`);
        if (m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) {
          C = m.e.c + 1 - or_C;
          midx = -1;
        }
        // console.log(`2 => ${JSON.stringify(m)}`);
      }
      /* TODO: figure out how to extract nonstandard mso- style */
      /* {8-3} */
      CS = elt.colspan ?? 1; // sheet_add_dom과 다른 부분
      // console.log(`CS => ${CS}`);
      /* {8-4} */
      if ((RS = +elt.rowspan || 1) > 1 || CS > 1) merges.push({ s: { r: R + or_R, c: C + or_C }, e: { r: R + or_R + (RS || 1) - 1, c: C + or_C + (CS || 1) - 1 } }); // sheet_add_dom과 다른 부분
      // console.log(`RS => ${RS}`);

      /* {15-2} */
      let columnWidth = byteSize(v, true) + 2;
      /* {15-2-1} */
      if (columnWidth >= 42 && columnWidth < 52) columnWidth = 42;
      else if (columnWidth >= 52 && columnWidth < 62) columnWidth = 52;
      else if (columnWidth > 62) columnWidth = 62;
      /* {15-3} */
      if (colLengthArr[C] < columnWidth || colLengthArr[C] === undefined) {
        colLengthArr[C] = columnWidth;
      }
      // {6} 참고
      var o /*:Cell*/ = { t: 's', v: v };
      // _t, 얘도 뭘해도 그냥 빈칸이 나옵니다.
      var _t /*:string*/ = ''; // sheet_add_dom과 다른 부분
      /* {8-5} */
      if (v != null) {
        if (v.length == 0) o.t = _t || 'z';
        else if (v === 'TRUE') o = { t: 'b', v: true };
        else if (v === 'FALSE') o = { t: 'b', v: false };
      }
      // console.log(`o => ${JSON.stringify(o)}`);
      if (o.z === undefined && z != null) o.z = z;
      /* The first link is used.  Links are assumed to be fully specified.
       * TODO: The right way to process relative links is to make a new <a> */
      // opts는 그저 빈 객체이므로 PASS
      if (opts.dense) {
        if (!ws[R + or_R]) ws[R + or_R] = [];
        ws[R + or_R][C + or_C] = o;
      } else ws[XLSX.utils.encode_cell({ c: C + or_C, r: R + or_R })] = o;
      /* {8-6} */
      if (range.e.c < C + or_C) range.e.c = C + or_C;
      /* {8-7} */
      C += CS;
    }
    ++R;
  }
  // console.log(`merges => ${JSON.stringify(merges)}`);
  /* {9} */
  if (merges.length) ws['!merges'] = (ws['!merges'] || []).concat(merges);
  range.e.r = Math.max(range.e.r, R - 1 + or_R);
  /* {10} */
  if (!noTitleHeader) {
    if (ws['!merges']) ws['!merges'].unshift({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
    /* {11} */ else ws['!merges'] = (ws['!merges'] || []).concat({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
  }
  // console.log(`range => ${JSON.stringify(range)}`);
  /* {12} */
  ws['!ref'] = XLSX.utils.encode_range(range);
  /* {13} */
  if (R >= sheetRows) ws['!fullref'] = XLSX.utils.encode_range(((range.e.r = rows.length - _R + R - 1 + or_R), range)); // We can count the real number of rows to parse but we don't to improve the performance
  /* {14} */
  const res = {
    // sheet_add_dom과 다른 부분
    // worksheet
    ws: ws,
    // 한 줄 최대 컬럼 길이
    colLen: range.e.c,
    // 헤더 줄 수는 바깥에서.
    // 헤더 포함 모든 로우 수
    totalRowLen: sheetRows,
    // 각 컬럼의 너비를 담은 배열
    colLengthArr: colLengthArr
  };
  return res;
}

// sheet Js 소스코드에서 가져왔습니다. 돔에 hidden 속성이 있는지 확인
function is_dom_element_hidden(element /*:HTMLElement*/) /*:boolean*/ {
  var display /*:string*/ = '';
  var get_computed_style /*:?function*/ = get_get_computed_style_function(element);
  if (get_computed_style) display = get_computed_style(element).getPropertyValue('display');
  if (!display) display = element.style && element.style.display;
  return display === 'none';
}
// sheet Js 소스코드에서 가져왔습니다.
function get_get_computed_style_function(element /*:HTMLElement*/) /*:?function*/ {
  // The proper getComputedStyle implementation is the one defined in the element window
  if (element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
  // If it is not available, try to get one from the global namespace
  if (typeof getComputedStyle === 'function') return getComputedStyle;
  return null;
}
// sheet Js 소스코드에서 가져왔습니다. textContent 뽑는애인듯합니다.
var htmldecode /*:{(s:string):string}*/ = /*#__PURE__*/ (function () {
  var entities /*:Array<[RegExp, string]>*/ = [
    ['nbsp', ' '],
    ['middot', '·'],
    ['quot', '"'],
    ['apos', "'"],
    ['gt', '>'],
    ['lt', '<'],
    ['amp', '&']
  ].map(function (x /*:[string, string]*/) {
    return [new RegExp('&' + x[0] + ';', 'ig'), x[1]];
  });
  return function htmldecode(str /*:string*/) /*:string*/ {
    var o = str
      // Remove new lines and spaces from start of content
      .replace(/^[\t\n\r ]+/, '')
      // Remove new lines and spaces from end of content
      .replace(/[\t\n\r ]+$/, '')
      // Added line which removes any white space characters after and before html tags
      .replace(/>\s+/g, '>')
      .replace(/\s+</g, '<')
      // Replace remaining new lines and spaces with space
      .replace(/[\t\n\r ]+/g, ' ')
      // Replace <br> tags with new lines
      .replace(/<\s*[bB][rR]\s*\/?>/g, '\n')
      // Strip HTML elements
      .replace(/<[^>]*>/g, '');
    for (var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
    return o;
  };
})();

// 추가 작성: 날짜데이터인데 월까지만 나오는 것인지 확인 후 맞다면 2023-03 형태로 반환
// function checkMonthIsDue(str) {
//   const reg = /\d{4}-?\d{2}/i;
//   console.log(`어음 ${str.match(reg)}`);
//   const res = str.match(reg) == str;
//   console.log(`어음 ${res}`);
//   return res;
// }

/**
 * 데이터로 된 엑셀 객체를 xlsx 파일 바이너리타입으로 변환한다.
 * @param workbook 엑셀 객체
 * @returns 바이너리 타입으로 변환된 xlsx 파일
 */
const workbook2blob = (workbook) => {
  const wopts = {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  };
  // 엑셀 파일로 변환
  const wbout = XLSX.write(workbook, wopts);
  // The application/octet-stream MIME type is used for unknown binary files.
  // It preserves the file contents, but requires the receiver to determine file type,
  // for example, from the filename extension.
  // 블롭 형태로 하여 다운로드 받기 위해 준비시킨다.
  const blob = new Blob([s2ab(wbout)], {
    type: 'application/octet-stream'
  });
  return blob;
};

// 입력한 걸 파일로 만드는 작업을 여기서 진행
const s2ab = (s) => {
  // The ArrayBuffer() constructor is used to create ArrayBuffer objects.
  // create an ArrayBuffer with a size in bytes
  const buf = new ArrayBuffer(s.length);
  //create a 8 bit integer array
  const view = new Uint8Array(buf);
  // 입력한 걸 파일로 만드는 작업을 여기서 진행
  //charCodeAt The charCodeAt() method returns an integer between 0 and 65535 representing the UTF-16 code
  for (let i = 0; i !== s.length; ++i) {
    //   console.log(s.charCodeAt(i));
    view[i] = s.charCodeAt(i);
  }
  return buf;
};

/**
 *
 * @param wb 엑셀파일이 될 객체
 * @param wsInfo 스타일링을 위해서 colLen, headerRowLen, totalRowLen 이 담겨있는 객체
 * @returns addStyle => 스타일을 먹이고 엑셀파일로 다운로드 받게 하는 메서드
 */
const handleExport = (wb, wsInfo, noTitleHeader) => {
  const workbookBlob = workbook2blob(wb);
  const colLen = wsInfo.colLen;
  const colAlphabet = getColAlphabetFromIdx(colLen);
  const dataStartIndex = noTitleHeader ? 0 : 3;
  const dataInfo = {
    titleRange: noTitleHeader ? undefined : `A1:${colAlphabet}2`,
    tbodyRange: `A${dataStartIndex + 1}:${colAlphabet}${dataStartIndex + wsInfo.totalRowLen}`,
    theadRange: `A${dataStartIndex + 1}:${colAlphabet}${dataStartIndex + wsInfo.headerRowLen}`
  };
  // return addStyle(workbookBlob, dataInfo, colLen);
  return addStyle(workbookBlob, dataInfo, wsInfo);
};
/**
 * 스타일링할 범위를 정해야 하는데, 이제부터는 숫자 인덱스 형태가 아닌 알파벳이 들어가야 한다.
 * @param colLen 최대 컬럼 길이
 * @returns 알파벳
 */
const getColAlphabetFromIdx = (maxColIdx) => {
  // ASCII 코드 A-Z => 65 ~ 90
  /** colLen은 최대, 즉 column의 최대 인덱스이다. 0부터 시작
     A : 65 + 0(colLen)
     B : 65 + 1(colLen)
     ... 
     Z : 65 + 25(colLen)
     AA : 65 + 26(colLen)
     => 25를 넘어가지 않는다면 65 + colLen한걸 ascii코드 사용해서 알파벳으로 변환시켜 반환하면된다.
     => 하지만 만약 25를 넘어간다면? colLen을 26으로 나눈 몫을 알파벳 앞자리에 두자. 
     => 그리고 65에 colLen을 25로 나눈 나머지를 더하는 것으로 한다.
     * */
  let length = maxColIdx + 1;
  let column = '';
  while (length > 0) {
    let remainder = (length - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    length = Math.floor((length - 1) / 26);
  }
  return column;
};

/**
 *
 * @param workbookBlob 바이너리 타입으로 변환된 xlsx 파일
 * @param {titleRange: 제목 범위, tbodyRange: 본문 범위, theadRange: 헤더 범위} dataInfo
 * @param wsInfo 스타일 정보
 * @returns 비동기로 엑셀파일을 다운로드 받게 한다.
 */
const addStyle = (workbookBlob, dataInfo, wsInfo) => {
  return XlsxPopulate.fromDataAsync(workbookBlob).then((workbook) => {
    workbook.sheets().forEach((sheet) => {
      sheet.usedRange().style({
        fontFamily: '맑은 고딕',
        fontSize: 11,
        verticalAlignment: 'center'
      });
      // 바디가 존재한다면 바디의 범위를 아래와 같은 속성들로 스타일링 한다.
      if (dataInfo.tbodyRange) {
        // 문자열의 길이가 같은 것은 가운데 정렬
        wsInfo.equalLengthColumnIndexList.forEach((idx) => {
          sheet.column(idx + 1).style({ horizontalAlignment: 'center' });
        });
        // 숫자 포맷은 오른쪽 정렬
        wsInfo.numberFormatIndexList.forEach((idx) => {
          // sheet.column(idx + 1).style({ horizontalAlignment: 'right', indent: 0.5 }); // 들여쓰기 포함 된 것
          sheet.column(idx + 1).style({ horizontalAlignment: 'right' });
        });
        // 그 외 바디 스타일 설정
        sheet.range(dataInfo.tbodyRange).style({
          border: true // 경계선 있음
          // wrapText: true // 자동 줄 바꿈
        });
      }
      // abcde =>
      // 모든 컬럼의 너비를 colLengthArr를 통해 지정
      for (let i = 0; i <= wsInfo.colLen; i++) {
        const colWidth = wsInfo.colLengthArr[i] ?? 5;
        sheet.column(getColAlphabetFromIdx(i)).width(colWidth);
      }
      // 제목의 범위를 아래와 같은 속성들로 스타일링 한다.
      // 폰트: 맑은 고딕, 굵게, 폰트사이즈: 18, 열높이:32
      // 옆 높이: 16 * 2 = 32;
      sheet.row(1).height(16);
      sheet.row(2).height(16);
      if (dataInfo.titleRange) {
        sheet.range(dataInfo.titleRange).style({
          underline: true,
          bold: true, // 볼드체
          fontSize: 18,
          horizontalAlignment: 'center' // 가로 중앙 정렬
          // verticalAlignment: 'center'
          // fill: 'FFFD04' // 배경 색
        });
      }

      // 헤더의 범위를 아래와 같은 속성들로 스타일링 한다.
      sheet.range(dataInfo.theadRange).style({
        fill: 'd9e1f2', // 배경 색
        bold: true, // 볼드체
        horizontalAlignment: 'center' // 가로 중앙 정렬
        // indent: 0 // 들여쓰기 포함된 것
        // fontColor: 'ffffff' // 폰트 색
      });

      //     if (dataInfo.theadRange1) {
      //       sheet.range(dataInfo.theadRange1).style({
      //         fill: "808080",
      //         bold: true,
      //         horizontalAlignment: "center",
      //         fontColor: "ffffff",
      //       });
      //     }

      //     if (dataInfo.tFirstColumnRange) {
      //       sheet.range(dataInfo.tFirstColumnRange).style({
      //         bold: true,
      //       });
      //     }

      //     if (dataInfo.tLastColumnRange) {
      //       sheet.range(dataInfo.tLastColumnRange).style({
      //         bold: true,
      //       });
      //     }

      //     if (dataInfo.tFirstColumnRange1) {
      //       sheet.range(dataInfo.tFirstColumnRange1).style({
      //         bold: true,
      //       });
      //     }

      //     if (dataInfo.tLastColumnRange1) {
      //       sheet.range(dataInfo.tLastColumnRange1).style({
      //         bold: true,
      //       });
      //     }
    });

    return workbook.outputAsync().then((workbookBlob) => URL.createObjectURL(workbookBlob));
  });
};

/**
 * table_to_book와 json_to_book를 하나의 메서드로 합치기 위해 만들었던 메서드이나 미완성입니다.
 * @param {*} title
 * @param {*} dataArray
 * @param {*} opts
 * @returns
 */
// function data_to_book(title /*:?string*/, dataArray /*:HTMLElement*/, opts /*:?any*/) /*:Workbook*/ {
//   //eslint-disable-line no-unused-vars
//   let sheetObj = parse_data(title, dataArray, opts);
//   return returnWorkbook(sheetObj, opts);
// }

/**
 * parse_dom_table와 parse_json를 하나의 메서드로 합치기 위해 만들었던 메서드이나 미완성입니다.
 * @param {*} title
 * @param {*} dataArray
 * @param {*} opts
 * @returns
 */
// function parse_data(title /*:?string*/, dataArray /*:HeaderArray*/, _opts /*:?any*/) /*:Worksheet*/ {
//   // opts: sheet 정보.
//   var opts = _opts || {};
//   // dense가 없으니 dictionary에. 이거는 sheet의 워크싯이된다.
//   var ws /*:Worksheet*/ = opts.dense ? ([] /*:any*/) : ({} /*:any*/);
//   return sheet_add_json(title, ws, dataArray, _opts);
// }

// function common_sheet_add_dom(title /*:?string*/, ws /*:Worksheet*/, dataArray /*:HeaderArray*/, _opts /*:?any*/) /*:Worksheet*/ {
//   //eslint-disable-line no-unused-vars
//   var opts = _opts || {};
//   const isDataTable = dataArray instanceof HTMLTableElement;
//   if (DENSE != null) opts.dense = DENSE;
//   var or_R = 3,
//     or_C = 0;
//   if (opts.origin != null) {
//     if (typeof opts.origin == 'number') or_R = opts.origin;
//     else {
//       var _origin /*:CellAddress*/ = typeof opts.origin == 'string' ? XLSX.utils.decode_cell(opts.origin) : opts.origin;
//       console.log(`_origin => ${_origin}`);
//       or_R = _origin.r;
//       or_C = _origin.c;
//     }
//   }
//   // var rows/*:wholeData*/ = dataArray; // 다름
//   let rows = isDataTable ? dataArray.getElementsByTagName('tr') : dataArray;
//   // console.log("rows => ", rows)
//   /**
//    * ex)
//    * [[{"rowpan":2,"value":"No."},{"colSpan":2,"value":"A"},{"colSpan":2,"value":"B"}],
//    * ["카드번호","상품명","유효기간","카드발급신청일자"]]
//    */
//   var sheetRows = Math.min(opts.sheetRows || 10000000, rows.length);
//   var range /*:Range*/ = { s: { r: 0, c: 0 }, e: { r: or_R, c: or_C } };
//   if (ws['!ref']) {
//     var _range /*:Range*/ = XLSX.utils.decode_range(ws['!ref']);
//     range.s.r = Math.min(range.s.r, _range.s.r);
//     range.s.c = Math.min(range.s.c, _range.s.c);
//     range.e.r = Math.max(range.e.r, _range.e.r);
//     range.e.c = Math.max(range.e.c, _range.e.c);
//     if (or_R == -1) range.e.r = or_R = _range.e.r + 1;
//   }
//   var merges /*:Array<Range>*/ = [],
//     midx = 0;
//   var rowinfo /*:Array<RowInfo>*/ = isDataTable ? ws['!rows'] || (ws['!rows'] = []) : undefined; // 다름
//   var _R = 0,
//     R = 0,
//     _C = 0,
//     C = 0,
//     RS = 0,
//     CS = 0;
//   var titleO /*:Cell*/ = { t: 's', v: title };
//   if (opts.dense) {
//     if (!ws[0]) ws[0] = [];
//     ws[0][0] = titleO;
//   } else {
//     ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = titleO;
//   }
//   if (!ws['!cols']) ws['!cols'] = [];
//   for (; _R < rows.length && R < sheetRows; ++_R) {
//     var row /*:HeaderArray*/ = rows[_R];
//     if (isDataTable && is_dom_element_hidden(row)) {
//       // 다름
//       if (opts.display) continue;
//       rowinfo[R] = { hidden: true };
//     }
//     var elts /*:HTMLCollection<HTMLTableCellElement>*/ = isDataTable ? (row.children /*:any*/) : undefined; // 다름
//     // var elts/*:HTMLCollection<HTMLTableCellElement>*/ = (row.children/*:any*/); // => th, td 같은 애들.
//     // 여기서는 wholeData의 각 element들 value가 있으면 하고 없으면 넘어가고 해야겠다.
//     for (_C = C = 0; _C < row.length; ++_C) {
//       var elt = isDataTable ? elts[_C] : row[_C];
//       if (isDataTable && opts.display && is_dom_element_hidden(elt)) continue; // 다름
//       console.log(`elt => ${JSON.stringify(elt)}`);
//       var v /*:?string*/ = isDataTable ? (elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmldecode(elt.innerHTML)) : elt.value ? elt.value : elt; // 다름
//       console.log(`v(textContent) => ${v}`);
//       var z /*:?string*/ = isDataTable ? elt.getAttribute('data-z') || elt.getAttribute('z') : null; // 다름
//       for (midx = 0; midx < merges.length; ++midx) {
//         var m /*:Range*/ = merges[midx];
//         console.log(`m => ${JSON.stringify(m)}`);
//         if (m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) {
//           C = m.e.c + 1 - or_C;
//           midx = -1;
//         }
//         console.log(`2 => ${JSON.stringify(m)}`);
//       }
//       /* TODO: figure out how to extract nonstandard mso- style */
//       CS = isDataTable ? +elt.getAttribute('colspan') ?? 1 : elt.colspan ?? 1; // 다름
//       console.log(`CS => ${CS}`);
//       if ((RS = isDataTable ? +elt.getAttribute('rowspan') || 1 : +elt.rowspan || 1) > 1 || CS > 1) {
//         merges.push({ s: { r: R + or_R, c: C + or_C }, e: { r: R + or_R + (RS || 1) - 1, c: C + or_C + (CS || 1) - 1 } });
//       } // 다름
//       console.log(`RS => ${RS}`);
//       var o /*:Cell*/ = { t: 's', v: v };
//       console.log(`o => ${JSON.stringify(o)}`);
//       var _t /*:string*/ = isDataTable ? elt.getAttribute('data-t') ?? elt.getAttribute('t') ?? '' : ''; // 다름
//       if (v != null) {
//         if (v.length == 0) o.t = _t || 'z';
//         else if (opts.raw || v.toString().trim().length == 0 || _t == 's') {
//           //eslint-disable-line no-unused-vars
//         } else if (v === 'TRUE') o = { t: 'b', v: true };
//         else if (v === 'FALSE') o = { t: 'b', v: false };
//         else if (!isNaN(fuzzynum(v))) o = { t: 'n', v: fuzzynum(v) };
//         else if (!isNaN(fuzzydate(v).getDate())) {
//           o = ({ t: 'd', v: parseDate(v) } /*:any*/);
//           // if(!opts.cellDates) o = ({t:'n', v:(o.v instanceof Date ? datenum(o.v) : o.v )}/*:any*/);
//           if (!opts.cellDates) o = ({ t: o.v instanceof Date ? 'n' : 's', v: o.v instanceof Date ? datenum(o.v) : o.v } /*:any*/);
//           o.z = opts.dateNF || table_fmt[14];
//         }
//       }
//       if (o.z === undefined && z != null) o.z = z;
//       /* The first link is used.  Links are assumed to be fully specified.
//        * TODO: The right way to process relative links is to make a new <a> */
//       if (isDataTable) {
//         var l = '',
//           Aelts = elt.getElementsByTagName('A');
//         if (Aelts && Aelts.length)
//           for (var Aelti = 0; Aelti < Aelts.length; ++Aelti)
//             if (Aelts[Aelti].hasAttribute('href')) {
//               // 다름
//               l = Aelts[Aelti].getAttribute('href');
//               if (l.charAt(0) != '#') break;
//             }
//         if (l && l.charAt(0) != '#') o.l = { Target: l }; // 다름
//       }
//       if (opts.dense) {
//         if (!ws[R + or_R]) ws[R + or_R] = [];
//         ws[R + or_R][C + or_C] = o;
//       } else ws[XLSX.utils.encode_cell({ c: C + or_C, r: R + or_R })] = o;
//       if (range.e.c < C + or_C) range.e.c = C + or_C;
//       C += CS;
//     }
//     ++R;
//   }
//   console.log(`merges => ${JSON.stringify(merges)}`);
//   if (merges.length) ws['!merges'] = (ws['!merges'] || []).concat(merges);
//   range.e.r = Math.max(range.e.r, R - 1 + or_R);
//   if (ws['!merges']) ws['!merges'].unshift({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
//   else ws['!merges'] = (ws['!merges'] || []).concat({ s: { r: 0, c: 0 }, e: { r: 1, c: range.e.c } });
//   console.log(`range => ${JSON.stringify(range)}`);
//   ws['!ref'] = XLSX.utils.encode_range(range);
//   if (R >= sheetRows) ws['!fullref'] = XLSX.utils.encode_range(((range.e.r = rows.length - _R + R - 1 + or_R), range)); // We can count the real number of rows to parse but we don't to improve the performance

//   const res = {
//     // 다름
//     // worksheet
//     ws: ws,
//     // 한 줄 최대 컬럼 길이
//     colLen: range.e.c,
//     // 헤더 줄 수는 바깥에서.
//     // 헤더 포함 모든 로우 수
//     totalRowLen: sheetRows
//   };
//   if (isDataTable) {
//     const thead = dataArray.tHead; // 다름
//     const headerRowLen = thead.rows.length; // 다름
//     res['headerRowLen'] = headerRowLen;
//   }
//   return res;
// }

export { json_to_book, table_to_book, handleExport };
