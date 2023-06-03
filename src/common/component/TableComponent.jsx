import React, { useCallback } from 'react';
import { no } from 'util/ComponentUtil';
// import { no } from 'util/ComponentUtil';
/**
 * {1}: 테이블 헤더의 정보를 가진 rowDef를 이용하여 hdArray를 만든다. 이후 이를 이용하여 thead를 만든다. arrangedHdArray은 rowDef의 요소들을 순서대로 넣는다.
 *      children이 있을 시 hdArray에는 [[],[]]의 형태가 되지만 arrangedHdArray는 [...]의 형태이다.
 *  {1-1}: 한 depth의 첫번째 요소가 올 때에는 undefined일 것이므로 리스트로 지정해주고 rowDef의 각 요소들을 리스트에 하나씩 push 한다.
 *  {1-2}: 헤더에 두번째 열이 올 경우 다른 그 다음 depth에 추가한다.
 * {2}: hdArray를 이용하여 thead의 th들을 반환하는 함수
 *  {2-1}: rowDef에 colSpan, rowSpan을 지정해주었을 경우 셀을 합친다.
 * {3}: list와 rowDef를 이용하여 tbody > tr > td를 반환하는 함수이다.
 */
const TableComponent = ({ rowDef, list, setList }) => {
  // {1}
  const hdArray = [];
  const arrangedHdArray = [];
  const researchHdArray = useCallback(
    (target, depth) => {
      target?.forEach((e) => {
        // {1-2}
        e?.children ? researchHdArray(e.children, depth + 1) : arrangedHdArray.push(e);
      });
      // {1-1}
      if (hdArray[depth] === undefined) hdArray[depth] = [];
      hdArray[depth].push(target);
    },
    [rowDef]
  );
  researchHdArray(rowDef, 0);
  // {2}
  const getHd = (innerIdx, colSpan, rowSpan, e) => {
    return (
      // {2-1}
      <th key={`${e}${innerIdx}`} className="p-1 border" colSpan={colSpan} rowSpan={rowSpan}>
        {e.labelName}
      </th>
    );
  };
  // {3}
  const getTd = (element, listIndex, headerElement, index, value, colSpan, tdClassName) => {
    const key = `${headerElement.labelName}${listIndex}${element}${index}`;
    // const isNoHide = headerElement.noHideCondition ? headerElement.noHideCondition(element) : false;
    let subtractionNoRemaining = listIndex + 1;
    switch (headerElement.type) {
      case no:
        // subtractionNoRemaining += headerElement.noHideCondition && headerElement.noHideCondition(element) ? 1 : 0;
        return (
          <td key={key} className={`${tdClassName} ${headerElement.textAlign ?? 'text-center'}`} colSpan={colSpan}>
            {/* {!isNoHide && ((paging?.currentPageNo ?? 1) - 1) * defaultPerPageCnt + listIndex + 1 - subtractionNoRemaining} */}
            {subtractionNoRemaining}
          </td>
        );
      default:
        return headerElement.colStyle === 'th' ? (
          <th key={key} className={`${tdClassName} ${headerElement.textAlign ?? 'text-center'} border-bottom-0`} name={headerElement.name} colSpan={colSpan}>
            {value}
          </th>
        ) : (
          <td key={key} className={`${tdClassName} ${headerElement.textAlign ?? 'text-center'}`} name={headerElement.name} colSpan={colSpan}>
            {value}
          </td>
        );
    }
  };

  return (
    <table className="table table-cell table-auto border border-collapse">
      <thead>
        {hdArray.map((e, idx) => {
          return (
            <tr key={`${idx}`} role="row">
              {e.map((coverE) =>
                coverE.map((innerE, innerIdx) => {
                  const colSpan = innerE.colSpan ?? 1;
                  const rowSpan = innerE.rowSpan ?? 1;
                  return <React.Fragment key={`${e}${idx}${innerE}${innerIdx}`}>{getHd(innerIdx, colSpan, rowSpan, innerE)}</React.Fragment>;
                })
              )}
            </tr>
          );
        })}
      </thead>
      <tbody>
        {list.map((element, listIndex) => {
          const elementKey = `${element}${listIndex}`;
          return (
            <tr
              key={elementKey}
              role="row"
              onClick={() => {
                setList && setList({ list });
              }}>
              {arrangedHdArray.map((headerElement, index) => {
                const colSpan = headerElement.colSpan ?? 1;
                // json리스트 중 rowDef의 name으로 지정된 키의 값을 가져와 target 값으로 한다.
                let targetValue = element[headerElement.name];
                if (headerElement.format) targetValue = headerElement.format(targetValue ?? '');
                const borderHide = headerElement.borderHideCondition && headerElement.borderHideCondition(element) ? 'border-right-0 border-left-0' : '';
                const tdClassName = `${borderHide} p-1 border`;
                return colSpan !== 0 && getTd(element, listIndex, headerElement, index, targetValue, colSpan, tdClassName);
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
};

export default TableComponent;
