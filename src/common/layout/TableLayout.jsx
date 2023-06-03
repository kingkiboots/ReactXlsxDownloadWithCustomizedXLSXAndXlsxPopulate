import ButtonComponent from 'common/component/ButtonComponent';
import TableComponent from 'common/component/TableComponent';
import React from 'react';

const TableLayout = ({ layoutHeaderName, rowDef, buttonsDef, list, setList }) => {
  return (
    <div className="m-8">
      <h2 className="text-[1.5rem]">{layoutHeaderName}</h2>
      {buttonsDef && (
        <div className="btn-group">
          {buttonsDef?.map((element, index) => {
            const elementKey = `${element.labelName}${index}`;
            return <ButtonComponent key={elementKey} option={element} />;
          })}
        </div>
      )}
      <div className="relative overflow-x-auto">
        <TableComponent rowDef={rowDef} list={list} setList={setList} />
      </div>
    </div>
  );
};

export default TableLayout;
