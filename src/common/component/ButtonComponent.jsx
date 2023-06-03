import React from 'react';

const ButtonComponent = ({ option }) => {
  return (
    <button type="button" className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded" onClick={option.onClick} disabled={option.isDisabled} hidden={option.hidden}>
      {option.labelName}
    </button>
  );
};

export default ButtonComponent;
