import * as React from 'react';
import { Multiselect } from 'react-widgets';


const MultiselectWrapper: React.FC<any> = (props) => {
  const defaultSelectIcon = (
    <span className="ms-Dropdown-caretDownWrapper">
      <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
    </span>
  );

  const defaultClearTagIcon = (
    <span className="ms-Button-flexContainer ">
      <i data-icon-name="Cancel" aria-hidden="true" className="clearTagIcon"></i>
    </span>
  );

  return (
    <Multiselect
      {...props}
      selectIcon={props.selectIcon || defaultSelectIcon}
      clearTagIcon={props.clearTagIcon || defaultClearTagIcon}
    />
  );
};

export default MultiselectWrapper;
