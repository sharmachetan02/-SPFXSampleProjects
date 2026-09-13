
import { IDropdownOption } from '@fluentui/react';
export enum LogicalOperator{
    And,
    Or
  }
export interface IFilter{
    name:string;
    displayName:string;
    placeHolder:string;
    fields:string[];
    values:IDropdownOption[] | string[];
    selectedValues:IDropdownOption[] | string[];
    operatorBetweenValues:LogicalOperator;
    assert?: (value, filter) => boolean;
}
