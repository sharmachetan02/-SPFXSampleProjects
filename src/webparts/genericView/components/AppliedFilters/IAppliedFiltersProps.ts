import { IFilter } from "../../../../common/models/IFilter";
export interface IAppliedFiltersProps {
  filters: IFilter[];
  callback?:  (param:IFilter[]) => void;
}
