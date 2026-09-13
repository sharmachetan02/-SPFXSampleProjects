
import * as React from 'react';
import { IAppliedFiltersProps } from './IAppliedFiltersProps';
import { IAppliedFiltersState } from './IAppliedFiltersState';
import styles from '../View.module.scss';
import { IFilter } from '../../../../common/models/IFilter';
import { IconButton, IDropdownOption } from '@fluentui/react';
export default class AppliedFilters extends React.Component<IAppliedFiltersProps, IAppliedFiltersState> {
  constructor(props) {
    super(props);
  }
  public async componentDidMount(): Promise<void> {
    //todo
  }

  public render(): React.ReactElement<IAppliedFiltersProps> {
    const { filters } = this.props;
    const appliedFiltersHtml = filters.map((filter: IFilter) => filter.selectedValues.map((val: string | IDropdownOption) => {
      const label = typeof val === 'string' ? val : (val as IDropdownOption).text;
      return (
        <span className={styles.appliedFilter__item} key={filter.name}>
           <span className={styles.appliedFilter__type} >
            {filter.displayName}
          </span>
          <span className={styles.appliedFilter__text} >
            {label}
          </span>
          <span className={styles.appliedFilter__close}>
            <IconButton iconProps={{ iconName: 'ChromeClose' }} onClick={(e) => { this.removeFilter(e, val, filter); }} />
          </span>
        </span>
      );
    }));
    return (<div className={styles.appliedFilter}>
      {appliedFiltersHtml}
    </div>);
  }
  private removeFilter(e, val: IDropdownOption | string, filter: IFilter) {
    e?.preventDefault?.();
    e?.stopPropagation?.();
    const { callback, filters } = this.props;
    const label = typeof val === 'string' ? val : (val as IDropdownOption).text;
    const index = filter.selectedValues.findIndex((selectedVal) => {
      const filterLabel = typeof selectedVal === 'string' ? selectedVal : (selectedVal as IDropdownOption).text;
      return filterLabel === label;
    });
    filter.selectedValues.splice(index, 1);
    if (filter.name === 'Search') {
      filter.selectedValues = [];
    }
    const updatedFilters = filters.map(((f) => ((f.name === filter.name) ? filter : f)));
    callback(updatedFilters);
    return;
  }
}
