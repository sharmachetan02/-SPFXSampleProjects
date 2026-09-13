import { DefaultButton, Icon, ITooltipHostStyles, TooltipHost } from "@fluentui/react";
import React from "react";
import styles from './Datasheet.module.scss';
import { getStatusViewProps } from "../../../common/helpers/StatusHelper";

interface ArrayPillsContainerProps<T extends string | number> {
  items: T[];
  format?: (v: T) => string;
  pillClassName?: string;
  maxPills?: number;
}

interface ArrayPillsContainerState {
  isExpanded: boolean;
}

interface ObjectPillsContainerProps<T extends object = Record<string, any>> {
  items: T[];
  renderItem: (item: T) => JSX.Element;
  maxPills?: number;
}

interface ObjectPillsContainerState {
  isExpanded: boolean;
}

class ArrayPillsContainer<T extends string | number> extends React.Component<
  ArrayPillsContainerProps<T>,
  ArrayPillsContainerState
> {
  constructor(props: ArrayPillsContainerProps<T>) {
    super(props);
    this.state = {
      isExpanded: false
    };
  }

  private toggleExpand = () => {
    this.setState((prevState) => ({ isExpanded: !prevState.isExpanded }));
  };

  render() {
    const { items, format, pillClassName, maxPills = 5 } = this.props;
    const { isExpanded } = this.state;
    const fmt = format ?? ((v: T) => (typeof v === 'number' ? v.toLocaleString() : String(v)));

    const values = (items || []).filter((v) => v !== null && v !== undefined) as T[];
    const shown = values.slice(0, maxPills);
    const hidden = values.slice(maxPills);
    const displayItems = isExpanded ? values : shown;

    const tooltipStyles: ITooltipHostStyles = { root: { display: 'inline-block' } };

    if (values.length === 0) {
      return (
        <TooltipHost content="No data" styles={tooltipStyles}>
          <span className={styles.detailsListValue}>
            <Icon iconName="Remove" className={styles.mutedIcon} />
            <span className={styles.srOnly}>No data</span>
          </span>
        </TooltipHost>
      );
    }

    return (
      <div className={styles.termPills}>
        {displayItems.map((v, i) => (
          <span
            key={`pill-${i}-${String(v)}`}
            className={`${styles.termPill} ${pillClassName ?? ''}`.trim()}
            title={fmt(v)}
          >
            <span className={styles.termPillText}>{fmt(v)}</span>
          </span>
        ))}

        {hidden.length > 0 && !isExpanded && (
          <DefaultButton
            onClick={this.toggleExpand}
            className={styles.termPillButton}
          >
            View more +{hidden.length}
          </DefaultButton>
        )}

        {isExpanded && hidden.length > 0 && (
          <DefaultButton
            onClick={this.toggleExpand}
            className={styles.termPillButton}
          >
            View less
          </DefaultButton>
        )}
      </div>
    );
  }
}

class ObjectPillsContainer<T extends object = Record<string, any>> extends React.Component<
  ObjectPillsContainerProps<T>,
  ObjectPillsContainerState
> {
  constructor(props: ObjectPillsContainerProps<T>) {
    super(props);
    this.state = {
      isExpanded: false
    };
  }

  private toggleExpand = () => {
    this.setState((prevState) => ({ isExpanded: !prevState.isExpanded }));
  };

  render() {
    const { items, renderItem, maxPills = 5 } = this.props;
    const { isExpanded } = this.state;

    const validItems = (items || []).filter((item) => item !== null && item !== undefined);
    const shown = validItems.slice(0, maxPills);
    const hidden = validItems.slice(maxPills);
    const displayItems = isExpanded ? validItems : shown;

    return (
      <div className={styles.termPills}>
        {displayItems.map((item, idx) => (
          <React.Fragment key={idx}>
            {renderItem(item)}
          </React.Fragment>
        ))}

        {hidden.length > 0 && !isExpanded && (
          <DefaultButton
            onClick={this.toggleExpand}
            className={styles.termPillButton}
          >
            View more +{hidden.length}
          </DefaultButton>
        )}

        {isExpanded && hidden.length > 0 && (
          <DefaultButton
            onClick={this.toggleExpand}
            className={styles.termPillButton}
          >
            View less
          </DefaultButton>
        )}
      </div>
    );
  }
}

export function renderArrayPills<T extends string | number>(
  arr: ReadonlyArray<T>,
  opts?: {
    maxPills?: number;
    format?: (v: T) => string;
    pillClassName?: string;
  }
): JSX.Element {
  return (
    <ArrayPillsContainer<T>
      items={Array.from(arr)}
      format={opts?.format}
      pillClassName={opts?.pillClassName}
      maxPills={opts?.maxPills}
    />
  );
}

export function StatusPill(value: string) {
  const statusStyles = getStatusViewProps(value);
  return (
    <span className={`${styles.statusPill} ${styles[statusStyles.className]}`}>
      {value}
    </span>
  );
}

export function renderObjectPills<T extends object = Record<string, any>>(
  items: T[],
  renderItem: (item: T) => JSX.Element,
  maxPills?: number
): JSX.Element {
  return <ObjectPillsContainer<T> items={items} renderItem={renderItem} maxPills={maxPills} />;
}
