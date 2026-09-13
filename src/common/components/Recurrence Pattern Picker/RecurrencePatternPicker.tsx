import * as React from 'react';
import { Dropdown, IDropdownOption, DatePicker, Label } from 'office-ui-fabric-react';
import strings from 'GenericFormWebPartStrings';
import { RecurrenceData, RecurrenceFrequency, RecurrencePatternPickerProps, RecurrencePatternPickerState } from './RecurrencePatternPicker.types';
import { Month, WeekDay } from '../../models/Enums';
import styles from '../../../webparts/genericForm/components/Form.module.scss';
import Tooltip from '../../../common/components/Tooltip/Tooltip';
import { UtilHelper } from '../../helpers/Util';
import { datepickerstyles } from '../../../webparts/genericForm/components/Form.utility';

export default class RecurrencePatternPicker extends React.Component<RecurrencePatternPickerProps, RecurrencePatternPickerState> {
  constructor(props: RecurrencePatternPickerProps) {
    super(props);
    // Set default frequencies if not provided
    const allFreq: RecurrenceFrequency[] = Object.values(RecurrenceFrequency);
    const frequencies = props.availableFrequencies && props.availableFrequencies.length > 0 ?
      props.availableFrequencies : allFreq;
    // Determine initial frequency
    let initFreq: RecurrenceFrequency = frequencies[0];
    if (props.recurrenceData && props.recurrenceData.frequency) {
      const freqValue = props.recurrenceData.frequency;
      // Use provided initial frequency if it's in allowed list
      if (frequencies.indexOf(freqValue as RecurrenceFrequency) !== -1) {
        initFreq = freqValue as RecurrenceFrequency;
      }
    }
    // Initialize state fields with defaults or provided values
    const isWeekly = initFreq === RecurrenceFrequency.Weekly;
    const isAnnual = initFreq === RecurrenceFrequency.Annual || initFreq === RecurrenceFrequency.MultiAnnual;
    const isMonthBased = initFreq === RecurrenceFrequency.Monthly || initFreq === RecurrenceFrequency.Quarterly || initFreq === RecurrenceFrequency.HalfAnnual;
    const provided = props.recurrenceData || {};
    this.state = {
      frequency: initFreq,
      weekDay: isWeekly ?
        (provided.weekDay || strings.WeekdayMonday) : null,
      month: isAnnual ?
        (provided.month || strings.MonthJanuary) : null,
      day: !isMonthBased && !isAnnual ? null : (provided.day !== undefined ? provided.day : 1),
      interval: (() => {
        // Determine default interval based on frequency
        if (initFreq === RecurrenceFrequency.Quarterly) return provided.interval || 3;
        if (initFreq === RecurrenceFrequency.HalfAnnual) return provided.interval || 6;
        if (initFreq === RecurrenceFrequency.Annual) return 1;
        if (initFreq === RecurrenceFrequency.MultiAnnual) {
          const val = provided.interval || 2;
          return val < 2 ? 2 : val;
        }
        // Weekly or Monthly
        return provided.interval || 1;
      })(),
      endDate: provided.endDate || null
    };
    this.raiseOnChange();
  }

  // Helper to trigger onChange callback with current state mapped to RecurrenceData shape
  private raiseOnChange = (): void => {
    const freq = this.state.frequency;
    const data: RecurrenceData = {
      frequency: freq,
      weekDay: freq === RecurrenceFrequency.Weekly ? this.state.weekDay : null,
      month: (freq === RecurrenceFrequency.Annual || freq === RecurrenceFrequency.MultiAnnual) ? this.state.month : null,
      day: freq !== RecurrenceFrequency.Weekly ? this.state.day : null,
      interval: this.state.interval,
      endDate: this.state.endDate
    };
    this.props.onChange(data);
  };
  private formatDate = (date?: Date): string => {
    if (!date) return '';
    return UtilHelper.formatDate(date, 'LL', 'en-us');
  };
  private onFrequencyChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const newFreq = option.key as RecurrenceFrequency;
    const oldFreq = this.state.frequency;
    if (newFreq === oldFreq) return;
    const updates: Partial<RecurrencePatternPickerState> = {
      frequency: newFreq
    };
    // Adjust fields based on new frequency
    switch (newFreq) {
      case RecurrenceFrequency.Weekly:
        // Ensure a weekday is set
        updates.weekDay = strings.WeekdayMonday;
        // No day-of-month or month needed
        // Set interval to 1 week by default
        updates.interval = 1;
        break;
      case RecurrenceFrequency.Monthly:
        // For monthly, ensure day-of-month is set (preserve if already set)
        updates.day = 1;
        // Not using month or weekday
        // Set interval to 1 month
        updates.interval = 1;
        break;
      case RecurrenceFrequency.Quarterly:
        // Ensure day-of-month set
        updates.day = 1;
        // Interval fixed to 3 months
        updates.interval = 3;
        break;
      case RecurrenceFrequency.HalfAnnual:
        // Ensure day-of-month set
        updates.day = 1;
        // Interval fixed to 6 months
        updates.interval = 6;
        break;
      case RecurrenceFrequency.Annual:
        // Ensure month and day are set
        updates.month = strings.MonthJanuary;
        updates.day = 1;
        // Interval fixed to 1 year
        updates.interval = 1;
        break;
      case RecurrenceFrequency.MultiAnnual:
        // Ensure month and day are set
        updates.month = strings.MonthJanuary;
        updates.day = 1;
        // Interval default to 2 years if less
        updates.interval = (oldFreq === RecurrenceFrequency.MultiAnnual) ? (this.state.interval < 2 ? 2 : this.state.interval) : 2;
        break;
    }
    // If switching away from Weekly, we no longer need weekDay (preserve existing for potential reuse)
    if (newFreq !== RecurrenceFrequency.Weekly) {
      updates.weekDay = this.state.weekDay; // (Keep current or null; no change)
    }
    // If switching away from Annual, preserve month (no change needed)
    // Apply state updates and trigger callback
    this.setState(updates as RecurrencePatternPickerState, this.raiseOnChange);
  };

  private onWeekDayChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const newDay = option.key as string;
    this.setState({ weekDay: newDay }, this.raiseOnChange);
  };


  // private onIncrement = (value: string, event?: React.SyntheticEvent<HTMLElement>): string | void => {
  //   console.log(value);
  // };

  // /** Decrement the value (or return nothing to keep the previous value if invalid) */
  // private onDecrement = (value: string, event?: React.SyntheticEvent<HTMLElement>): string | void => {
  //   console.log(value);
  // };

  // /**
  //  * Clamp the value within the valid range (or return nothing to keep the previous value
  //  * if there's not valid numeric input)
  //  */
  // private onValidate = (value: string, event?: React.SyntheticEvent<HTMLElement>): string | void => {
  //   console.log(value);
  // };

  // private handleChange: React.FocusEventHandler<HTMLInputElement> = (event) => {
  //   this.onDayChange(event.currentTarget.value as string);
  // };
  private onDayChange = (event: React.SyntheticEvent<HTMLElement>, newValue?: string) => {
    if (!newValue) return;
    let val = parseInt(newValue, 10);
    if (isNaN(val)) {
      return; // ignore invalid input
    }
    // Clamp day between 1 and 31
    if (val < 1) val = 1;
    if (val > 31) val = 31;
    this.setState({ day: val }, this.raiseOnChange);
  };
  // private Test = (newValue:string) => {
  //   if (!newValue) return;
  //  console.log(newValue);
  //  const val = parseInt(newValue, 10);
  //  this.setState({ day: val }, this.raiseOnChange);
  // };
  private onIntervalChange = (event: React.SyntheticEvent<HTMLElement>, newValue?: string) => {
    if (!newValue) return;
    let val = parseInt(newValue, 10);
    if (isNaN(val)) {
      return;
    }
    // Enforce minimum interval based on frequency
    if (val < 1) val = 1;
    if (this.state.frequency === RecurrenceFrequency.MultiAnnual) {
      if (val < 2) val = 2;  // Multi-annual interval must be at least 2 years
    }
    this.setState({ interval: val }, this.raiseOnChange);
  };

  private onEndDateSelect = (date: Date | null | undefined): void => {
    // Allow clearing the date (null) or selecting a new date
    this.setState({ endDate: date || null }, this.raiseOnChange);
  };

  private onMonthChange = (_e, option) => {
    if (option) {
      this.setState({ month: option.key as string }, this.raiseOnChange);
    }
  };

  public render(): JSX.Element {
    const freq = this.state.frequency;
    // Build frequency dropdown options from allowed frequencies (localized text)
    const frequencyOptions: IDropdownOption[] = [];
    const freqList = this.props.availableFrequencies && this.props.availableFrequencies.length > 0 ? this.props.availableFrequencies :
      [RecurrenceFrequency.Weekly, RecurrenceFrequency.Monthly, RecurrenceFrequency.Quarterly, RecurrenceFrequency.HalfAnnual, RecurrenceFrequency.Annual, RecurrenceFrequency.MultiAnnual];
    freqList.forEach((f) => {
      switch (f) {
        case RecurrenceFrequency.Weekly: frequencyOptions.push({ key: RecurrenceFrequency.Weekly, text: strings.FrequencyWeekly }); break;
        case RecurrenceFrequency.Monthly: frequencyOptions.push({ key: RecurrenceFrequency.Monthly, text: strings.FrequencyMonthly }); break;
        case RecurrenceFrequency.Quarterly: frequencyOptions.push({ key: RecurrenceFrequency.Quarterly, text: strings.FrequencyQuarterly }); break;
        case RecurrenceFrequency.HalfAnnual: frequencyOptions.push({ key: RecurrenceFrequency.HalfAnnual, text: strings.FrequencyHalfAnnual }); break;
        case RecurrenceFrequency.Annual: frequencyOptions.push({ key: RecurrenceFrequency.Annual, text: strings.FrequencyAnnual }); break;
        case RecurrenceFrequency.MultiAnnual: frequencyOptions.push({ key: RecurrenceFrequency.MultiAnnual, text: strings.FrequencyMultiAnnual }); break;
      }
    });
    // Weekday options (Monday–Sunday)
    const weekdayOptions: IDropdownOption[] = [
      { key: WeekDay.Monday, text: strings.WeekdayMonday },
      { key: WeekDay.Tuesday, text: strings.WeekdayTuesday },
      { key: WeekDay.Wednesday, text: strings.WeekdayWednesday },
      { key: WeekDay.Thursday, text: strings.WeekdayThursday },
      { key: WeekDay.Friday, text: strings.WeekdayFriday },
      { key: WeekDay.Saturday, text: strings.WeekdaySaturday },
      { key: WeekDay.Sunday, text: strings.WeekdaySunday }
    ];
    // Month options (January–December)
    const monthOptions: IDropdownOption[] = [
      { key: Month.January, text: strings.MonthJanuary },
      { key: Month.February, text: strings.MonthFebruary },
      { key: Month.March, text: strings.MonthMarch },
      { key: Month.April, text: strings.MonthApril },
      { key: Month.May, text: strings.MonthMay },
      { key: Month.June, text: strings.MonthJune },
      { key: Month.July, text: strings.MonthJuly },
      { key: Month.August, text: strings.MonthAugust },
      { key: Month.September, text: strings.MonthSeptember },
      { key: Month.October, text: strings.MonthOctober },
      { key: Month.November, text: strings.MonthNovember },
      { key: Month.December, text: strings.MonthDecember }
    ];
    // Determine interval spin configuration
    const intervalMin = (freq === RecurrenceFrequency.MultiAnnual) ? 2 : 1;
    let intervalSuffix = '';
    if (freq === RecurrenceFrequency.Weekly) {
      intervalSuffix = strings.IntervalSuffixWeeks;     // "week(s)"
    } else if (freq === RecurrenceFrequency.Monthly || freq === RecurrenceFrequency.Quarterly || freq === RecurrenceFrequency.HalfAnnual) {
      intervalSuffix = strings.IntervalSuffixMonths;    // "month(s)"
    } else if (freq === RecurrenceFrequency.Annual || freq === RecurrenceFrequency.MultiAnnual) {
      intervalSuffix = strings.IntervalSuffixYears;     // "year(s)"
    }
    const intervalDisabled = (freq === RecurrenceFrequency.Quarterly || freq === RecurrenceFrequency.HalfAnnual || freq === RecurrenceFrequency.Annual);
    return (
      <div className="recurrencePatternPicker">
        {/* Frequency selection */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.RecurrenceFrequencyLabel}</Label>
            </div>
            <Tooltip content={strings.RecurrenceTooltipFrequency} ></Tooltip>
          </div>
          <Dropdown
            selectedKey={this.state.frequency}
            options={frequencyOptions}
            onChange={this.onFrequencyChange} />
        </div>

        {/* Weekly: weekday + interval in weeks */}
        {freq === RecurrenceFrequency.Weekly && (
          <>
            <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.RecurrencePattern}</Label>
                </div>
                <Tooltip content={strings.RecurrenceTooltipPattern} ></Tooltip>
              </div>
              <div className={styles.recurrencePatternContainer}>
                <div className={styles.recurrencePatternItem}>
                  <Dropdown
                    selectedKey={this.state.weekDay || undefined}
                    options={weekdayOptions}
                    onChange={this.onWeekDayChange}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {strings.RecurrenceIntervalPrefix}
                </div>
                <div className={styles.recurrencePatternItem}>
                  {/* <SpinButton
                    min={intervalMin}
                    step={1}
                    disabled={intervalDisabled}
                    value={String(this.state.interval)}
                    styles={{
                      root: { width: 50 },
                      input: { width: 40 }
                    }}
                    onChange={this.onIntervalChange}
                  /> */}
                  <input
                    type="number"
                    min={intervalMin}
                    disabled={intervalDisabled}
                    step={1}
                    value={String(this.state.interval)}
                    onChange={(e) => this.onIntervalChange(null, e.target.value)}
                    className={styles.numberInput}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {intervalSuffix}
                </div>
              </div>

            </div>
          </>
        )
        }

        {/* Monthly/Quarterly/Half Annual: day of month + interval in months */}
        {(freq === RecurrenceFrequency.Monthly || freq === RecurrenceFrequency.Quarterly || freq === RecurrenceFrequency.HalfAnnual) && (
          <>
            <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.RecurrencePattern}</Label>
                </div>
                <Tooltip content={strings.RecurrenceTooltipPattern} ></Tooltip>
              </div>
              <div className={styles.recurrencePatternContainer}>
                <div className={styles.recurrencePatternItem}>
                  Day
                </div>
                <div className={styles.recurrencePatternItem}>
                  {/* <SpinButton
                    min={1}
                    max={31}
                    step={1}
                    value={String(this.state.day || 1)}
                    styles={{
                      root: { width: 50 },
                      input: { width: 40 },
                    }}

                    // onIncrement={(value) => {
                    //   const v = parseInt(value || "0", 10) + 1;
                    //   this.Test(v.toString());
                    // }}
                    // onDecrement={(value) => {
                    //   const v = parseInt(value || "0", 10) - 1;
                    //   this.Test(v.toString());
                    // }}
                    // onChange={this.handleChange}
                    // onBlur={this.handleChange}
                    onChange={this.onDayChange}
                  // inputProps={{ onChange: this.handleChange }}
                  /> */}
                  <input
                    type="number"
                    min={1}
                    max={31}
                    step={1}
                    value={String(this.state.day || 1)}
                    onChange={(e) => this.onDayChange(null, e.target.value)}
                    className={styles.numberInput}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {strings.RecurrenceIntervalPrefix}
                </div>
                <div className={styles.recurrencePatternItem}>
                  {/* <SpinButton
                    min={intervalMin}
                    step={1}
                    disabled={intervalDisabled}
                    value={String(this.state.interval)}
                    styles={{
                      root: { width: 50 },
                      input: { width: 40 }
                    }}
                    onChange={this.onIntervalChange}
                  /> */}
                  <input
                    type="number"
                    min={intervalMin}
                    disabled={intervalDisabled}
                    step={1}
                    value={String(this.state.interval)}
                    onChange={(e) => this.onIntervalChange(null, e.target.value)}
                    className={styles.numberInput}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {intervalSuffix}
                </div>
              </div>

            </div>
          </>
        )
        }

        {/* Annual/Multi Annual: month + day + interval in years */}
        {(freq === RecurrenceFrequency.Annual || freq === RecurrenceFrequency.MultiAnnual) && (
          <>
            <div className={styles.field}>
              <div className={styles.fieldLabel}>
                <div className={styles.fieldLabelContainer}>
                  <Label>{strings.RecurrencePattern}</Label>
                </div>
                <Tooltip content={strings.RecurrenceTooltipPattern} ></Tooltip>
              </div>
              <div className={styles.recurrencePatternContainer}>
                <div className={styles.recurrencePatternItem}>
                  Every
                </div>
                <div className={styles.recurrencePatternItem}>
                  <Dropdown
                    selectedKey={this.state.month || undefined}
                    options={monthOptions}
                    onChange={this.onMonthChange}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {/* <SpinButton
                    min={1}
                    max={31}
                    step={1}
                    value={String(this.state.day || '')}
                    styles={{
                      root: { width: 50 },
                      input: { width: 40 }
                    }}
                    // onIncrement={(value) => {
                    //   const v = parseInt(value || "0", 10) + 1;
                    //   this.Test(v.toString());
                    // }}
                    // onDecrement={(value) => {
                    //   const v = parseInt(value || "0", 10) - 1;
                    //   this.Test(v.toString());
                    // }}
                    // onChange={this.handleChange}
                    // onBlur={this.handleChange}
                    // inputProps={{ onChange: this.handleChange }}
                    // onValidate={this.onValidate}
                    // onIncrement={this.onIncrement}
                    // onDecrement={this.onDecrement}
                    onChange={this.onDayChange}
                  /> */}
                  <input
                    type="number"
                    min={1}
                    max={31}
                    step={1}
                    value={String(this.state.day || 1)}
                    onChange={(e) => this.onDayChange(null, e.target.value)}
                    className={styles.numberInput}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {strings.RecurrenceIntervalPrefix}
                </div>
                <div className={styles.recurrencePatternItem}>
                  {/* <SpinButton
                    min={intervalMin}
                    step={1}
                    disabled={intervalDisabled}
                    value={String(this.state.interval)}
                    styles={{
                      root: { width: 50 },
                      input: { width: 40 }
                    }}
                    onChange={this.onIntervalChange}
                  /> */}
                  <input
                    type="number"
                    min={intervalMin}
                    disabled={intervalDisabled}
                    step={1}
                    value={String(this.state.interval)}
                    onChange={(e) => this.onIntervalChange(null, e.target.value)}
                    className={styles.numberInput}
                  />
                </div>
                <div className={styles.recurrencePatternItem}>
                  {intervalSuffix}
                </div>
              </div>
            </div>
          </>
        )
        }
        {/* Optional end date */}
        {/* Interval*/}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.EndDateFieldLabel}</Label>
            </div>
            <Tooltip content={strings.RecurrenceTooltipEndDate} ></Tooltip>
          </div>
          <DatePicker
            value={this.state.endDate}
            onSelectDate={this.onEndDateSelect}
            formatDate={this.formatDate}
            allowTextInput
            styles={datepickerstyles}
          />
        </div>
      </div>
    );
  }
  // eslint-disable-next-line camelcase
  // UNSAFE_componentWillReceiveProps(nextProps) {
  //   // You don't have to do this check first, but it can help prevent an unneeded render
  //   if (nextProps.recurrenceData.day !== this.state.day) {
  //     this.setState({ day: nextProps.recurrenceData.day });
  //   }
  //   if (nextProps.recurrenceData.endDate !== this.state.endDate) {
  //     this.setState({ endDate: nextProps.recurrenceData.endDate });
  //   }
  //   if (nextProps.recurrenceData.frequency !== this.state.frequency) {
  //     this.setState({ frequency: nextProps.recurrenceData.frequency });
  //   }
  //   if (nextProps.recurrenceData.interval !== this.state.interval) {
  //     this.setState({ interval: nextProps.recurrenceData.interval });
  //   }
  //   if (nextProps.recurrenceData.month !== this.state.month) {
  //     this.setState({ month: nextProps.recurrenceData.month });
  //   }
  //   if (nextProps.recurrenceData.weekDay !== this.state.weekDay) {
  //     this.setState({ weekDay: nextProps.recurrenceData.weekDay });
  //   }
  // }
}
