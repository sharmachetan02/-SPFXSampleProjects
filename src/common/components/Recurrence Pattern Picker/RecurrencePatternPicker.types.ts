// export type RecurrenceFrequency = 'Weekly' | 'Monthly' | 'Quarterly' | 'Half Annual' | 'Annual' | 'Multi Annual';
export enum RecurrenceFrequency {
    Weekly = 'Weekly',
    Monthly = 'Monthly',
    Quarterly = 'Quarterly',
    HalfAnnual = 'Half Annual',
    Annual = 'Annual',
    MultiAnnual = 'Multi Annual'
  }
export interface RecurrenceData {
  frequency: string;
  weekDay: string | null;
  month: string | null;
  day: number | null;
  interval: number | null;
  endDate: Date | null;
}

export interface RecurrencePatternPickerProps {
  availableFrequencies?: RecurrenceFrequency[];  // Frequencies to allow (default: all)
  recurrenceData?: Partial<RecurrenceData>;      // Initial values (optional)
  onChange: (data: RecurrenceData) => void;
}

export interface RecurrencePatternPickerState {
  frequency: RecurrenceFrequency;
  weekDay: string | null;
  month: string | null;
  day: number | null;
  interval: number;
  endDate: Date | null;
}
