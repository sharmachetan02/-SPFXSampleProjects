import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface ILoadingSpinnerProps {
  label?: string;
  size?: SpinnerSize;
  className?: string;
}

const LoadingSpinner: React.FunctionComponent<ILoadingSpinnerProps> = ({
  label = 'Loading...',
  size = SpinnerSize.medium,
  className = ''
}) => (
  <div className={className}>
    <Spinner label={label} size={size} />
  </div>
);

export default LoadingSpinner;
