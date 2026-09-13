import * as React from 'react';
import { DirectionalHint, IconButton, TooltipHost } from 'office-ui-fabric-react';

interface ITooltipProps extends React.SVGProps<SVGSVGElement> {
  content: string;
}

export default class Tooltip extends React.Component<ITooltipProps> {

  constructor(props) {
    super(props);
  }

  public async componentDidMount(): Promise<void> {
    //todo
  }

  public render(): React.ReactElement<ITooltipProps> {
    const calloutProps = { gapSpace: 0 };
    const hostStyles = { display: 'inline-block'  };
    const { content } = this.props;
    return (
      <TooltipHost
        calloutProps={calloutProps}
        tooltipProps={{
          onRenderContent: () => (
            <div className="content" dangerouslySetInnerHTML={{ __html: content }}></div>
          )
        }}
        directionalHint={DirectionalHint.rightCenter}
        style={hostStyles}>
        <IconButton iconProps={{ iconName: 'UnknownSolid' }} className='tooltipButton' />
      </TooltipHost>);
  }
}
