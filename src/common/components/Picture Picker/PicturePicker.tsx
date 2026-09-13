import * as React from 'react';
import { IPicturePickerProps, IPicturePickerState } from './PicturePicker.types';
import Dropzone from 'react-dropzone';
import {
  IconButton,
} from 'office-ui-fabric-react/lib/index';
import '@pnp/sp/files';
export default class PicturePicker extends React.Component<
  IPicturePickerProps,
  IPicturePickerState
> {

  // private logoInput: HTMLInputElement;
  constructor(props: Readonly<IPicturePickerProps>) {
    super(props);
    this.state = {
      maxSizeExceeded: false,
      url: ''
    };
    this.savePicture = this.savePicture.bind(this);


  }

  public async componentDidMount() {
    //empty
  }

  public render(): React.ReactElement<IPicturePickerProps> {
    const {
      maxSizeExceeded,
      url
    } = this.state;
    return (
      <div>
        <div>
          <div className='picture-picker'>
            <Dropzone onDrop={this.savePicture} multiple={false} accept={{ 'mime/type': ['.gif', '.jpg', '.jpeg', '.bmp', '.dib', '.tif', '.tiff', '.ico', '.png', '.jxr', '.svg'] }}>
              {({ getRootProps, getInputProps, isDragActive }) =>
              (
                <div
                  {...getRootProps({
                    className: `dropzone ${isDragActive ? 'activeStyle' : ''
                      }`
                  })}
                >
                  <input {...getInputProps()} />
                  {isDragActive ? (
                    <div>
                      <IconButton
                        iconProps={{ iconName: 'CloudUpload' }}
                      />
                      Drop file here...
                    </div>
                  ) : (
                    <div>
                      <IconButton
                        iconProps={{ iconName: 'CloudUpload' }}
                      />
                      Drag &apos;n&apos; drop a picture here, or click to select
                      a picture
                    </div>
                  )}
                </div>
              )
              }
            </Dropzone>
          </div>
          {url && <div className='picture-preview'>
            <img src={url} />
          </div>}

        </div>
        {
          maxSizeExceeded &&
          <div>
            <p className='ms-TextField-errorMessage errorMessage '>
              <span>
                Max Size 2 Mo
              </span>
            </p>
          </div>
        }

      </div>
    );
  }

  private async savePicture(files: File[]) {
    const { onChange, pnpService } = this.props;
    let { maxSizeExceeded } = this.state;
    const file = files[0];
    let url = '';
    if (file) {
      if (file.size > 2097152) {
        maxSizeExceeded = true;
        this.setState({ maxSizeExceeded });
        onChange('');
        return;
      } else {
        maxSizeExceeded = false;
        this.setState({ maxSizeExceeded });
      }
      //url = await this.getBase64(file);
      //this.setState({ url });

      const fileresult = await pnpService.getFolder(pnpService.getContext().pageContext.web.serverRelativeUrl + '/Pictures').files.addChunked(
        this.getRandomFileName(), file, () => { console.log('progress'); }, true);
      //const item = await fileresult.file.getItem();
      // await item.update({
      //   Tsm_PictureCategory: category
      // });
      url = document.location.origin + fileresult.data.ServerRelativeUrl;
      this.setState({ url });
      onChange(url);
    }

  }
  // private async getBase64(file):Promise<string>{
  //   return new Promise(resolve => {
  //     let fileInfo;
  //     let baseURL = '';
  //     // Make new FileReader
  //     let reader = new FileReader();

  //     // Convert the file to base64 text
  //     reader.readAsDataURL(file);

  //     // on reader load somthing...
  //     reader.onload = () => {
  //       // Make a fileInfo Object
  //       console.log('Called', reader);
  //       baseURL = reader.result as string;
  //       console.log(baseURL);
  //       return resolve(baseURL);
  //     };
  //     console.log(fileInfo);
  //   });
  // }
  private getRandomFileName() {
    const timestamp = new Date()?.toISOString().replace(/[-:.]/g, '');
    const random = ('' + Math.random()).substring(2, 8);
    const randomNumber = timestamp + random;
    return randomNumber;
  }
}
