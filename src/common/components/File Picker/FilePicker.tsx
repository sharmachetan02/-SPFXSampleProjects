import * as React from 'react';
import { IFilePickerProps } from './IFilePickerProps';
import { IFilePickerState } from './IFilePickerState';
import Dropzone from 'react-dropzone';
import { IconButton } from '@fluentui/react';
import styles from './FilePicker.module.scss';

export default class FilePicker extends React.Component<IFilePickerProps, IFilePickerState> {
  // private logoInput: HTMLInputElement;
  constructor(props: Readonly<IFilePickerProps>) {
    super(props);
    this.handleFileDrop = this.handleFileDrop.bind(this);
    // this.deleteAttachement = this.deleteAttachement.bind(this);
  }

  public async componentDidMount() {

    //empty
  }

  public render(): React.ReactElement<IFilePickerProps> {
    const {
      files
    } = this.props;
    return (
      <div>
        <div>
          <div>
            <Dropzone onDrop={this.handleFileDrop} multiple={this.props.multiple} disabled={this.props.disabled}
            >
              {({ getRootProps, getInputProps, isDragActive }) => (
                <div
                  {...getRootProps({
                    className: `dropzone ${this.props.disabled ? 'disabled' : isDragActive ? 'activeStyle' : ''}`
                  })}
                >
                  <input {...getInputProps()} />
                  {isDragActive ? (
                    <div>
                      <IconButton
                        iconProps={{ iconName: 'CloudUpload' }} />
                      Drop file here...
                    </div>
                  ) : (
                    <div>
                      <IconButton iconProps={{ iconName: 'CloudUpload' }} />{`Drag 'n' drop some file here, or click to select file`}</div>
                  )}
                </div>
              )
              }
            </Dropzone>
          </div>
          {files &&
            files.map((file) => (
              <div key={file.name} className={styles.attachement}>
                <div className={styles.attachementText}>{file.name}</div>
                {/* <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  onClick={() => this.deleteAttachement(file.name)} */}
                {/* /> */}
              </div>
            ))}

        </div>
        {

        }

      </div>
    );
  }

  private async handleFileDrop(droppedFiles: File[]) {
    const { onChange, multiple } = this.props;
    let { files } = this.props;
    if (droppedFiles.length > 0) {
      if (!multiple) {
        files = [];
        await Promise.all(droppedFiles.map(async (file: File) => {
          const filebuffer = await file.arrayBuffer();
          files.push({
            name: file.name,
            content: filebuffer
          });
        }));
        onChange(files);
      } else {

        droppedFiles.map((droppedFile) => {
          const fetchedFileIndex = files.findIndex((file) => {
            file.name === droppedFile.name;
          });
          if (fetchedFileIndex > -1) {
            files.splice(fetchedFileIndex, 1);
          }
        });

        await Promise.all(droppedFiles.map(async (file: File) => {
          const filebuffer = await file.arrayBuffer();
          files.push({
            name: file.name,
            content: filebuffer
          });
        }));
        onChange(files);
      }
    }




  }
  // private deleteAttachement(fileName: string) {
  //   const { files, onChange } = this.props;
  //   const fetchedFileIndex = files.findIndex((file) => file.name === fileName);
  //   if (fetchedFileIndex > -1) {
  //     files.splice(fetchedFileIndex, 1);
  //   }
  //   onChange(files);

  // }
}


