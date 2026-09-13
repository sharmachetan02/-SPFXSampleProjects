import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  DialogFooter,
  Spinner,
  Stack,
  SpinnerSize,
  Icon
} from '@fluentui/react';
import { IFormWizardProps, IFormWizardState } from './FormWizard.types';
import strings from 'GenericFormWebPartStrings';
import { Config } from '../../../../common/config/Config';
import { Consts } from '../../../../common/consts/Consts';
import styles from '../Form.module.scss';
import { navigateToDatasheet } from '../Form.utility';
import { FormOrigin, IMessageBanner } from '../Form.types';
import { MessageBanner } from '../../../../common/components/Message Banner/MessageBanner';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
export default class FormWizard extends React.Component<IFormWizardProps, IFormWizardState> {
  constructor(props: Readonly<IFormWizardProps>) {
    super(props);
    const messageBanner: IMessageBanner = {
      message: '',
      type: 'info',
      visible: false
    };
    this.state = {
      formStepHidden: false,
      creationConfirmationStepHidden: true,
      editConfirmationStepHidden: true,
      redirectionStepHidden: true,
      isFormProcessing: false,
      isConfirmButtonDisabled: false,
      suggestedItems: [],
      messageBanner
    };
  }

  public async componentDidMount(): Promise<void> {
    //empty
  }

  public render(): React.ReactElement<IFormWizardProps> {
    const {
      formStepHidden,
      creationConfirmationStepHidden,
      editConfirmationStepHidden,
      redirectionStepHidden,
      isFormProcessing,
      isConfirmButtonDisabled,
      messageBanner
    } = this.state;
    const {
      getSummary,
      isFormReady
    } = this.props;
    return (
      <div>
        <MessageBanner
          message={messageBanner.message}
          type={messageBanner.type}
          visible={messageBanner.visible}
          onDismiss={() => this.setState((prevState) => ({
            messageBanner: {
              ...prevState.messageBanner,
              visible: false
            }
          }))}
        />
        <div className="ms-Grid">
          {/* pop up form */}
          {
            !formStepHidden &&
            <div className={styles.wizardStep}>
              {this.props.children}
              {
                isFormReady &&
                <DialogFooter>
                  <PrimaryButton
                    text="Submit"
                    disabled={false}
                    onClick={this.goToNextStep}
                  />
                  <DefaultButton
                    text="Cancel"
                    disabled={false}
                    onClick={this.cancelForm}
                  />
                </DialogFooter>
              }

            </div>
          }

          {/* pop up creation */}
          {!creationConfirmationStepHidden &&
            <div className={styles.wizardStep}>
              {this.getSuggestions()}
              <p>
                {strings.FormCreationSummaryText}
              </p>
              {getSummary()}
              <p>
                {strings.FormCreationConfirmation}
              </p>
              <DialogFooter>
                <PrimaryButton
                  disabled={isConfirmButtonDisabled}
                  onClick={this.handleCreationConfirmation}>
                  {isFormProcessing ? (
                    <Stack horizontal verticalAlign="center">
                      <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                      <span style={{ marginLeft: 8 }}>Processing</span>
                    </Stack>
                  ) : (
                    'Confirm'
                  )}
                </PrimaryButton>
                <DefaultButton
                  onClick={this.goToPreviousStep}
                  text="Go Back"
                />
              </DialogFooter>
            </div>
          }

          {/* pop up edit */}
          {!editConfirmationStepHidden &&
            <div className={styles.wizardStep}>
              <p>
                {strings.FormEditSummaryText}
              </p>
              {getSummary()}
              <p>
                {strings.FormEditConfirmation}
              </p>

              <DialogFooter>
                <PrimaryButton
                  disabled={isConfirmButtonDisabled}
                  onClick={this.handleEditConfirmation}>
                  {isFormProcessing ? (
                    <Stack horizontal verticalAlign="center">
                      <Spinner size={SpinnerSize.xSmall} labelPosition="right" />
                      <span style={{ marginLeft: 8 }}>Processing</span>
                    </Stack>
                  ) : (
                    'Confirm'
                  )}
                </PrimaryButton>
                <DefaultButton
                  onClick={this.goToPreviousStep}
                  text="Go Back"
                />
              </DialogFooter>
            </div>
          }

          {/* pop up redirection*/}
          {!redirectionStepHidden &&
            <div className={styles.wizardStep}>
              <p>{strings.FormRedirectionText}</p>
              {getSummary()}
              <DialogFooter>
                <PrimaryButton
                  onClick={
                    this.handleRedirection
                  }
                  text="Go To Datasheet"
                />
                <DefaultButton
                  onClick={this.goToPreviousStep}
                  text="Go Back" />
              </DialogFooter>
            </div>}
        </div>
      </div>
    );
  }

  private setStep(step: 'form' | 'create' | 'edit' | 'redirect') {
    this.setState({
      formStepHidden: step !== 'form',
      creationConfirmationStepHidden: step !== 'create',
      editConfirmationStepHidden: step !== 'edit',
      redirectionStepHidden: step !== 'redirect'
    });
  }
  private scrollToTop = () => {
    const { isEdit, formOrigin } = this.props;
    switch (formOrigin) {
      case (FormOrigin.Datasheet): {
        document.querySelector('*[class*="ms-Modal-scrollableContent"]').scrollTop = 0;
        break;
      }
      default: {
        if (!isEdit)
          document.querySelector('[data-automation-id="contentScrollRegion"]').scrollTop = 0;
        else
          document.querySelector('*[class*="ms-Modal-scrollableContent"]').scrollTop = 0;
      }
    }

  };
  private handleCreationConfirmation = async () => {
    try {
      const { formConfig, formOrigin, processCreation, callback } = this.props;
      const { isFormProcessing } = this.state;
      if (isFormProcessing) return;

      this.setState({ isFormProcessing: true });
      const item = await processCreation();
      const messageBanner: IMessageBanner = {
        message: strings.FormCreationSuccess,
        type: 'success',
        visible: true
      };
      this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
      this.scrollToTop();
      switch (formOrigin) {
        case (FormOrigin.Datasheet): {
          setTimeout(() => { callback(window.location); }, 3000);
          break;
        }
        default: {
          setTimeout(() => { navigateToDatasheet(formConfig?.type, item?.data.Id); }, 3000);
        }
      }

    } catch (error) {
      this.scrollToTop();
      console.error(strings.FormCreationError, error);
      const messageBanner: IMessageBanner = {
        message: strings.FormCreationFailure,
        type: 'error',
        visible: true
      };
      this.setState({ isFormProcessing: false, messageBanner });
      return null;
    }
  };

  private handleEditConfirmation = async () => {
    try {

      const { processEdit, callback } = this.props;
      const { isFormProcessing } = this.state;
      if (isFormProcessing) return;
      this.setState({ isFormProcessing: true });
      await processEdit();
      const messageBanner: IMessageBanner = {
        message: strings.FormEditSuccess,
        type: 'success',
        visible: true
      };
      this.setState({ isFormProcessing: false, isConfirmButtonDisabled: true, messageBanner });
      this.scrollToTop();
      setTimeout(() => { callback(window.location); }, 3000);
    } catch (error) {
      this.scrollToTop();
      console.error(strings.FormEditError, error);
      const messageBanner: IMessageBanner = {
        message: strings.FormEditFailure,
        type: 'error',
        visible: true
      };
      this.setState({ isFormProcessing: false, messageBanner });
      return null;
    }
  };

  private getSuggestions = (): JSX.Element => {
    const { allowDuplicateItems, formConfig } = this.props;
    const { suggestedItems } = this.state;

    if (!allowDuplicateItems) {
      return null;
    }

    if (suggestedItems.length === 0) {
      return null;
    }
    const html = suggestedItems.map((item) => {
      const href = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${formConfig.type}&itemId=${item.Id}`;
      return (
        <li key={item.Id} className={styles.suggestedItem}>
          <p>
            <b>
              <a className="ms-Link" data-interception='off' target="_blank" rel="noreferrer" href={href}>
                {item.Title}
              </a>
            </b>
          </p>
        </li>
      );
    });
    return (
      <div>
        <div className="ms-MessageBar ms-MessageBar--severeWarning">
          <div className="ms-MessageBar-content">
            <div className="ms-MessageBar-icon">
              <Icon iconName='Error' className={styles.alertIcon} />
            </div>
            <div className="ms-MessageBar-text">
              <div className="ms-MessageBar-title">
                {strings.FormSuggestionText}
              </div>
              <ul>{html}</ul>
            </div>
          </div>
        </div>
      </div>
    );
  };

  private goToPreviousStep = () => {
    this.setStep('form');
  };

  private goToNextStep = async () => {
    const {
      isEdit,
      allowDuplicateItems,
      validateForm
    } = this.props;
    let {
      suggestedItems
    } = this.state;
    try {
      const isValid = await Promise.resolve(validateForm());
      if (isValid) {
        this.scrollToTop();
        const { getDuplicateItems } = this.props;
        if (getDuplicateItems)
          suggestedItems = await getDuplicateItems();
        if (isEdit) {
          // edit
          if (suggestedItems.length > 0 && !allowDuplicateItems) {
            this.setStep('redirect');
            this.setState({ suggestedItems });
          } else {
            this.setStep('edit');
          }
        } else if (suggestedItems.length === 0) {
          // creation
          this.setStep('create');
        } else if (suggestedItems.length > 0 && allowDuplicateItems) {
          // creation
          this.setStep('create');
          this.setState({ suggestedItems });
        } else if (suggestedItems.length > 0 && !allowDuplicateItems) {
          // redirection
          this.setStep('redirect');
          this.setState({ suggestedItems });
        }
      }

    } catch (error) {
      console.log(error);
    }
  };

  private handleRedirection = () => {
    const { formConfig } = this.props;
    const { suggestedItems } = this.state;
    const duplicatedItem = suggestedItems[0];
    navigateToDatasheet(formConfig?.type, duplicatedItem?.Id);
  };

  private cancelForm = () => {
    const {
      isEdit,
      callback,
    } = this.props;
    if (isEdit)
      callback();
    else
      callback(Config.SITE.ROOT_WEB_URL);
  };
}
