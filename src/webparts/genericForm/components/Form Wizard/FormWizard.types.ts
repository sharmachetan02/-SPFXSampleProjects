import { IItemAddResult } from '@pnp/sp/items';
import { FormOrigin, IFormConfig, IMessageBanner } from '../Form.types';
import { IBaseItem } from '../../../../common/models/IBusiness';

export interface IFormWizardProps {
  validateForm: () => boolean | Promise<boolean>
  processEdit: () => void;
  processCreation: () => Promise<IItemAddResult>;
  getSummary: () => JSX.Element;
  getDuplicateItems: () => Promise<IBaseItem[]>;
  isEdit: boolean;
  isFormReady: boolean;
  formConfig: IFormConfig;
  allowDuplicateItems: boolean;
  formOrigin:FormOrigin;
  callback?: (string?) => void;
}

export interface IFormWizardState {
  formStepHidden: boolean;
  creationConfirmationStepHidden: boolean;
  editConfirmationStepHidden: boolean;
  redirectionStepHidden: boolean;
  isFormProcessing: boolean;
  isConfirmButtonDisabled: boolean;
  suggestedItems: IBaseItem[];
  messageBanner: IMessageBanner;
}
