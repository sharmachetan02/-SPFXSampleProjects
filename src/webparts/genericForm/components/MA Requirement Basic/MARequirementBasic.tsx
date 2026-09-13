import * as React from 'react';
import { Dropdown, Icon, Label, TextField } from '@fluentui/react';
import strings from 'GenericFormWebPartStrings';
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import styles from '../Form.module.scss';
import { MARequirementBasicState, MARequirementBasicProps, ICategorizedDropdownOption } from './MARequirementBasic.types';
import { Combobox } from 'react-widgets/cjs';
import MultiselectWrapper from '../../../../common/components/MultiselectWrapper/MultiselectWrapper';
import Tooltip from '../../../../common/components/Tooltip/Tooltip';
import { RichText } from '../../../../common/components/RichText/RichText';
import { ChangeHandler } from 'react-widgets/cjs/List';
import { MARequirementResponsibleParty, Vertical } from '../../../../common/models/Enums';

export default class MARequirementBasic extends React.Component<MARequirementBasicProps, MARequirementBasicState> {
  constructor(props: MARequirementBasicProps) {
    super(props);
    const maRequirementBasicData = this.props.maRequirementBasicData;
    this.state = {
      maVertical: maRequirementBasicData.maVertical || [],
      maVerticalChoices: maRequirementBasicData.maVerticalChoices || [],
      snps: maRequirementBasicData.snps || [],
      snpsChoices: maRequirementBasicData.snpsChoices || [],
      geoLeo: maRequirementBasicData.geoLeo || '',
      geoLeoChoices: maRequirementBasicData.geoLeoChoices || [],
      comments: maRequirementBasicData.comments || '',
      name: maRequirementBasicData.name || '',
      type: maRequirementBasicData.type || '',
      typeChoices: maRequirementBasicData.typeChoices || [],
      requirementType: maRequirementBasicData.requirementType || '',
      requirementTypeChoices: maRequirementBasicData.requirementTypeChoices || [],
      summary: maRequirementBasicData.summary || '',
      markets: maRequirementBasicData.markets || [],
      marketsChoices: maRequirementBasicData.marketsChoices || [],
      responsibleParty: maRequirementBasicData.responsibleParty || '',
      responsiblePartyChoices: maRequirementBasicData.responsiblePartyChoices || [],
      responsiblePartyItem: maRequirementBasicData.responsiblePartyItem || null,
      responsiblePartyItemChoices: maRequirementBasicData.responsiblePartyItemChoices || [],
      eutelsatOwners: maRequirementBasicData.eutelsatOwners || [],
      eutelsatOwnersChoices: maRequirementBasicData.eutelsatOwnersChoices || [],
      responsiblePartyContacts: maRequirementBasicData.responsiblePartyContacts || [],
      responsiblePartyContactsChoices: maRequirementBasicData.responsiblePartyContactsChoices || [],
      synthesisStatus: maRequirementBasicData.synthesisStatus || '',
      synthesisStatusChoices: maRequirementBasicData.synthesisStatusChoices || [],
      ragStatus: maRequirementBasicData.ragStatus || '',
      ragStatusChoices: maRequirementBasicData.ragStatusChoices || []
    };
    this.raiseOnChange();
  }

  // Helper to trigger onChange callback with current state mapped to RecurrenceData shape
  private raiseOnChange = (): void => {
    const maRequirementBasicData = this.state;

    this.props.onChange(maRequirementBasicData);
  };


  // Individual handlers
  private onChangeName = (event, name?: string): void => {
    this.setState({ name }, this.raiseOnChange);
  };

  private onChangeType = (event, option?: IDropdownOption): void => {
    if (option) this.setState({ type: option?.key?.toString() }, this.raiseOnChange);
  };

  private onChangeRequirementType: ChangeHandler<ICategorizedDropdownOption> = (value: ICategorizedDropdownOption) => {
    if (value) {
      const maVertical: IDropdownOption[] = [];
      value?.maVerticals?.forEach((item) => {
        maVertical.push({ key: item, text: item });
      });
      this.setState({ requirementType: value?.key?.toString(), maVertical }, this.raiseOnChange);
    }
  };

  private onChangeSummary = (summary: string): string => {
    this.setState({ summary }, this.raiseOnChange);
    return summary;
  };

  private onChangeMarket = (markets: IDropdownOption[]): void => {
    this.setState({ markets }, this.raiseOnChange);
  };

  private onChangeresponsibleParty = (event, option?: IDropdownOption): void => {
    if (option) this.setState({ responsibleParty: option?.key?.toString(), responsiblePartyItem: null, responsiblePartyContacts: [] }, this.raiseOnChange);
  };

  private onChangeChargeByItem = (responsiblePartyItem: IDropdownOption): void => {
    this.setState({ responsiblePartyItem, responsiblePartyContacts: [] }, this.raiseOnChange);
  };

  private onChangeEutelsatOwners = (eutelsatOwners: IDropdownOption[]): void => {
    this.setState({ eutelsatOwners }, this.raiseOnChange);
  };

  private onChangeMAVerticals = (maVertical: IDropdownOption[]): void => {
    if (!maVertical.find((v) => v.text === Vertical.SNP)) {
      this.setState({ maVertical, snps: [] }, this.raiseOnChange);
    } else {
      this.setState({ maVertical, }, this.raiseOnChange);
    }
  };

  private onChangeContacts = (responsiblePartyContacts: IDropdownOption[]): void => {
    this.setState({ responsiblePartyContacts }, this.raiseOnChange);
  };


  private onChangeSynthesis = (event, option?: IDropdownOption): void => {
    if (option) this.setState({ synthesisStatus: option?.key?.toString() }, this.raiseOnChange);
  };

  private onChangeRagStatus = (event, option?: IDropdownOption): void => {
    if (option) this.setState({ ragStatus: option.text }, this.raiseOnChange);
  };

  private onChangeSnps = (snps: IDropdownOption[]): void => {
    if (snps) this.setState({ snps }, this.raiseOnChange);
  };

  private onChangeGeoLeo = (event, option: IDropdownOption): void => {
    if (option) this.setState({ geoLeo: option.text }, this.raiseOnChange);
  };
  private onChangeComments = (comments: string): string => {
    this.setState({ comments }, this.raiseOnChange);
    return comments;
  };

  public render(): JSX.Element {
    const {
      name,
      type,
      typeChoices,
      summary,
      markets,
      marketsChoices,
      responsibleParty,
      responsiblePartyChoices,
      responsiblePartyItem,
      responsiblePartyItemChoices,
      eutelsatOwners,
      eutelsatOwnersChoices,
      responsiblePartyContacts,
      responsiblePartyContactsChoices,
      ragStatus,
      ragStatusChoices,
      synthesisStatus,
      synthesisStatusChoices,
      maVertical,
      maVerticalChoices,
      snps,
      snpsChoices,
      geoLeo,
      geoLeoChoices,
      comments,
      requirementTypeChoices,
      requirementType
    } = this.state;
    const { errors } = this.props;
    const selectedRequirementType = (requirementTypeChoices?.find((item) => item.key.toString() === requirementType) || []);
    return (
      <>
        {/* Name */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelName}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipName} ></Tooltip>
          </div>
          <TextField
            value={name}
            onChange={this.onChangeName}
          />
          {errors.name &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.name}
              </span>
            </div>
          }
        </div>
        {/* Type */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelType}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipType} ></Tooltip>
          </div>
          <Dropdown
            disabled={true}
            selectedKey={type}
            onChange={this.onChangeType}
            options={typeChoices} />
        </div>
        {/* Summary */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelSummary}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipSummary} ></Tooltip>
          </div>
          <RichText isEditMode={true} value={summary} onChange={this.onChangeSummary} />
        </div>
        {/* Market */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelMarkets}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipMarkets} ></Tooltip>
          </div>
          <MultiselectWrapper
            data={marketsChoices}
            dataKey={(item: IDropdownOption) => item.key}
            textField={(item: IDropdownOption) => item.text}
            value={markets}
            filter="contains"
            onChange={this.onChangeMarket} />
          {errors.markets &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.markets}
              </span>
            </div>
          }
        </div>


        {/* Requirement Type */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelRequirementType}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipRequirementType} ></Tooltip>
          </div>
          <Combobox
            data={requirementTypeChoices}
            value={selectedRequirementType}
            dataKey="key"
            textField="text"
            selectIcon={
              <span className="ms-Dropdown-caretDownWrapper">
                <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
              </span>
            }


            filter="contains"
            groupBy={(type) => (type as ICategorizedDropdownOption).category}
            onChange={this.onChangeRequirementType}
          />

          {errors.requirementType &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.requirementType}
              </span>
            </div>
          }
        </div>
        {/* MA Verticals */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelMAVertical}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipMAVertical} ></Tooltip>
          </div>
          <MultiselectWrapper
            value={maVertical}
            data={maVerticalChoices}
            dataKey={(item: IDropdownOption) => item.key}
            textField={(item: IDropdownOption) => item.text}
            filter="contains"
            onChange={this.onChangeMAVerticals} />

          {errors.maVertical &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.maVertical}
              </span>
            </div>
          }

        </div>


        {/* Synthesis Status */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelSynthesis}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipSynthesis} ></Tooltip>
          </div>
          <Dropdown

            selectedKey={synthesisStatus}
            onChange={this.onChangeSynthesis}
            options={synthesisStatusChoices} />

          {errors.synthesisStatus &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.synthesisStatus}
              </span>
            </div>
          }
        </div>
        {/* RAG Status */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelRagStatus}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipRagStatus} ></Tooltip>
          </div>
          <Dropdown
            selectedKey={ragStatus}
            onChange={this.onChangeRagStatus}
            options={ragStatusChoices} />
        </div>


        {/* SNPs*/}
        {maVertical?.some((v) => v.key === "SNP") && (<div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelSNPS}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipSNPS} ></Tooltip>
          </div>
          <MultiselectWrapper
            data={snpsChoices}
            dataKey={(item: IDropdownOption) => item.key}
            textField={(item: IDropdownOption) => item.text}
            value={snps}
            filter="contains"
            onChange={this.onChangeSnps} />
          {errors.snps &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.snps}
              </span>
            </div>
          }
        </div>)}





        {/* Responsible */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelResponsibleParty}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipResponsibleParty} ></Tooltip>
          </div>
          <Dropdown
            selectedKey={responsibleParty}
            onChange={this.onChangeresponsibleParty}
            options={responsiblePartyChoices} />
        </div>


        {/* Responsible Item */}

        <div className={styles.field} >
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelResponsiblePartyItem}</Label>
            </div>
            {<Tooltip content={strings.FormMARequirementTooltipResponsiblePartyItem} ></Tooltip>}
          </div>
          <Combobox
            data={responsiblePartyItemChoices}
            value={responsiblePartyItem}
            selectIcon={
              <span className="ms-Dropdown-caretDownWrapper">
                <i data-icon-name="ChevronDown" aria-hidden="true" className="ms-Dropdown-caretDown"></i>
              </span>
            }
            textField="text"
            filter="contains"
            key="key"
            onChange={this.onChangeChargeByItem}
          />
          {errors.responsiblePartyItem &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.responsiblePartyItem}
              </span>
            </div>
          }
        </div>
        {/* Contacts*/}
        {responsibleParty !== MARequirementResponsibleParty.Eutelsat && (<div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelContacts}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipContacts} ></Tooltip>
          </div>
          <MultiselectWrapper
            data={responsiblePartyContactsChoices}
            dataKey={(item: IDropdownOption) => item.key}
            textField={(item: IDropdownOption) => item.text}
            value={responsiblePartyContacts}
            filter="contains"
            onChange={this.onChangeContacts} />
          {/*errors.responsiblePartyContacts &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.responsiblePartyContacts}
              </span>
            </div>*/
          }
        </div>)}
        {/* Eutlsat Owners*/}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label required>{strings.FormMARequirementLabelEutelsatOwners}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipEutelsatOwners} ></Tooltip>
          </div>
          <MultiselectWrapper
            data={eutelsatOwnersChoices}
            dataKey={(item: IDropdownOption) => item.key}
            textField={(item: IDropdownOption) => item.text}
            value={eutelsatOwners}
            filter="contains"
            onChange={this.onChangeEutelsatOwners} />
          {errors.eutelsatOwners &&
            <div className={styles.errorContainer}>
              <Icon iconName="Error" className={styles.errorIcon} />
              <span className={styles.errorMessage}>
                {errors.eutelsatOwners}
              </span>
            </div>
          }
        </div>




        {/* GEO/LEO */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelGeoLeo}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipGeoLeo} ></Tooltip>
          </div>
          <Dropdown

            selectedKey={geoLeo}
            onChange={this.onChangeGeoLeo}
            options={geoLeoChoices} />
        </div>
        {/* Comments */}
        <div className={styles.field}>
          <div className={styles.fieldLabel}>
            <div className={styles.fieldLabelContainer}>
              <Label>{strings.FormMARequirementLabelComments}</Label>
            </div>
            <Tooltip content={strings.FormMARequirementTooltipComments} ></Tooltip>
          </div>
          <RichText isEditMode={true} value={comments} onChange={this.onChangeComments} />
        </div>
      </>
    );
  }
}
