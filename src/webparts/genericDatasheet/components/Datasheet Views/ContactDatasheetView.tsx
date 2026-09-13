import * as React from 'react';
import { IContactDatasheetViewProps, IContactDatasheetViewState } from './DatasheetViews.types';
import styles from '../Datasheet.module.scss';
import * as strings from 'GenericDatasheetStrings';
import { Icon, ActionButton, Modal } from '@fluentui/react';
import { IContact } from '../../../../common/models/IBusiness';
import Contact from '../../../genericForm/components/Contact/Contact';
import { FormOrigin } from '../../../genericForm/components/Form.types';
import { UserRole } from '../../../../common/models/Enums';
import { Config } from '../../../../common/config/Config';
import { Consts } from '../../../../common/consts/Consts';
import { DatasheetType } from '../../../../webparts/genericDatasheet/components/Datasheet.types';
import { UrlHelper } from '../../../../common/helpers/UrlHelper';

export default class ContactDatasheetView extends React.Component<IContactDatasheetViewProps, IContactDatasheetViewState> {
    constructor(props: IContactDatasheetViewProps) {
        super(props);
        this.state = {
            showFormDialog: false,
            itemId: undefined
        };
    }
    private hasAnyRole = (...rolesToCheck: UserRole[]): boolean => {
        const userRoles = this.props.userService.userContext?.userRoles ?? [];
        return rolesToCheck.some((role) => userRoles.includes(role));
    };

    private renderContactCard = (contact: IContact): JSX.Element => (
        <div key={contact.Id} className={styles.contactCardItem} style={{ cursor: 'pointer' }}
            onClick={(e) => {
                e.preventDefault();
                const url = `${Config.SITE.ROOT_WEB_URL}/${Consts.Pages.DATASHEET_URL}?config=${encodeURIComponent(DatasheetType.Contact)}&itemId=${contact.Id}`;
                UrlHelper.navigate(url, false);
            }}>
            <div className={styles.contactCardContent}>
                <div className={styles.contactCardHeader}>
                    <Icon iconName="Contact" className={styles.contactIconLarge} />
                    <div className={styles.contactCardName}>
                        {(contact.FirstName || contact.LastName) ?
                            `${contact.FirstName || ''} ${contact.LastName || ''}`.trim() :
                            contact.Title || 'Unknown'}
                    </div>
                </div>
                <div className={styles.contactCardSeparator}></div>
                <div className={styles.contactCardProperties}>
                    <div className={styles.contactProperty}>
                        <Icon iconName="Work" className={styles.propertyIconSmall} />
                        <span>{contact.JobTitle || ''}</span>
                    </div>
                    <div className={styles.contactProperty}>
                        <Icon iconName="Org" className={styles.propertyIconSmall} />
                        <span>{contact.Department || ''}</span>
                    </div>
                    {/* {contact.Email && (
                        <div className={styles.contactProperty}>
                            <Icon iconName="Mail" className={styles.propertyIconSmall} />
                            <span>{contact.Email}</span>
                        </div>
                    )} */}
                </div>

            </div>
        </div>
    );

    public render(): React.ReactElement<IContactDatasheetViewProps> {
        const { contacts } = this.props;
        const { showFormDialog } = this.state;

        if (!contacts || contacts.length === 0) {
            return (
                <>
                    {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.actionButton}>
                        <ActionButton
                            iconProps={{ iconName: 'Add' }}
                            onClick={this.openContactNewForm}
                            title={strings.ContactDataSheetNewContact}
                        >
                            {strings.ContactDataSheetNewContact}
                        </ActionButton>
                    </div>}
                    <div className={styles.noDataMessage}>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                        <span>{strings.DataSheetNoContacts || 'No record found'}</span>
                        <Icon iconName="Remove" className={styles.mutedIcon} />
                    </div>
                    {showFormDialog && this.renderFormDialog()}
                </>
            );
        }

        return (
            <>
                {this.hasAnyRole(UserRole.Contributor, UserRole.FinanceContributor, UserRole.Owner, UserRole.Admin) && <div className={styles.actionButton}>
                    <ActionButton
                        iconProps={{ iconName: 'Add' }}
                        onClick={this.openContactNewForm}
                        title={strings.ContactDataSheetNewContact}
                    >
                        {strings.ContactDataSheetNewContact}
                    </ActionButton>
                </div>}
                <div className={styles.contactsContainer}>
                    {contacts.map((contact) => this.renderContactCard(contact))}
                </div>
                {showFormDialog && this.renderFormDialog()}
            </>
        );
    }

    private openContactNewForm = () => {
        this.setState({
            showFormDialog: true,
            itemId: undefined,
        });
    };

    private closeFormDialog = () => {
        this.setState({ showFormDialog: false });
    };

    private refresh = async (refreshData: boolean) => {
        this.closeFormDialog();
        if (refreshData && this.props.refresh) {
            await this.props.refresh();
        }
    };

    private renderFormDialog() {
        const { pnpService, userService, relatedItemId, relatedItemType } = this.props;
        const { showFormDialog } = this.state;
        let formControl = null;
        let formTitle = null;
        formControl = (
            <Contact
                pnpService={pnpService}
                userService={userService}
                preselectedItemId={relatedItemId}
                preselectedType={relatedItemType}
                callback={() => this.refresh(true)}
                formOrigin={FormOrigin.Datasheet}
            />
        );
        formTitle = 'New Contact';
        return (
            <Modal
                isOpen={showFormDialog}
                onDismiss={this.closeFormDialog}
                containerClassName={styles.modalContainer}
                isBlocking={true}
            >
                <div>
                    <div className={styles.modalHeader}>
                        {formTitle}
                        <ActionButton
                            iconProps={{ iconName: 'Cancel' }}
                            onClick={this.closeFormDialog}
                            className={styles.iconButton}
                        >
                            {strings.DataSheetCancelButton}
                        </ActionButton>
                    </div>
                    <div className={styles.modalBody}>
                        {formControl}
                    </div>
                </div>
            </Modal>
        );
    }
}
