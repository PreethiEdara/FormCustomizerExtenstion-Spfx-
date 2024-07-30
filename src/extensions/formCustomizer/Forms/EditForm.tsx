import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import { PrimaryButton, DefaultButton, Panel, PanelType } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import MainForm from './MainForm';
import { useFormContext } from './FormContext';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import styles from '../components/FormCustomizer.module.scss';

export interface IEditFormProps {
    sp: SPFI;
    listGuid: Guid;
    context: WebPartContext;
    itemId: number;
    onSave: () => void;
    onClose: () => void;
}

const EditForm: FC<IEditFormProps> = (props) => {
    const { title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue, selectedUsers, setSelectedUsers, setPeoplePickerKey, maxRole, setMaxRole, Appointments, setAppointments, isPanelOpen, setIsPanelOpen, showEditPanel, setShowEditPanel } = useFormContext();
    const [msg, setMsg] = useState<any>(undefined);

    const clearControls = () => {
        setTitle('');
        setRoleTitle('');
        setDateValue(undefined);
        setSelectedUsers([]);
        setPeoplePickerKey(Math.random().toString());
        setMaxRole(undefined);
        setAppointments('');
        setIsPanelOpen(false);
        props.onClose();
    };

    const getUserId = async (loginName: string) => {
        try {
            const user = await props.sp.web.ensureUser(loginName);
            return user.data.Id;
        } catch (error) {
            console.error('Error getting user ID:', error);
            return null;
        }
    };

    const saveListItem = async () => {
        if (!roleTitle) {
            setMsg({ scope: MessageBarType.error, Message: 'Role Title is required.' });
            return;
        }

        const formattedDate = dateValue ? dateValue : undefined;

        let incumbentField = {};
        if (selectedUsers && selectedUsers.length > 0) {
            const userId = await getUserId(selectedUsers[0].loginName);
            if (userId) {
                incumbentField = { IncumbentId: userId };
            } else {
                setMsg({ scope: MessageBarType.error, Message: 'Error getting user ID.' });
                return;
            }
        }

        try {
            await props.sp.web.lists.getById(props.listGuid.toString()).items.getById(props.itemId).update({
                Title: title,
                RoleTitle: roleTitle,
                DateOfBoardRatificationLevel: formattedDate,
                ...incumbentField,
                MaxRole: maxRole,
                CurrentAppointments: Appointments
            });
            setMsg({ scope: MessageBarType.success, Message: 'Save successful!' });

            if (showEditPanel) {
                setShowEditPanel(false);
                props.onClose();
            } else {
                clearControls();
            }
        } catch (error) {
            console.error('Error saving item:', error);
            setMsg({ scope: MessageBarType.error, Message: 'Error saving item.' });
        }
    };

    const populateItemForEdit = async () => {
        if (props.itemId) {
            let itemToUpdate: any = await props.sp.web.lists.getById(props.listGuid.toString()).items
                .select('ID', 'Title', 'RoleTitle', 'DateOfBoardRatificationLevel', 'IncumbentId', 'MaxRole', 'CurrentAppointments')
                .getById(props.itemId)();            
            if (itemToUpdate) {
                setTitle(itemToUpdate.Title);
                setRoleTitle(itemToUpdate.RoleTitle);
                setDateValue(itemToUpdate.DateOfBoardRatificationLevel ? new Date(itemToUpdate.DateOfBoardRatificationLevel) : undefined);
                if (itemToUpdate.IncumbentId) {
                    const user = await props.sp.web.siteUsers.getById(itemToUpdate.IncumbentId)();
                    setSelectedUsers([{ loginName: user.Title }]); 
                } else {
                    setSelectedUsers([]);
                }
                console.log("Incumbent:", itemToUpdate.Incumbent); 
                console.log("Selected Users:", selectedUsers);
                setMaxRole(itemToUpdate.MaxRole);
                setAppointments(itemToUpdate.CurrentAppointments);
            }
        } else {
            setMsg({ scope: MessageBarType.error, Message: 'Sorry, item not found!' });
        }
    };

    useEffect(() => {
        populateItemForEdit();
    }, []);

    const onClosePanel = () => {
        setIsPanelOpen(false);
        props.onClose();
    };

    return (
        <React.Fragment>
            <Panel
                isOpen={isPanelOpen}
                type={PanelType.custom}
                isLightDismiss
                customWidth="700px"
                onDismiss={onClosePanel}
            >
                <div className={styles.mainForm}>
                    <h2>{title}</h2>
                    <MainForm
                        sp={props.sp}
                        context={props.context}
                        listGuid={props.listGuid}
                        onClose={props.onClose}
                        onSave={saveListItem}
                    />
                    {msg && (
                        <MessageBar className={styles.msgBar} messageBarType={msg.scope}>
                            {msg.Message}
                        </MessageBar>
                    )}
                </div>
                <PrimaryButton className={styles.btn} text="Save" onClick={saveListItem} />
                <DefaultButton text="Cancell" onClick={props.onClose} />
            </Panel>
        </React.Fragment>
    );
};

export default EditForm;
