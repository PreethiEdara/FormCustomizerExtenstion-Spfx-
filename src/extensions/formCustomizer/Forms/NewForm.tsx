import * as React from 'react';
import { useState, FC } from 'react';
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


export interface INewFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}

const NewForm: FC<INewFormProps> = (props) => {
    const { title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue, selectedUsers,setSelectedUsers, setPeoplePickerKey, maxRole, setMaxRole,Appointments, setAppointments, setErrMsg,isPanelOpen, setIsPanelOpen} = useFormContext();
    const [msg, setMsg] = useState<any>(undefined);
    

    const clearControls = () => {
        setTitle('');
        setRoleTitle('');
        setDateValue(undefined);
        setSelectedUsers([])
        setPeoplePickerKey(Math.random().toString());
        setMaxRole(undefined);
        setAppointments('');
        setErrMsg(false);
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
        setMsg(undefined);

        if (!roleTitle) {
            setErrMsg(true);
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
            await props.sp.web.lists.getById(props.listGuid.toString()).items.add({
                Title: title,
                RoleTitle: roleTitle,
                DateOfBoardRatificationLevel: formattedDate,
                ...incumbentField,
                MaxRole: maxRole,
                CurrentAppointments: Appointments
            });
            setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
            clearControls();
        } catch (error) {
            console.error(error);
            setMsg({ scope: MessageBarType.error, Message: 'Error creating item.' });
        }
    };

    const onClosePanel = () => {
        setIsPanelOpen(false);
        props.onClose();
    };

    return (
        <React.Fragment>
            <Panel
                isOpen={isPanelOpen}
                type={PanelType.custom}
                customWidth="700px"
                onDismiss={onClosePanel}>
            <div className={styles.mainForm}>
            <h2>New Item</h2>
            <MainForm 
                sp={props.sp} 
                context={props.context} 
                listGuid={props.listGuid} 
                onClose={props.onClose} 
                onSave={saveListItem} 
            />
            {msg && (
                <MessageBar className={styles.msgBar}
                    messageBarType={msg.scope}>
                    {msg.Message}
                </MessageBar>
            )}
            </div>
            <PrimaryButton className={styles.btn} text="Save" onClick={saveListItem} />
            <DefaultButton text="Cancel" onClick={props.onClose} />
            </Panel>
        </React.Fragment>
    );
};

export default NewForm;
