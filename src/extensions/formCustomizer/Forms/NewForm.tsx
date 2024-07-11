import * as React from 'react';
import { useState, FC } from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react';
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
// import styles from './FormCustomizer.module.scss';


export interface INewFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}

const NewForm: FC<INewFormProps> = (props) => {
    const { title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue, selectedUsers,setSelectedUsers, setPeoplePickerKey, maxRole, setMaxRole,Appointments, setAppointments } = useFormContext();
    const [msg, setMsg] = useState<any>(undefined);

    const clearControls = () => {
        setTitle('');
        setRoleTitle('');
        setDateValue(undefined);
        setSelectedUsers([])
        setPeoplePickerKey(Math.random().toString());
        setMaxRole(undefined);
        setAppointments('')
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
        
        // Ensure dateValue is a valid Date object
        const formattedDate = dateValue ? dateValue : undefined;

        let incumbentField = {};
        if (selectedUsers) {
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
                DateofBoardRatificationLevel: formattedDate,
                ...incumbentField,
                MaxRoleTermLength: maxRole,
                CurrentAppointments: Appointments
            });
            setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
            clearControls();
        } catch (error) {
            console.error(error);
            setMsg({ scope: MessageBarType.error, Message: 'Error creating item.' });
        }
    };

    return (
        <React.Fragment>
            <div>New Form</div>
            <div style={{margin:'1%'}}>
            <MainForm 
                sp={props.sp} 
                context={props.context} 
                listGuid={props.listGuid} 
                onClose={props.onClose} 
                onSave={saveListItem} 
            />
            {msg && (
                <MessageBar messageBarType={msg.scope}>
                    {msg.Message}
                </MessageBar>
            )}
            </div>
            <PrimaryButton style={{margin:'0.5% 2%'}} text="Save" onClick={saveListItem} />
            <DefaultButton text="Cancell" onClick={props.onClose} />
        </React.Fragment>
    );
};

export default NewForm;
