import * as React from 'react';
import { useState, FC } from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import MainForm from './MainForm';
import { useFormContext } from './FormContext';

export interface INewFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}

const NewForm: FC<INewFormProps> = (props) => {
    const { title, setTitle,roleTitle, setRoleTitle } = useFormContext();
    const [msg, setMsg] = useState<any>(undefined);

    const clearControls = () => {
        setTitle('');
        setRoleTitle('')
    };

    const saveListItem = async () => {
        setMsg(undefined);
        await props.sp.web.lists.getById(props.listGuid.toString()).items.add({
            Title: title,
            RoleTitle : roleTitle
        });
        console.log(title, "from NewForm");
        console.log(roleTitle, "from new")
        setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
        clearControls();
    };

    return (
        <React.Fragment>
            <div>New Form</div>
            <MainForm 
                sp={props.sp} 
                context={props.context} 
                listGuid={props.listGuid} 
                onClose={props.onClose} 
                onSave={props.onSave}
            />
            <PrimaryButton text="Save" onClick={saveListItem} />
            <DefaultButton text="Cancel" onClick={props.onClose} style={{ marginLeft: '10px' }} /> 
                    
            {msg && msg.Message &&
            <MessageBar messageBarType={msg.scope ? msg.scope : MessageBarType.info}>{msg.Message}</MessageBar>
            } 
        </React.Fragment>
    );
};

export default NewForm;
