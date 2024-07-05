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
export interface INewFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}


const NewForm: FC<INewFormProps> = (props) => {
    const [title, setTitle] = useState<string>('');
    const [msg, setMsg] = useState<any>(undefined);

    const clearControls = () => {
        setTitle('');
    };


    const saveListItem = async () => {
        setMsg(undefined);
        await props.sp.web.lists.getById(props.listGuid.toString()).items.add({
            Title: title
        });
        setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
        clearControls();
    };

    return (
        <React.Fragment>
            <div>New Form</div>
            <MainForm sp={props.sp} context={props.context} listGuid={props.listGuid} onClose={props.onClose} onSave={props.onSave}/>
            <PrimaryButton text="Savee" onClick={saveListItem} />
            <DefaultButton text="Cancell" onClick={props.onClose} style={{ marginLeft: '10px' }} /> 
                    
            {msg && msg.Message &&
            <MessageBar messageBarType={msg.scope ? msg.scope : MessageBarType.info}>{msg.Message}</MessageBar>
            } 
            
        </React.Fragment>
    );
};

export default NewForm;
