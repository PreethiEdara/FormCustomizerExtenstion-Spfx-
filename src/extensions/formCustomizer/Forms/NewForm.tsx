import * as React from 'react';
import { useState, FC } from 'react';
import { TextField } from '@fluentui/react';
import { PrimaryButton, DefaultButton } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
 
export interface INewFormProps {
    sp: SPFI;
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
            <div style={{ margin: '10px' }}>
                <TextField label="Enter Title:" value={title} onChange={(e, v) => setTitle(v !== undefined ? v : '')} />
                <PrimaryButton text="Save" onClick={saveListItem} />
                <DefaultButton text="Cancel" onClick={props.onClose} style={{ marginLeft: '10px' }} />
            </div>
            {msg && msg.Message &&
                <MessageBar messageBarType={msg.scope ? msg.scope : MessageBarType.info}>{msg.Message}</MessageBar>
            }
        </React.Fragment>
    );
};
 
export default NewForm;