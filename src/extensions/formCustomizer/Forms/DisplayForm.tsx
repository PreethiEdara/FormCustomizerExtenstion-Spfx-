import * as React from 'react';
import { useEffect, useState, FC } from 'react';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import MainForm from './MainForm';
import { useFormContext } from './FormContext';
import "@pnp/sp/site-users/web";
import styles from '../components/FormCustomizer.module.scss';
import { MessageBar, MessageBarType, Panel, PanelType, PrimaryButton } from '@fluentui/react';
import EditForm from './EditForm';

export interface IDisplayFormProps {
    sp: SPFI;
    listGuid: Guid;
    context: WebPartContext;
    itemId: number;
    onClose: () => void;
}

const DisplayForm: FC<IDisplayFormProps> = (props) => {
    const { title, setTitle, setRoleTitle, setDateValue, setSelectedUsers, setMaxRole, setAppointments, isPanelOpen, setIsPanelOpen, showEditPanel, setShowEditPanel } = useFormContext();
    const [msg, setMsg] = useState<any>(undefined);

    const populateItemForDisplay = async () => {
        try {
            console.log("populateItemForDisplay called");
            let itemToUpdate: any = await props.sp.web.lists.getById(props.listGuid.toString()).items
                .select('ID', 'Title', 'RoleTitle', 'DateOfBoardRatificationLevel', 'IncumbentId', 'MaxRole', 'CurrentAppointments')
                .getById(props.itemId)();
            console.log("Item to update:", itemToUpdate);

            if (itemToUpdate) {
                setTitle(itemToUpdate.Title || '');
                setRoleTitle(itemToUpdate.RoleTitle || '');
                setDateValue(itemToUpdate.DateOfBoardRatificationLevel ? new Date(itemToUpdate.DateOfBoardRatificationLevel) : undefined);
                
                if (itemToUpdate.IncumbentId) {
                    const user = await props.sp.web.siteUsers.getById(itemToUpdate.IncumbentId)();
                    console.log("User data:", user);
                    setSelectedUsers([{ loginName: user.Title }]);
                } else {
                    setSelectedUsers([]);
                    console.log("No IncumbentId found");
                }
                
                setMaxRole(itemToUpdate.MaxRole || null);
                setAppointments(itemToUpdate.CurrentAppointments || '');
                
                console.log("Form Data Set: ", {
                    title: itemToUpdate.Title || '',
                    roleTitle: itemToUpdate.RoleTitle || '',
                    dateValue: itemToUpdate.DateOfBoardRatificationLevel ? new Date(itemToUpdate.DateOfBoardRatificationLevel) : undefined,
                    selectedUsers: itemToUpdate.IncumbentId ? [{ loginName: (await props.sp.web.siteUsers.getById(itemToUpdate.IncumbentId)()).Title }] : [],
                    maxRole: itemToUpdate.MaxRole || null,
                    appointments: itemToUpdate.CurrentAppointments || ''
                });
            } else {
                setMsg({ scope: MessageBarType.error, Message: 'Sorry, item not found!' });
                console.log("Item not found");
            }
        } catch (error) {
            setMsg({ scope: MessageBarType.error, Message: 'Error loading item: ' + error.message });
            console.error("Error loading item:", error);
        }
    };

    useEffect(() => {
        console.log("isPanelOpen changed:", isPanelOpen);
        if (isPanelOpen) {
            populateItemForDisplay();
        }
    }, [isPanelOpen]);

    const onClosePanel = () => {
        console.log("Closing panel");
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
                        onSave={() => {}}
                    />
                    {msg && (
                        <MessageBar className={styles.msgBar} messageBarType={msg.scope}>
                            {msg.Message}
                        </MessageBar>
                    )}
                </div>
                <div className={styles.displayBtn}>
                    <PrimaryButton text="Edit All" onClick={() => setShowEditPanel(true)} />
                </div>
                {showEditPanel && (
                    <Panel
                        isOpen={showEditPanel}
                        onDismiss={() => setShowEditPanel(false)}
                        type={PanelType.medium}
                        headerText="Edit Form"
                        closeButtonAriaLabel="Close"
                    >
                        <EditForm
                            sp={props.sp}
                            context={props.context}
                            listGuid={props.listGuid}
                            itemId={props.itemId}
                            onSave={() => setShowEditPanel(false)}
                            onClose={() => setShowEditPanel(false)}
                        />
                    </Panel>
                )}
            </Panel>
        </React.Fragment>
    );
};

export default DisplayForm;
