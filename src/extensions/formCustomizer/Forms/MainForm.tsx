import * as React from 'react';
import { FC, useEffect, useRef, useState } from 'react';
import { TextField, Dropdown, IDropdownOption, IDropdownStyles, DatePicker, defaultDatePickerStrings, Stack, IStackProps, IStackStyles } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPFI } from '@pnp/sp';
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useFormContext } from './FormContext';
import { getChoicesFromSharePointList } from './spHelper';

export interface IMainFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
    tokens: { childrenGap: 15 },
    styles: { root: { width: 300 } },
};

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

// const dropdownControlledExampleOptions = [
//     { key: 'Business Leader', text: 'Business Leader' },
//     { key: 'Office Leader', text: 'Office Leader' },
//     { key: 'Americas BIM Manager', text: 'Americas BIM Manager' },
// ];

const MainForm: FC<IMainFormProps> = (props) => {
    const { title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue, selectedUsers, setSelectedUsers, maxRole,setMaxRole,peoplePickerKey,Appointments, setAppointments } = useFormContext();
    const [dropdownOptions, setDropdownOptions] = useState<IDropdownOption[]>([]);
    const [isLoading, setIsLoading] = useState(true);

    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setRoleTitle(item.text);
    };

    const onSelectDate = (date: Date | null | undefined): void => {
        if (date) {
            setDateValue(date);
        }
    };

    const _getPeoplePickerItems = (items: any[]) => {
        setSelectedUsers(items);
    };

    useEffect(() => {
        console.log('Selected User:', selectedUsers);
    }, [selectedUsers]);

    useEffect(() => {
        const fetchDropdownChoices = async () => {
            try {
                console.log(isLoading)
                const choices = await getChoicesFromSharePointList(props.sp, props.listGuid.toString(), 'RoleTitle');
                const options = choices.map(choice => ({ key: choice, text: choice }));
                setDropdownOptions(options);
                setIsLoading(false);
            } catch (error) {
                console.error('Error fetching choices from SharePoint list:', error);
                setIsLoading(false);
            }
        };
        fetchDropdownChoices();
    }, [props.sp, props.listGuid]);

    const containerRef = useRef<HTMLDivElement>(null);
    
    return (
        <React.Fragment>
            <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                <Stack {...columnProps}>
                    <div style={{ marginLeft: '10px', marginRight: '10px' }}>
                        <TextField label="Enter Title (Single Line Text):" value={title} onChange={(e, v) => setTitle(v !== undefined ? v : '')} />
                        <Dropdown
                            label="Enter Role Title (Dropdown)"
                            selectedKey={roleTitle}
                            onChange={onChange}
                            placeholder="Select an option"
                            options={dropdownOptions}
                            styles={dropdownStyles}
                            required={true}
                        />
                        <DatePicker
                            placeholder="Select a date"
                            label='Date of Board Ratification Level'
                            strings={defaultDatePickerStrings}
                            value={dateValue}
                            onSelectDate={onSelectDate}
                        />
                        <PeoplePicker
                            key={peoplePickerKey}
                            context={{
                                absoluteUrl: props.context.pageContext.web.absoluteUrl,
                                msGraphClientFactory: props.context.msGraphClientFactory,
                                spHttpClient: props.context.spHttpClient
                            }}
                            titleText="Incumbent (People Picker)"
                            personSelectionLimit={1}
                            groupName=''
                            showtooltip={true}
                            required={false}
                            disabled={false}
                            searchTextLimit={3}
                            onChange={_getPeoplePickerItems}
                            principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                            resolveDelay={1000}
                            placeholder='Enter a name or email address'
                            
                        />
                    </div>
                </Stack>
                <Stack {...columnProps}>
                    <TextField 
                        label="Max Role Term Length (Number Field)" 
                        value={maxRole !== undefined ? maxRole.toString() : ''}
                        type='number' 
                        placeholder='Enter a value here'
                        onChange={(e, v) => setMaxRole(v !== undefined ? parseInt(v) : undefined)}  
                    />
                    <TextField 
                        label="Current Appointments (Multi lines)" 
                        multiline 
                        autoAdjustHeight 
                        scrollContainerRef={containerRef} 
                        value={Appointments} 
                        onChange={(e, v) => setAppointments(v !== undefined ? v : '')} 
                        />
                </Stack>
            </Stack>
        </React.Fragment>
    );
};

export default MainForm;
