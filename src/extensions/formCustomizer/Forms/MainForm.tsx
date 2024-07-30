import * as React from 'react';
import { FC, useEffect, useRef, useState } from 'react';
import { TextField, Dropdown, IDropdownOption, IDropdownStyles, DatePicker, defaultDatePickerStrings, Stack, IStackTokens, IDatePicker } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPFI } from '@pnp/sp';
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useFormContext } from './FormContext';
import { getChoicesFromSharePointList } from './spHelper';
// import styles from '../components/FormCustomizer.module.scss';


export interface IMainFormProps {
    sp: SPFI;
    context: WebPartContext;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}


const stackTokens: IStackTokens = { childrenGap: 15 };

const dropdownStyles: Partial<IDropdownStyles> = { 
    dropdown: { width: 638 }, 
    callout: { maxHeight: 200 } 
};


const MainForm: FC<IMainFormProps> = (props) => {
    const { title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue, selectedUsers, setSelectedUsers, maxRole, setMaxRole, peoplePickerKey, Appointments, setAppointments, errmsg } = useFormContext();
    const [dropdownOptions, setDropdownOptions] = useState<IDropdownOption[]>([]);
    const [isLoading, setIsLoading] = useState(true);

    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setRoleTitle(item.text);
    };

    const onFormatDate = (date?: Date): string => {
        return !date ? '' : date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    };

    const onParseDateFromString = React.useCallback(
        (newValue: string): Date => {
            const previousValue = dateValue || new Date();
            const newValueParts = (newValue || '').trim().split('/');
            const day =
                newValueParts.length > 0 ? Math.max(1, Math.min(31, parseInt(newValueParts[0], 10))) : previousValue.getDate();
            const month =
                newValueParts.length > 1
                    ? Math.max(1, Math.min(12, parseInt(newValueParts[1], 10))) - 1
                    : previousValue.getMonth();
            let year = newValueParts.length > 2 ? parseInt(newValueParts[2], 10) : previousValue.getFullYear();
            if (year < 100) {
                year += previousValue.getFullYear() - (previousValue.getFullYear() % 100);
            }
            return new Date(year, month, day);
        },
        [dateValue],
    );

    const _getPeoplePickerItems = (items: any[]) => {
        setSelectedUsers(items);
    };

    useEffect(() => {
        console.log('Selected User:', selectedUsers);
    }, [selectedUsers]);

    useEffect(() => {
        const fetchDropdownChoices = async () => {
            try {
                console.log(isLoading);
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
    const datePickerRef = React.useRef<IDatePicker>(null);

    return (
        <React.Fragment>
            <Stack tokens={stackTokens}>
                <TextField 
                    label="Enter Title (Single Line Text):" 
                    placeholder='Enter the Title'
                    value={title} 
                    onChange={(e, v) => setTitle(v !== undefined ? v : '')} 
                />
                <Dropdown 
                    label="Enter Role Title (Dropdown)"
                    selectedKey={roleTitle}
                    onChange={onChange}
                    placeholder="Select an option"
                    options={dropdownOptions}
                    styles={dropdownStyles}
                    errorMessage={errmsg ? 'This field is required' : undefined}
                    required={true}
                />
                <DatePicker
                    componentRef={datePickerRef}
                    label="Date of Board Ratification Level"
                    allowTextInput
                    placeholder='select a date'
                    value={dateValue}
                    onSelectDate={setDateValue as (date?: Date) => void}
                    formatDate={onFormatDate}
                    parseDateFromString={onParseDateFromString}
                    strings={defaultDatePickerStrings}
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
                    defaultSelectedUsers={selectedUsers.map(user => user.loginName)}
                />
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
        </React.Fragment>
    );
};

export default MainForm;
