import * as React from 'react';
import { FC, useRef} from 'react';
import { TextField } from '@fluentui/react';
import { Dropdown, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { DatePicker, defaultDatePickerStrings } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
import { Stack, IStackProps, IStackStyles } from '@fluentui/react/lib/Stack';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useFormContext } from './FormContext';

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

const dropdownControlledExampleOptions = [
    { key: 'Business Leader', text: 'Business Leader' },
    { key: 'Office Leader', text: 'Office Leader' },
    { key: 'Americas BIM Manager', text: 'Americas BIM Manager' },
];

const MainForm: FC<IMainFormProps> = (props) => {
    const { title, setTitle,roleTitle, setRoleTitle} = useFormContext();
    console.log(title, "from MainForm");
    console.log(roleTitle,"from main")

    const onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
        setRoleTitle(item.text);
    };

    const _getPeoplePickerItems = (items: any[]) => {
        console.log('Items:', items);
    };

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
                            options={dropdownControlledExampleOptions}
                            styles={dropdownStyles}
                            required={true}
                        />
                        <DatePicker
                            placeholder="Select a date"
                            label='Date of Board Ratification Level'
                            strings={defaultDatePickerStrings}
                        />
                        <PeoplePicker
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
                        value={title}
                        type='number' 
                        placeholder='Enter a value here'
                        onChange={(e, v) => setTitle(v !== undefined ? v : '')} 
                    />
                    <TextField label="Current Appointments (Multi lines)" multiline autoAdjustHeight scrollContainerRef={containerRef} />
                </Stack>
            </Stack>
        </React.Fragment>
    );
};

export default MainForm;
