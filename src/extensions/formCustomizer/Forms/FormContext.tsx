
import * as React from 'react';

interface IFormContext {
    title: string;
    setTitle: React.Dispatch<React.SetStateAction<string>>;
    maxRole: number|undefined;
    setMaxRole: React.Dispatch<React.SetStateAction<number|undefined>>;
    roleTitle: string;
    setRoleTitle: React.Dispatch<React.SetStateAction<string>>;
    dateValue : Date|undefined;
    setDateValue: React.Dispatch<React.SetStateAction<Date | undefined>>;
    selectedUsers: any[];
    setSelectedUsers: React.Dispatch<React.SetStateAction<any[]>>;
    peoplePickerKey: string;
    setPeoplePickerKey: React.Dispatch<React.SetStateAction<string>>;
    Appointments: string;
    setAppointments: React.Dispatch<React.SetStateAction<string>>;
}

const FormContext = React.createContext<IFormContext | undefined>(undefined);

const FormProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [title, setTitle] = React.useState<string>('');
    const [maxRole, setMaxRole] = React.useState<number|undefined>(undefined)
    const [roleTitle, setRoleTitle] = React.useState<string>('');
    const [dateValue, setDateValue] = React.useState<Date|undefined>(undefined);
    const [selectedUsers, setSelectedUsers] = React.useState<any[]>([]);
    const [peoplePickerKey, setPeoplePickerKey] = React.useState<string>(Math.random().toString());
    const [Appointments, setAppointments] = React.useState<string>('');

    return (
        <FormContext.Provider value={{ title, setTitle, roleTitle, setRoleTitle, dateValue, setDateValue,selectedUsers, setSelectedUsers, maxRole, setMaxRole,peoplePickerKey, setPeoplePickerKey,Appointments,setAppointments}}>
            {children}
        </FormContext.Provider>
    );
};

export const useFormContext = () => {
    const context = React.useContext(FormContext);
    if (!context) {
        throw new Error('useFormContext must be used within a FormProvider');
    }
    return context;
};
export default FormProvider;
