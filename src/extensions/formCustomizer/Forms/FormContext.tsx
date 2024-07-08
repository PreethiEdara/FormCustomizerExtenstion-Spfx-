import * as React from 'react';

interface IFormContext {
    title: string;
    setTitle: React.Dispatch<React.SetStateAction<string>>;
    roleTitle: string;
    setRoleTitle: React.Dispatch<React.SetStateAction<string>>;
}

const FormContext = React.createContext<IFormContext | undefined>(undefined);

const FormProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [title, setTitle] = React.useState<string>('');
    const [roleTitle, setRoleTitle] = React.useState<string>('');

    return (
        <FormContext.Provider value={{ title, setTitle, roleTitle, setRoleTitle }}>
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
