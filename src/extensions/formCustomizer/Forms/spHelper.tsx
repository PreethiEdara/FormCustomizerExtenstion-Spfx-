import { SPFI } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/fields";

export const getChoicesFromSharePointList = async (sp: SPFI, listGuid: string, fieldName: string): Promise<string[]> => {
    try {
        const field = await sp.web.lists.getById(listGuid).fields.getByInternalNameOrTitle(fieldName)();
        if (field && field.Choices) {
            return field.Choices;
        } else {
            throw new Error(`Field ${fieldName} does not have choices or does not exist.`);
        }
    } catch (error) {
        console.error('Error fetching choices:', error);
        throw error;
    }
};
