type CreateFieldsPayload = {
    siteId: string;
    listId: string;
    fields: Record<string, any>;
};
export declare function createListItem(token: string, input: CreateFieldsPayload): Promise<any>;
export {};
//# sourceMappingURL=sharepoint.d.ts.map