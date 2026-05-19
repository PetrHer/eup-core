export interface ISPComment {
    text: string;
    author: {
        email: string;
        name: string;
        isExternal: boolean;
    };
    createdDate: string;
    id: number;
}