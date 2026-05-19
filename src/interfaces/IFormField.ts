export interface IFormField {
    type: string,
    label: string,
    field: string,
    options?: { label: string; value: string | number }[],
}
