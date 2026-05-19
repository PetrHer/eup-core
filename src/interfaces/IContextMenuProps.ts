import { MenuEventArgs, MenuItemModel } from "@syncfusion/ej2-react-navigations";

export interface IContextMenuProps {
    rect: DOMRect;
    items: MenuItemModel[];
    contextMenuSelect: (args: MenuEventArgs) => Promise<void>
}