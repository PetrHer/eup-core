import { IStructureNode } from '../../interfaces';

export interface IStructureApiClient {
    getStructure(): Promise<IStructureNode[]>;
}
