import { ColumnsCategories, FolderStatus, Language } from "../../enums";
import { IStructureNode } from "../../interfaces";
import { cleanFolderName } from "../../utils/utils";

/**
 * Returns a new array of structure nodes sorted alphabetically by their `name` property.
 *
 * @param nodes - Array of structure nodes to sort.
 * @returns A new array sorted by node name in ascending order.
 */
const sortStructureNodesByName = (nodes: IStructureNode[]): IStructureNode[] => {
    const sorted = [...nodes].sort((a, b) => a.name.localeCompare(b.name));
    return sorted;
};

/**
 * Maps a raw data item to a category folder object.
 * @param item The raw data item.
 * @returns A mapped category folder object.
 */
const mapCategoryNode = (item: any): IStructureNode => {
    return {
        name: item.Title,
        id: item.Id,
        uploadedFiles: [],
        nodes: [],
        status: item[ColumnsCategories.CategoryStatus] ? item[ColumnsCategories.CategoryStatus] : FolderStatus.Default,
        description: {
            [Language.CS]: item[ColumnsCategories.Description],
            [Language.EN]: item[ColumnsCategories.DescriptionEn]
        },
        isCategory: true,
        localization: {
            [Language.CS]: cleanFolderName(item.Title) ?? '',
            [Language.EN]: cleanFolderName(item[ColumnsCategories.CategoryEn]) ?? ''
        },
        requiredFiles: [],
        order: item[ColumnsCategories.Order]
    };
};

/**
 * Maps raw data to an array of category folders.
 * @param data The raw data array.
 * @returns An array of mapped category folders.
 */
export const mapCategoryNodes = (data: any[]): IStructureNode[] => {
    return data.map(item => mapCategoryNode(item));
};

/**
    * Maps raw data to a structured hierarchy of folders categorized by stage, phase, and category.
    * @param categories The raw data array.
    * @returns An array of structured folder objects.
    */
export const mapStructure = (categories: any[], phasesQnA: any[]): IStructureNode[] => {
    const nodes: IStructureNode[] = [];
    for (let i = 0; i < categories.length; i++) {
        const item = categories[i];
        let stageNode = nodes.find(fol => fol.name === item[ColumnsCategories.Stage]);
        if (!stageNode) {
            stageNode = {
                name: item[ColumnsCategories.Stage],
                uploadedFiles: [],
                nodes: [],
                status: FolderStatus.Default,
                localization: {
                    [Language.CS]: cleanFolderName(item[ColumnsCategories.Stage]) ?? '',
                    [Language.EN]: cleanFolderName(item[ColumnsCategories.StageEn]) ?? ''
                }
            };
            nodes.push(stageNode);
        }
        let phaseNode = stageNode.nodes.find(fol => fol.name === item[ColumnsCategories.Phase]);
        if (!phaseNode) {
            phaseNode = {
                name: item[ColumnsCategories.Phase],
                uploadedFiles: [],
                nodes: [],
                status: FolderStatus.Default,
                localization: {
                    [Language.CS]: cleanFolderName(item[ColumnsCategories.Phase]) ?? '',
                    [Language.EN]: cleanFolderName(item[ColumnsCategories.PhaseEn]) ?? ''
                }
            };
            stageNode.nodes.push(phaseNode);
        }
        let categoryNode = phaseNode.nodes.find(fol => fol.name === item.Title);

        if (!categoryNode) {
            categoryNode = mapCategoryNode(item);
            phaseNode.nodes.push(categoryNode);
        }
    }

    phasesQnA.forEach(phaseQnA => {
        const stage = nodes.find(node => node.name === phaseQnA.Stage);
        if (stage) {
            const phase = stage.nodes.find(node => node.name === phaseQnA.Phase);
            if (!phase) {
                stage.nodes = [
                    ...stage.nodes,
                    {
                        name: phaseQnA.Phase,
                        nodes: [],
                        uploadedFiles: [],
                        status: FolderStatus.Default,
                        localization: {
                            [Language.CS]: cleanFolderName(phaseQnA.Phase) ?? '',
                            [Language.EN]: cleanFolderName(phaseQnA.Phase_en) ?? ''
                        }
                    }
                ];
            }
        }
    });

    nodes.forEach(node => {
        node.nodes = sortStructureNodesByName(node.nodes);
    });
    return sortStructureNodesByName(nodes);
};