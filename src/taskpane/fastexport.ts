import {
  parseAnnotation,
  TagAnnotation
} from "./annotations";

import { taxonomy, Datatype } from "./datamodel";
import showFlashMessage from "./flashmessage";
import { getDefaultChoices } from "./word";



// build fast lookup dict
const taxonomyMap = new Map(
    taxonomy.datapoints.map(({ reference, ...rest }) => [reference, rest])
);


// data model of json export
interface ExportFile {
    sourceHost: string;
    timestamp: string;
    tagCount?: number;
    validationErrors?: string[];
    tags: {
        [tagName: string]: Tag;
    };
}

interface Tag {
    text?: string;
    value?: string | string[];
    dimensions?: Record<string, Dimension>;
    contentControlId: number;
    taxonomyFields: TaxonomyMetadata;
}

interface Dimension {
    name: string;
    member: string;
}

interface TaxonomyMetadata {
    conceptId: string;
    name: string;
    qname: string;
    standard: string;
    DR: string;
    type: Datatype;
}

const numericTypes = [
  "percent",
  "monetary",
  "ghgemissions",
  "decimal",
  "energy",
  "mass",
  "volume",
  "integer",
] as const

function getDimensionSignature(dimensions?: Record<string, Dimension>) {
    if (!dimensions || Object.keys(dimensions).length === 0) {
        return "";
    }

    return Object.entries(dimensions)
        .sort(([leftId], [rightId]) => leftId.localeCompare(rightId))
        .map(([dimensionId, dimension]) => `${dimensionId}:${dimension.member}`)
        .join("|");
}

function getTagUniquenessKey(tag: Tag) {
    return `${tag.taxonomyFields.conceptId}|${getDimensionSignature(tag.dimensions)}`;
}
 

export default function Export() {
    const exportData: ExportFile = {
        sourceHost: window.location.hostname,
        timestamp: new Date().toISOString(),
        tags: {}
    };
    const validationErrors: string[] = [];

    Word.run(async (context) => {
        const { contentControls } = context.document // object destructuring

        contentControls.load('items');

        let tagIndex = 1
        let annotatedTagCount = 0;
        const seenTagKeys = new Set<string>();

        await context.sync();
            
        for (const contentControl of contentControls.items) {
            const annotation: TagAnnotation | undefined = parseAnnotation(contentControl.tag)
            console.log(annotation);
            // skip other content controls
            if (annotation === undefined) {
                continue;
            }
            annotatedTagCount += 1;

            const currentTaxonomyFields = taxonomyMap.get(annotation.tagId);
            if (!currentTaxonomyFields) {
                validationErrors.push(`Invalid datapoint: '${annotation.tagId}'`);
                continue;
            }

            const tag: Tag = {
                contentControlId: contentControl.id,
                taxonomyFields: {
                    conceptId: annotation.tagId,
                    name: currentTaxonomyFields.name,
                    qname: currentTaxonomyFields.qname,
                    standard: currentTaxonomyFields.standard,
                    DR: currentTaxonomyFields.dr,
                    type: currentTaxonomyFields.datatype,
                }
            }

            // adding CC text to value only when it's a number 
            const text = contentControl.text.trim()
            if ((numericTypes as readonly string[]).includes(tag.taxonomyFields.type)) {
                const numericValue = Number(text);
                if (text === "" || !Number.isFinite(numericValue)) {
                    validationErrors.push(`'${text}' is not a valid number for '${annotation.tagId}'`);
                    continue;
                }
                tag.value = text;
            } else {
                tag.text = contentControl.text.trim()
            }

            // adding choices if applicable
            const val = annotation.valueChosen
            if (val) {
                tag.value = val;
            }

            // adding dimensions if applicable
            if (annotation.dimensions) {
                const dims = annotation.dimensions;

                const allDimensions = currentTaxonomyFields.dimensions;
                if (allDimensions) {

                    const dimMap = new Map(
                        allDimensions.map(({ dimension_id, ...rest }) => [dimension_id, rest])
                    );

                    dims.forEach(dim => {
                        const dimFields = dimMap.get(dim.id);
                        if(!dimFields) {
                            validationErrors.push(`'${dim.id}' is an invalid dimension for datapoint '${annotation.tagId}'`);
                            return;
                        }

                        const title = (dimFields) ? dimFields.dimension_title : undefined;
                        const member = (dim.memberValue) ? dim.memberValue : ((dim.memberLabel) ? dim.memberLabel : undefined);

                        if (title && member) {
                            if (dim.mode === "default" && getDefaultChoices().includes(member)) {
                                return;
                            }
                            tag.dimensions ??= {};
                            tag.dimensions[dim.id] = {
                                name: title,
                                member: member
                            }
                        }
                    });
                }
            }

            const uniquenessKey = getTagUniquenessKey(tag);
            if (seenTagKeys.has(uniquenessKey)) {
                validationErrors.push(`Duplicate tag detected for ${tag.taxonomyFields.conceptId}`);
                continue;
            }

            seenTagKeys.add(uniquenessKey);
            exportData.tags[`tag${tagIndex}`] = tag;

            tagIndex += 1;
        }

        const tagCount = Object.keys(exportData.tags).length;
            
        if (tagCount === 0) {
            if (annotatedTagCount > 0 && validationErrors.length > 0) {
                console.error("Export could not produce any valid tags:", validationErrors);
                showFlashMessage(
                    `No valid tags could be exported. Found ${validationErrors.length} validation error(s). See the console for details.`
                );
            } else {
                showFlashMessage("No tags have been applied yet.")
            }
            return;
        }

        exportData.tagCount = tagCount;
        if (validationErrors.length > 0) {
            exportData.validationErrors = validationErrors;
            console.error("Export completed with validation errors:", validationErrors);
            showFlashMessage(
                `Export completed with ${validationErrors.length} validation error(s). See the downloaded file and console for details.`
            );
        }

        // turns object into JSON string, and then into something the browser can treat as a file (blob)
        const blob = new Blob([JSON.stringify(exportData, null, 2)], {type: 'application/json'});

        // creates temporary URL pointing to file
        const url = URL.createObjectURL(blob);

        // creates invisible link to file and assign filename
        const a = document.createElement("a");
        a.href = url;
        const timestamp = new Date()
            .toISOString()
            .replace(/[:.]/g, "-");
        a.download = `taxonomy-export-${timestamp}.json`;

        // triggers download
        a.click();

        // close temporary url
        URL.revokeObjectURL(url);
        });
}