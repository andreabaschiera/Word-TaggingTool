import {
  parseAnnotation,
  TagAnnotation
} from "./annotations";

import { taxonomy, Datatype } from "./datamodel";

import showFlashMessage from "./flashmessage";


// build fast lookup dict
const taxonomyMap = new Map(
    taxonomy.datapoints.map(({ reference, ...rest }) => [reference, rest])
);


// data model of json export
interface ExportFile {
    sourceHost: string;
    timestamp: string;
    tagCount?: number;
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

    Word.run(async (context) => {
        const { contentControls } = context.document // object destructuring

        contentControls.load('items');

        let tagIndex = 1
        const seenTagKeys = new Set<string>();

        await context.sync();
            
        for (const contentControl of contentControls.items) {
            const annotation: TagAnnotation | undefined = parseAnnotation(contentControl.tag)
            // skip other content controls
            if (annotation === undefined) {
                continue;
            }

            const currentTaxonomyFields = taxonomyMap.get(annotation.tagId);
            if (!currentTaxonomyFields) {
                showFlashMessage(`Invalid datapoint: '${annotation.tagId}'`);
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
                if (isNaN(Number(text))) {
                    showFlashMessage(`'${text}' is not a valid number`);
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
                            showFlashMessage(`'${dim.id}' is an invalid dimension for datapoint '${annotation.tagId}'`);
                            return;
                        }

                        const title = (dimFields) ? dimFields.dimension_title : undefined;
                        const member = (dim.memberValue) ? dim.memberValue : ((dim.memberLabel) ? dim.memberLabel : undefined);

                        if (title && member) {
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
                showFlashMessage(`Duplicate tag detected for ${tag.taxonomyFields.conceptId}`);
                return;
            }

            seenTagKeys.add(uniquenessKey);
            exportData.tags[`tag${tagIndex}`] = tag;

            tagIndex += 1;
        }

        const tagCount = Object.keys(exportData.tags).length;
            
        if (tagCount === 0) {
            showFlashMessage("No tags have been applied yet.")
            return;
        }

        exportData.tagCount = tagCount;

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