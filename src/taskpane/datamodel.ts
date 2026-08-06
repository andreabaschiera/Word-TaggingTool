export type Datatype =
  | "enumeration"
  | "narrative"
  | "boolean"
  | "enumerationset"
  | "percent"
  | "monetary"
  | "date"
  | "string"
  | "undefined"
  | "year"
  | "ghgemissions"
  | "decimal"
  | "energy"
  | "mass"
  | "volume"
  | "integer";

export interface Taxonomy {
  datapoints: Datapoint[];
}

export interface Datapoint {
  reference: string;
  dr: string;
  datatype: Datatype;
  name: string;
  qname: string;
  enum?: Enumeration;
  dimensions?: Dimension[];
}

export interface ValueLabel {
  value: string;
  label: string;
}

export interface Enumeration {
  enum_id: string;
  enum_members: ValueLabel[];
}

export interface Dimension {
  required: boolean;
  dimension_id: string;
  dimension_mode: "fixed" | "typed";
  dimension_title: string;
  dimension_members?: ValueLabel[];
}

import data from "../python/taxonomy.json";

export const taxonomy: Taxonomy = {datapoints: data as Datapoint[]};


// importing and parsing value-label JSON
import ValueLabels from "../python/ValueLabels.json"
if (!Array.isArray(ValueLabels)) {
  throw new Error('Value-labels JSON must be an array of objects.');
}

export const valueLabelPairs: Record<string, string> = {};
for (const item of ValueLabels) {
  if (typeof item.value !== 'string' || typeof item.label !== 'string') {
      console.warn('Skipping invalid item:', item);
      continue;
  }
  valueLabelPairs[item.value] = item.label;
}

// importing and parsing dimensions' id-label JSON
import dimensionLabels from "../python/dimensionLabels.json"
if (!Array.isArray(dimensionLabels)) {
  throw new Error('Dimension ID-labels JSON must be an array of objects.');
}

export const dimensionIDLabelPairs: Record<string, string> = {};
for (const item of dimensionLabels) {
  if (typeof item.dimension_id !== 'string' || typeof item.dimension_name !== 'string') {
      console.warn('Skipping invalid item:', item);
      continue;
  }
  dimensionIDLabelPairs[item.dimension_id] = item.dimension_name;
}