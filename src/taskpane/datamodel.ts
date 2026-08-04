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