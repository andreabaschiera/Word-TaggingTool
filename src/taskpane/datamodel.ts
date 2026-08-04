export type TagType =
  | "textBlock"
  | "string"
  | "boolean"
  | "enumeration"
  | "enumerationSet"
  | "date"
  | "year"
  | "numeric"
  | "decimal"
  | "percentage";

export interface Taxonomy {
  tags: TaxonomyTag[];
}

export interface TaxonomyTag {
  id: string;
  name: string;
  type: TagType;
  values?: TaxonomyValue[];
  dimension?: Dimension;
}

export interface TaxonomyValue {
  id: string;
  label: string;
}

export interface Dimension {
  axis: string;
  mode: "explicit" | "typed";
  members?: DimensionMember[];
}

export interface DimensionMember {
  memberId: string;
  memberLabel: string;
}

export const taxonomy: Taxonomy = {
  tags: [
    {
      id: "text_simple",
      name: "Datapoint representing a textblock",
      type: "textBlock"
    },

    {
      id: "string_simple",
      name: "Datapoint representing a string",
      type: "string",
    },

    {
      id: "boolean_simple",
      name: "Datapoint representing a boolean",
      type: "boolean"
    },

    {
      id: "enum_short",
      name: "Datapoint representing a short enumeration",
      type: "enumeration",
      values: [
        { id: "LOW", label: "Low" },
        { id: "MEDIUM", label: "Medium" },
        { id: "HIGH", label: "High" }
      ]
    },

    {
      id: "enum_long",
      name: "Datapoint representing a long enumeration",
      type: "enumeration",
      values: [
        { id: "LOW", label: "Low" },
        { id: "MEDIUM", label: "Medium" },
        { id: "HIGH", label: "High" },
        { id: "SUPERHIGH", label: "Super High" },
        { id: "SUPERLOW", label: "Super Low" }
      ]
    },

    {
      id: "enum_set",
      name: "Datapoint representing an enumeration Set",
      type: "enumerationSet",
      values: [
        { id: "DE", label: "Germany" },
        { id: "IT", label: "Italy" },
        { id: "FR", label: "France" },
        { id: "PT", label: "Portugal" },
        { id: "ES", label: "Spain" }
      ]
    },

    {
      id: "number_disag-expl",
      name: "Datapoint representing number disaggregated by fixed dim members",
      type: "numeric",
      dimension: {
        axis: "Geography",
        mode: "explicit",
        members: [
          { memberId: "BE", memberLabel: "Belgium" },
          { memberId: "FR", memberLabel: "France" },
          { memberId: "DE", memberLabel: "Germany" }
        ]
      }
    },

    {
      id: "number_disag-typed",
      name: "Datapoint representing number disaggregated by typed dim",
      type: "numeric",
      dimension: {
        axis: "Department",
        mode: "typed"
      }
    },

    {
      id: "enum-short_disag-expl",
      name: "Datapoint representing a short enumeration disaggregated by fixed dim members",
      type: "enumeration",
      values: [
        { id: "LOW", label: "Low" },
        { id: "HIGH", label: "High" }
      ],
      dimension: {
        axis: "Geography",
        mode: "explicit",
        members: [
          { memberId: "BE", memberLabel: "Belgium" },
          { memberId: "DE", memberLabel: "Germany" }
        ]
      }
    },

    {
      id: "enum-set_disag-expl",
      name: "Datapoint representing an enumeration Set disaggregated by fixed dim members",
      type: "enumerationSet",
      values: [
        { id: "LOW", label: "Low" },
        { id: "HIGH", label: "High" },
        { id: "HIGHDRY", label: "High and dry" },
        { id: "SUPERHIGH", label: "Super High" },
        { id: "SUPERLOW", label: "Super Low" }
      ],
      dimension: {
        axis: "Geography",
        mode: "explicit",
        members: [
          { memberId: "BE", memberLabel: "Belgium" },
          { memberId: "DE", memberLabel: "Germany" },
          { memberId: "IT", memberLabel: "Italy" },
          { memberId: "CI", memberLabel: "something else" },
        ]
      }
    },

    {
      id: "text_disag-expl",
      name: "Datapoint representing a text disaggregated by fixed dim members",
      type: "textBlock",
      dimension: {
        axis: "Business unit",
        mode: "explicit",
        members: [
          { memberId: "HR", memberLabel: "Human Resources" },
          { memberId: "IT", memberLabel: "IT" }
        ]
      }
    }
  ]
};