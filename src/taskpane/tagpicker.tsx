import * as React from "react";

import { Datapoint } from "./datamodel";

type TagPickerProps = {
  concepts: Datapoint[];
  onSelect: (conceptReference: string) => void;
};

const STANDARD_ORDER = ["ESRS2", "E1", "E2", "E3", "E4", "E5", "S1", "S2", "S3", "S4", "G1"];
const ESRS2_DR_ORDER = ["BP-1", "BP-2", "GOV-1", "GOV-2", "GOV-3", "GOV-4", "SBM-1", "SBM-2", "SBM-3", "IRO-1", "IRO-2", "GDR", "GDR-P", "GDR-A", "GDR-M", "GDR-T"];

const STANDARD_ORDER_MAP = new Map(STANDARD_ORDER.map((standard, index) => [standard, index]));
const ESRS2_DR_ORDER_MAP = new Map(ESRS2_DR_ORDER.map((dr, index) => [dr, index]));

function uniqueValues(values: string[]): string[] {
  return Array.from(new Set(values.filter((value) => value.trim().length > 0))).sort(compareStandards);
}

function compareStandards(left: string, right: string): number {
  const leftIndex = STANDARD_ORDER_MAP.get(left);
  const rightIndex = STANDARD_ORDER_MAP.get(right);

  if (leftIndex !== undefined && rightIndex !== undefined) {
    return leftIndex - rightIndex;
  }

  if (leftIndex !== undefined) {
    return -1;
  }

  if (rightIndex !== undefined) {
    return 1;
  }

  return left.localeCompare(right);
}

function compareDrs(standard: string, left: string, right: string): number {
  if (standard === "ESRS2") {
    const leftIndex = ESRS2_DR_ORDER_MAP.get(left);
    const rightIndex = ESRS2_DR_ORDER_MAP.get(right);

    if (leftIndex !== undefined && rightIndex !== undefined) {
      return leftIndex - rightIndex;
    }

    if (leftIndex !== undefined) {
      return -1;
    }

    if (rightIndex !== undefined) {
      return 1;
    }

    return left.localeCompare(right);
  }

  const standardPrefix = `${standard}-`;
  const leftMatch = left.startsWith(standardPrefix) ? Number(left.slice(standardPrefix.length)) : Number.NaN;
  const rightMatch = right.startsWith(standardPrefix) ? Number(right.slice(standardPrefix.length)) : Number.NaN;

  if (!Number.isNaN(leftMatch) && !Number.isNaN(rightMatch)) {
    return leftMatch - rightMatch;
  }

  if (!Number.isNaN(leftMatch)) {
    return -1;
  }

  if (!Number.isNaN(rightMatch)) {
    return 1;
  }

  return left.localeCompare(right);
}

function getHiddenTagSelect(): HTMLSelectElement | null {
  return document.getElementById("tag-select") as HTMLSelectElement | null;
}

function readHiddenTagSelectValue(): string {
  return getHiddenTagSelect()?.value || "";
}

export function TagPicker({ concepts, onSelect }: TagPickerProps) {
  const [selectedStandard, setSelectedStandard] = React.useState<string>("all");
  const [query, setQuery] = React.useState("");
  const [selectedReference, setSelectedReference] = React.useState<string>(readHiddenTagSelectValue());

  React.useEffect(() => {
    const select = getHiddenTagSelect();

    if (!select) {
      return;
    }

    const handleChange = () => {
      setSelectedReference(select.value);
    };

    handleChange();
    select.addEventListener("change", handleChange);

    return () => {
      select.removeEventListener("change", handleChange);
    };
  }, []);

  const standards = React.useMemo(
    () => uniqueValues(concepts.map((concept) => concept.standard)),
    [concepts]
  );

  const filtered = React.useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();

    return concepts.filter((concept) => {
      if (selectedStandard !== "all" && concept.standard !== selectedStandard) {
        return false;
      }

      if (!normalizedQuery) {
        return true;
      }

      return [
        concept.reference,
        concept.standard,
        concept.dr,
        concept.name,
        concept.qname,
      ]
        .filter((value): value is string => typeof value === "string")
        .some((value) => value.toLowerCase().includes(normalizedQuery));
    });
  }, [concepts, query, selectedStandard]);

  const byDr = React.useMemo(() => {
    const map = new Map<string, { standard: string; label: string; items: Datapoint[] }>();

    filtered.forEach((concept: Datapoint) => {
      const subgroup = concept.dr.trim() || "Ungrouped";
      const key = selectedStandard === "all"
        ? `${concept.standard}::${subgroup}`
        : subgroup;

      if (!map.has(key)) {
        map.set(key, {
          standard: concept.standard,
          label: subgroup,
          items: [],
        });
      }

      map.get(key)!.items.push(concept);
    });

    return Array.from(map.values()).sort((left, right) => {
      const standardOrder = compareStandards(left.standard, right.standard);

      if (standardOrder !== 0) {
        return standardOrder;
      }

      return compareDrs(left.standard, left.label, right.label);
    });
  }, [filtered]);

  const handleSelectChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const conceptReference = event.target.value;
    setSelectedReference(conceptReference);
    onSelect(conceptReference);
  };

  return React.createElement(
    "div",
    { className: "tag-picker" },
    React.createElement(
      "div",
      { className: "tag-picker__filters" },
      React.createElement("label", { htmlFor: "tag-picker-standard" }, "Standard"),
      React.createElement(
        "select",
        {
          id: "tag-picker-standard",
          value: selectedStandard,
          onChange: (event: React.ChangeEvent<HTMLSelectElement>) => setSelectedStandard(event.target.value),
        },
        React.createElement("option", { value: "all" }, "All standards"),
        standards.map((standard) =>
          React.createElement("option", { key: standard, value: standard }, standard)
        )
      ),
      React.createElement("label", { htmlFor: "tag-picker-query" }, "Search"),
      React.createElement("input", {
        id: "tag-picker-query",
        type: "search",
        placeholder: "Search concepts",
        value: query,
        onChange: (event: React.ChangeEvent<HTMLInputElement>) => setQuery(event.target.value),
      })
    ),
    React.createElement(
      "label",
      { className: "tag-picker__label", htmlFor: "tag-picker-select" },
      "Tag datapoint"
    ),
    React.createElement(
      "select",
      {
        id: "tag-picker-select",
        className: "tag-picker__select",
        value: selectedReference,
        onChange: handleSelectChange,
      },
      React.createElement("option", { value: "" }, "-- Select a datapoint --"),
      byDr.map(({ label, items, standard }) =>
        React.createElement(
          "optgroup",
          { key: `${standard}::${label}`, label: label },
          items.map((concept) =>
            React.createElement(
              "option",
              { key: concept.reference, value: concept.reference },
              `${concept.name}`
            )
          )
        )
      )
    ),
    React.createElement(
      "div",
      { className: "tag-picker__results", "aria-hidden": "true" }
    )
  );
}