export interface TagAnnotation {
  tagId: string;
  valueChosen?: string | string[];
  dimensions?: DimensionMemberSelection[];
}

export interface DimensionMemberSelection {
  id: string;
  mode: "fixed" | "typed";
  memberValue?: string;
  memberLabel?: string;
}
// explicit -> memberValue required (users select it - and thus we have the techn value)
// typed -> memberLabel required (users type it - and thus we only have the label)

export function serializeAnnotation(
  annotation: TagAnnotation
): string {
  const parts: string[] = [];

  parts.push(`taxonomyId=${encodeURIComponent(annotation.tagId)}`);

  if (annotation.valueChosen !== undefined) {
    const value = Array.isArray(annotation.valueChosen)
      ? annotation.valueChosen.join(",")
      : annotation.valueChosen;

    parts.push(`valueChosen=${encodeURIComponent(value)}`);
  }

  const dimensions = annotation.dimensions;

  dimensions?.forEach((dimension) => {
    if (dimension.mode === "fixed" && dimension.memberValue) {
      parts.push(
        `${dimension.id}_memberValue=${encodeURIComponent(
          dimension.memberValue
        )}`
      );
    }

    if (dimension.mode === "typed" && dimension.memberLabel) {
      parts.push(
        `${dimension.id}_memberLabel=${encodeURIComponent(
          dimension.memberLabel
        )}`
      );
    }
  });

  return parts.join("|");
}

export function parseAnnotation(
  metadata: string
): TagAnnotation | undefined {

  const fields: Record<string, string> = {};

  for (const part of metadata.split("|")) {
    const separatorIndex = part.indexOf("=");

    if (separatorIndex === -1) {
      continue;
    }

    const key = part.substring(0, separatorIndex);
    const value = part.substring(separatorIndex + 1);

    fields[key] = decodeURIComponent(value);
  }

  if (!fields.taxonomyId) {
    return undefined;
  }

  const annotation: TagAnnotation = {
    tagId: fields.taxonomyId
  };

  if (fields.valueChosen !== undefined) {
    annotation.valueChosen = fields.valueChosen.includes(",")
      ? fields.valueChosen.split(",")
      : fields.valueChosen;
  }

  let list_dimensions: Array<DimensionMemberSelection> = []
  for (const key in fields) {
    if (key.includes("memberValue")) {
      const dimension_id = key.split("_")[0];
      const dim: DimensionMemberSelection = {
        id: dimension_id,
        mode: "fixed",
        memberValue: fields[key],
      }
      list_dimensions.push(dim);
    } else if (key.includes("memberLabel")) {
      const dimension_id = key.split("_")[0];
      const dim: DimensionMemberSelection = {
        id: dimension_id,
        mode: "typed",
        memberLabel: fields[key],
      }
      list_dimensions.push(dim);
    }
  }

  if (list_dimensions.length !== 0) {
    annotation.dimensions = list_dimensions;
  }

  return annotation;
}