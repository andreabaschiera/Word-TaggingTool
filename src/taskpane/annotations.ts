export interface TagAnnotation {
  tagId: string;
  value?: string | string[];
  dimensionMember?: DimensionMemberSelection;
}

export interface DimensionMemberSelection {
  mode: "explicit" | "typed";
  memberId?: string;
  memberLabel?: string;
}
// explicit -> memberID required (users select it - and thus we have the id)
// typed -> memberLabel required (users type it - and thus we only have the label)

export function serializeAnnotation(
  annotation: TagAnnotation
): string {
  const parts: string[] = [];

  parts.push(`taxonomyId=${encodeURIComponent(annotation.tagId)}`);

  if (annotation.value !== undefined) {
    const value = Array.isArray(annotation.value)
      ? annotation.value.join(",")
      : annotation.value;

    parts.push(`value=${encodeURIComponent(value)}`);
  }

  if (annotation.dimensionMember) {
    parts.push(
      `dimensionMode=${annotation.dimensionMember.mode}`
    );

    if (
      annotation.dimensionMember.mode === "explicit" &&
      annotation.dimensionMember.memberId
    ) {
      parts.push(
        `memberId=${encodeURIComponent(
          annotation.dimensionMember.memberId
        )}`
      );
    }

    if (
      annotation.dimensionMember.mode === "typed" &&
      annotation.dimensionMember.memberLabel
    ) {
      parts.push(
        `memberLabel=${encodeURIComponent(
          annotation.dimensionMember.memberLabel
        )}`
      );
    }
  }

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

  if (fields.value !== undefined) {
    annotation.value = fields.value.includes(",")
      ? fields.value.split(",")
      : fields.value;
  }

  if (fields.dimensionMode === "explicit") {
    annotation.dimensionMember = {
      mode: "explicit",
      memberId: fields.memberId
    };
  }

  if (fields.dimensionMode === "typed") {
    annotation.dimensionMember = {
      mode: "typed",
      memberLabel: fields.memberLabel
    };
  }

  return annotation;
}