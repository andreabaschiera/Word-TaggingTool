import {
  DimensionMemberSelection,
  parseAnnotation,
  serializeAnnotation,
  TagAnnotation
} from "./annotations";

import {
  Dimension,
  TagType,
  TaxonomyTag,
  TaxonomyValue,
  taxonomy
} from "./datamodel";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";

    populateTagSelector();

    document.getElementById("tag-select").onchange = updateValueSelector;

    document.getElementById("apply-tag").onclick = applyTag;

    document.getElementById("inspect-tag").onclick = inspectSelectedTag;

    document.getElementById("remove-tag").onclick = removeSelectedTag;
  }
});

function populateTagSelector() {
  const select = document.getElementById(
    "tag-select"
  ) as HTMLSelectElement;

  taxonomy.tags.forEach((tag) => {
    const option = document.createElement("option");

    option.value = tag.id;
    option.textContent = tag.name;

    select.appendChild(option);
  });
}

function getSelectedTag(): TaxonomyTag | undefined {
  const select = document.getElementById(
    "tag-select"
  ) as HTMLSelectElement;

  return taxonomy.tags.find(tag => tag.id === select.value);
}

function getSelectedValue(
  tag: TaxonomyTag
): string | string[] | undefined {

  switch (tag.type) {

    case "textBlock":
    case "string":
      return undefined;

    case "boolean":
    case "enumeration": {
      const radio = document.querySelector(
        'input[name="tag-value"]:checked'
      ) as HTMLInputElement | null;

      if (radio) {
        return radio.value;
      }

      const select = document.getElementById(
        "tag-value-select"
      ) as HTMLSelectElement | null;

      if (select && select.value) {
        return select.value;
      }

      return undefined;
    }

    case "enumerationSet": {
      const checkboxes = Array.from(
        document.querySelectorAll(
          'input[name="tag-values"]:checked'
        )
      ) as HTMLInputElement[];

      return checkboxes.map(
        checkbox => checkbox.value
      );
    }

    case "date": {
      const input = document.getElementById(
        "tag-value-date"
      ) as HTMLInputElement | null;

      return input?.value || undefined;
    }

    case "year": {
      const input = document.getElementById(
        "tag-value-year"
      ) as HTMLInputElement | null;

      return input?.value || undefined;
    }

    case "numeric":
    case "decimal":
    case "percentage": {
      const input = document.getElementById(
        "tag-value-number"
      ) as HTMLInputElement | null;

      return input?.value || undefined;
    }
  }
}

function getSelectedDimensionMember(tag: TaxonomyTag):
  DimensionMemberSelection | undefined {

  if (!tag?.dimension) {
    return undefined;
  }

  const dimension = tag.dimension;

  if (dimension.mode === "explicit") {

    const selected = document.querySelector(
      'input[name="dimension-member"]:checked'
    ) as HTMLInputElement | null;

    if (!selected) {
      return undefined;
    }

    return {
      mode: "explicit",
      memberId: selected.value
    };
  }

  const input = document.getElementById(
    "dimension-member-input"
  ) as HTMLInputElement | null;

  if (!input || !input.value.trim()) {
    return undefined;
  }

  return {
    mode: "typed",
    memberLabel: input.value.trim()
  };
}

// async function getSelectedContentControl():
//   Promise<Word.ContentControl | undefined> {

//   return Word.run(async (context) => {

//     const selection = context.document.getSelection();

//     const controls =
//       selection.parentContentControl;

//     controls.load("items/tag,title,text");

//     await context.sync();

//     if (controls.items.length === 0) {
//       return undefined;
//     }

//     return controls.items[0];
//   });
// }

export async function applyTag() {

  const selectedTag = getSelectedTag();

  if (!selectedTag) {
    console.log("No tag selected.");
    return;
  }

  const selectedValue = getSelectedValue(selectedTag);

  const selectedDimension = getSelectedDimensionMember(selectedTag);

  // Values are mandatory only for these types.
  if (
    (
      selectedTag.type === "boolean" ||
      selectedTag.type === "enumeration" ||
      selectedTag.type === "enumerationSet"
    ) &&
    (
      selectedValue === undefined ||
      (Array.isArray(selectedValue) && selectedValue.length === 0)
    )
  ) {
    console.log("A value must be selected.");
    return;
  }

  if (
    selectedTag.dimension &&
    !selectedDimension
  ) {
    console.log("A dimension member must be selected.");
    return;
  }

  return Word.run(async (context) => {
    const range = context.document.getSelection();

    range.load("text");

    await context.sync();

    if (!range.text.trim()) {
      console.log("No text selected.");
      return;
    }

    const annotation: TagAnnotation = {
      tagId: selectedTag.id,
      value: selectedValue,
      dimensionMember: selectedDimension
    };

    // const existingControl = await getSelectedContentControl();

    // if (existingControl) {
    //   existingControl.tag =
    //     serializeAnnotation(annotation);

    //   existingControl.title = selectedTag.name;

    //   await context.sync();

    //   return;
    // }

    const contentControl = range.insertContentControl(
      Word.ContentControlType.richText
    );

    contentControl.title = selectedTag.name;
    contentControl.tag = serializeAnnotation(annotation);
    contentControl.appearance =
      Word.ContentControlAppearance.boundingBox;

    await context.sync();

    console.log("Applied:", serializeAnnotation(annotation));
  });
}

async function inspectSelectedTag() {

  const details =
    document.getElementById("tag-details");

  if (!details) {
    return;
  }

  details.innerHTML = "";

  await Word.run(async (context) => {

    const selection =
      context.document.getSelection();

    const selectedContentControl =
      selection.getContentControls().getFirstOrNullObject();

    selectedContentControl.load("tag,title,text");

    const parentContentControl =
      selection.parentContentControlOrNullObject;

    parentContentControl.load("tag,title,text");

    await context.sync();

    const control = selectedContentControl.isNullObject
      ? parentContentControl.isNullObject
        ? undefined
        : parentContentControl
      : selectedContentControl;

    if (!control) {
      details.textContent =
        "The current selection is not tagged.";

      return;
    }

    const annotation =
      parseAnnotation(control.tag);

    if (!annotation) {
      details.textContent =
        "Content control found, but it is not a taxonomy annotation.";

      return;
    }

    const tag = taxonomy.tags.find(
      t => t.id === annotation.tagId
    );

    if (!tag) {
      details.textContent =
        `Unknown taxonomy ID: ${annotation.tagId}`;

      return;
    }

    loadAnnotationIntoUI(
      annotation,
      tag
    );

    renderInspectionDetails(details, {
      title: tag.name,
      type: tag.type,
      text: control.text,
      value: formatAnnotationValue(annotation.value),
      dimension: formatAnnotationDimension(annotation.dimensionMember)
    });
  });
}

function renderInspectionDetails(
  container: HTMLElement,
  details: {
    title: string;
    type: string;
    text: string;
    value?: string;
    dimension?: string;
  }
) {
  container.innerHTML = "";

  const box = document.createElement("div");
  box.className = "inspection-details-box";

  const header = document.createElement("div");
  header.className = "inspection-details-header";

  const title = document.createElement("strong");
  title.textContent = details.title;

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "inspection-details-close";
  closeButton.setAttribute("aria-label", "Close details");
  closeButton.textContent = "X";
  closeButton.onclick = () => clearInspectionDetails(container);

  header.appendChild(title);
  header.appendChild(closeButton);

  const body = document.createElement("div");
  body.className = "inspection-details-body";

  const typeLine = document.createElement("div");
  typeLine.textContent = `Type: ${details.type}`;

  const textLine = document.createElement("div");
  textLine.textContent = `Text: ${details.text}`;

  body.appendChild(typeLine);
  body.appendChild(textLine);

  if (details.value !== undefined) {
    const valueLine = document.createElement("div");
    valueLine.textContent = `Value: ${details.value}`;
    body.appendChild(valueLine);
  }

  if (details.dimension !== undefined) {
    const dimensionLine = document.createElement("div");
    dimensionLine.textContent = `Dimension: ${details.dimension}`;
    body.appendChild(dimensionLine);
  }

  box.appendChild(header);
  box.appendChild(body);
  container.appendChild(box);
}

function clearInspectionDetails(container: HTMLElement) {
  container.innerHTML = "";
}

function formatAnnotationValue(
  value: TagAnnotation["value"]
): string | undefined {

  if (value === undefined) {
    return undefined;
  }

  return Array.isArray(value)
    ? value.join(", ")
    : value;
}

function formatAnnotationDimension(
  dimensionMember: TagAnnotation["dimensionMember"]
): string | undefined {

  if (!dimensionMember) {
    return undefined;
  }

  if (dimensionMember.mode === "explicit") {
    return `Dimension member: ${dimensionMember.memberId}`;
  }

  return `Dimension member: ${dimensionMember.memberLabel}`;
}

function loadAnnotationIntoUI(
  annotation: TagAnnotation,
  tag: TaxonomyTag
) {
  const select = document.getElementById(
    "tag-select"
  ) as HTMLSelectElement;

  select.value = tag.id;

  updateValueSelector();

  // Value
  if (annotation.value !== undefined) {

    const values = Array.isArray(annotation.value)
      ? annotation.value
      : [annotation.value];
    
    // attribute selector
    document
      .querySelectorAll(
        'input[name="tag-value"]'
      )
      .forEach(element => {
        const input = element as HTMLInputElement;

        input.checked =
          values.includes(input.value);
      });

    const dropdown =
      document.getElementById(
        "tag-value-select"
      ) as HTMLSelectElement | null;

    if (dropdown) {
      dropdown.value = values[0];
    }

    document
      .querySelectorAll(
        'input[name="tag-values"]'
      )
      .forEach(element => {
        const input = element as HTMLInputElement;

        input.checked =
          values.includes(input.value);
      });
  }

  // Dimension
  if (annotation.dimensionMember) {

    if (
      annotation.dimensionMember.mode === "explicit"
    ) {
      const radio = document.querySelector(
        `input[name="dimension-member"][value="${annotation.dimensionMember.memberId}"]`
      ) as HTMLInputElement | null;

      if (radio) {
        radio.checked = true;
      }
    }

    if (
      annotation.dimensionMember.mode === "typed"
    ) {
      const input =
        document.getElementById(
          "dimension-member-input"
        ) as HTMLInputElement | null;

      if (input) {
        input.value =
          annotation.dimensionMember.memberLabel ?? "";
      }
    }
  }
}

async function removeSelectedTag() {

  await Word.run(async (context) => {

    const selection =
      context.document.getSelection();

    const selectedContentControl =
      selection.getContentControls().getFirstOrNullObject();

    selectedContentControl.load("tag");

    const parentContentControl =
      selection.parentContentControlOrNullObject;

    parentContentControl.load("tag");

    await context.sync();

    const control = selectedContentControl.isNullObject
      ? parentContentControl.isNullObject
        ? undefined
        : parentContentControl
      : selectedContentControl;

    if (!control) {
      return;
    }

    // Keep the text, remove only the content control.
    control.delete(true);

    await context.sync();
  });
}

function updateValueSelector() {
  const tag = getSelectedTag();
  const container = document.getElementById("value-container");

  if (!container) {
    return;
  }

  container.innerHTML = "";

  if (!tag) {
    updateDimensionSelector();
    return;
  }

  switch (tag.type) {
    case "textBlock":
    case "string":
      // No additional value is required.
      break;

    case "boolean":
    case "enumeration":
      renderSingleChoice(tag, container);
      break;

    case "enumerationSet":
      renderMultipleChoice(tag, container);
      break;

    case "date":
      renderDateInput(container);
      break;

    case "year":
      renderYearInput(container);
      break;

    case "numeric":
    case "decimal":
    case "percentage":
      renderNumberInput(tag.type, container);
      break;
  }

  updateDimensionSelector();
}

function updateDimensionSelector() {
  const tag = getSelectedTag();
  const container = document.getElementById("dimension-container");

  if (!container) {
    return;
  }

  container.innerHTML = "";

  if (!tag?.dimension) {
    return;
  }

  const dimension = tag.dimension;

  const heading = document.createElement("div");
  heading.textContent = `Dimension: ${dimension.axis}`;
  container.appendChild(heading);

  if (dimension.mode === "explicit") {
    renderExplicitDimension(dimension, container);
  } else {
    renderTypedDimension(container);
  }
}

function renderSingleChoice(
  tag: TaxonomyTag,
  container: HTMLElement
) {
  if (!tag.values) {
    return;
  }

  const label = document.createElement("div");
  label.textContent = "Value:";
  container.appendChild(label);

  if (tag.type === "boolean") {
    tag = {
      ...tag,
      values: [
        { id: "YES", label: "Yes" },
        { id: "NO", label: "No" }
      ]
    };
  }

  if (tag.type === "enumeration" && tag.values.length > 4) {
    renderDropdown(tag.values, container);
    return;
  }

  tag.values.forEach((value) => {
    const wrapper = document.createElement("div");

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "tag-value";
    radio.value = value.id;
    radio.id = `value-${value.id}`;

    const valueLabel = document.createElement("label");
    valueLabel.htmlFor = radio.id;
    valueLabel.textContent = value.label;

    wrapper.appendChild(radio);
    wrapper.appendChild(valueLabel);

    container.appendChild(wrapper);
  });
}

function renderDropdown(
  values: TaxonomyValue[],
  container: HTMLElement
) {
  const select = document.createElement("select");
  select.id = "tag-value-select";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "-- Select a value --";

  select.appendChild(placeholder);

  values.forEach((value) => {
    const option = document.createElement("option");

    option.value = value.id;
    option.textContent = value.label;

    select.appendChild(option);
  });

  container.appendChild(select);
}

function renderMultipleChoice(
  tag: TaxonomyTag,
  container: HTMLElement
) {
  if (!tag.values) {
    return;
  }

  const label = document.createElement("div");
  label.textContent = "Values:";
  container.appendChild(label);

  tag.values.forEach((value) => {
    const wrapper = document.createElement("div");

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "tag-values";
    checkbox.value = value.id;
    checkbox.id = `value-${value.id}`;

    const valueLabel = document.createElement("label");
    valueLabel.htmlFor = checkbox.id;
    valueLabel.textContent = value.label;

    wrapper.appendChild(checkbox);
    wrapper.appendChild(valueLabel);

    container.appendChild(wrapper);
  });
}

function renderDateInput(container: HTMLElement) {
  const label = document.createElement("label");
  label.textContent = "Date:";

  const input = document.createElement("input");
  input.type = "date";
  input.id = "tag-value-date";

  container.appendChild(label);
  container.appendChild(input);
}

function renderYearInput(container: HTMLElement) {
  const label = document.createElement("label");
  label.textContent = "Year:";

  const input = document.createElement("input");
  input.type = "number";
  input.id = "tag-value-year";
  input.min = "1";
  input.max = "9999";
  input.step = "1";

  container.appendChild(label);
  container.appendChild(input);
}

function renderNumberInput(
  type: TagType,
  container: HTMLElement
) {
  const label = document.createElement("label");

  switch (type) {
    case "numeric":
      label.textContent = "Number:";
      break;

    case "decimal":
      label.textContent = "Decimal:";
      break;

    case "percentage":
      label.textContent = "Percentage:";
      break;
  }

  const input = document.createElement("input");
  input.type = "number";
  input.id = "tag-value-number";

  if (type === "numeric") {
    input.step = "1";
  } else if (type === "decimal") {
    input.step = "any";
  } else if (type === "percentage") {
    input.step = "any";
  }

  container.appendChild(label);
  container.appendChild(input);
}

function renderExplicitDimension(
  dimension: Dimension,
  container: HTMLElement
) {
  if (!dimension.members) {
    return;
  }

  dimension.members.forEach((member) => {
    const wrapper = document.createElement("div");

    const radio = document.createElement("input");

    radio.type = "radio";
    radio.name = "dimension-member";
    radio.value = member.memberId;
    radio.id = `dimension-${member.memberId}`;

    const label = document.createElement("label");

    label.htmlFor = radio.id;
    label.textContent = member.memberLabel;

    wrapper.appendChild(radio);
    wrapper.appendChild(label);

    container.appendChild(wrapper);
  });
}

function renderTypedDimension(
  container: HTMLElement
) {
  const label = document.createElement("label");

  label.textContent = "Dimension member:";

  const input = document.createElement("input");

  input.type = "text";
  input.id = "dimension-member-input";

  container.appendChild(label);
  container.appendChild(input);
}