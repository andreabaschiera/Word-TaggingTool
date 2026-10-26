import {
  DimensionMemberSelection,
  parseAnnotation,
  serializeAnnotation,
  TagAnnotation
} from "./annotations";

import {
  Dimension,
  Datapoint,
  Datatype,
  ValueLabel,
  taxonomy,
  valueLabelPairs,
  dimensionIDLabelPairs
} from "./datamodel";

import exportData from "./fastexport";

import showFlashMessage from "./flashmessage";


function getTagPickerSelect(): HTMLSelectElement | null {
  return document.getElementById("tag-picker-select") as HTMLSelectElement | null;
}

function bindTagPickerSelection() {
  const attachListener = () => {
    const select = getTagPickerSelect();

    if (!select) {
      return false;
    }

    select.onchange = updateValueSelector;
    return true;
  };

  if (attachListener()) {
    return;
  }

  const observer = new MutationObserver(() => {
    if (attachListener()) {
      observer.disconnect();
    }
  });

  observer.observe(document.body, { childList: true, subtree: true });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";

    bindTagPickerSelection();

    document.getElementById("apply-tag").onclick = applyTag;

    document.getElementById("inspect-tag").onclick = inspectSelectedTag;

    document.getElementById("remove-tag").onclick = removeSelectedTag;

    document.getElementById("export-data").onclick = exportData;
  }
});

function getSelectedDatapoint(): Datapoint | undefined {
  const select = getTagPickerSelect();

  if (!select) {
    return undefined;
  }

  return taxonomy.datapoints.find(dp => dp.reference === select.value);
}

function getSelectedChoice(
  dp: Datapoint
): string | string[] | undefined {

  switch (dp.datatype) {

    case "narrative":
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

    case "enumerationset": {
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
  }
}

function getSelectedDimensionMembers(dp: Datapoint):
  DimensionMemberSelection[] | undefined {

  if (!dp?.dimensions) {
    return undefined;
  }

  let annotated_dimensions: Array<DimensionMemberSelection> = [];
  const dimensions = dp.dimensions;

  dimensions.forEach((dim) => {
    if (dim.dimension_mode === "fixed" || dim.dimension_mode === "default") {

      const selectedRadio = document.querySelector(
        `input[name="dimension-member_${dim.dimension_id}"]:checked`
      ) as HTMLInputElement | null;

      let selectedValue: string | undefined = selectedRadio?.value;

      if (!selectedValue) {
        const selectedDropdown = document.getElementById(
          `dimension-member_${dim.dimension_id}`
        ) as HTMLSelectElement | null;

        if (selectedDropdown && selectedDropdown.value.trim()) {
          selectedValue = selectedDropdown.value;
        }
      }

      // Special handling for year picker (dim44)
      if (!selectedValue && dim.dimension_id === "dim44") {
        const yearInput = document.getElementById(
          `dimension-member-input_${dim.dimension_id}`
        ) as HTMLInputElement | null;

        if (yearInput && yearInput.value.trim()) {
          selectedValue = yearInput.value.trim();
        }
      }

      if (!selectedValue || getDefaultChoices().includes(selectedValue)) {
        return;
      }

      const selected_dim: DimensionMemberSelection = {
        id: dim.dimension_id,
        mode: "fixed",
        memberValue: selectedValue
      };

      annotated_dimensions.push(selected_dim);
      return;
    }


    const input = document.getElementById(
      `dimension-member-input_${dim.dimension_id}`
    ) as HTMLInputElement | null;

    if (!input || !input.value.trim()) {
      return;
    }

    const typed_dim: DimensionMemberSelection = {
      id: dim.dimension_id,
      mode: "typed",
      memberLabel: input.value.trim()
    };

    annotated_dimensions.push(typed_dim);
  });

  return annotated_dimensions;
}

function areRequiredDimensionsPresent(dp: Datapoint, selected_dims?: DimensionMemberSelection[]): string[] {
  if (!dp.dimensions) {
    return [];
  }

  const dimensions = dp.dimensions;
  const required_ids: Array<string> = [];

  dimensions.forEach((dimension) => {
    if (dimension.required === true) {
      required_ids.push(dimension.dimension_id);
    }
  });

  if (!selected_dims) {
    return required_ids;
  }
  const annotated_ids = new Set(selected_dims.map(dim => dim.id));

  return required_ids.filter(item => !annotated_ids.has(item));
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

  const selectedDatapoint = getSelectedDatapoint();

  if (!selectedDatapoint) {
    showFlashMessage("No datapoint selected.");
    return;
  }

  const selectedEnumerations = getSelectedChoice(selectedDatapoint);

  const selectedDimensions = getSelectedDimensionMembers(selectedDatapoint);

  if (
    (
      selectedDatapoint.datatype === "boolean" ||
      selectedDatapoint.datatype === "enumeration" ||
      selectedDatapoint.datatype === "enumerationset"
    ) &&
    (
      selectedEnumerations === undefined ||
      (Array.isArray(selectedEnumerations) && selectedEnumerations.length === 0)
    )
  ) {
    showFlashMessage("At least one enumeration value must be selected.");
    return;
  }
  
  if ((
    ["year", "date"].includes(selectedDatapoint.datatype)
  ) &&
  (
    selectedEnumerations === undefined
  )) {
    showFlashMessage("One year/date must be picked.");
    return;
  }

  const missing_dims = areRequiredDimensionsPresent(selectedDatapoint, selectedDimensions);

  if (missing_dims.length !== 0) {
    console.log(`Misssing dim IDs: ${missing_dims}`)
    const missing_dims_labels = missing_dims.map(dimID => dimensionIDLabelPairs[dimID]);
    console.log(`Misssing dim labels: ${missing_dims_labels}`)
    showFlashMessage(`A dimension member for each of the required dimensions must be selected. Missing choices for dimensions: ${missing_dims_labels.join(", ")}`);
    return;
  }

  return Word.run(async (context) => {
    const range = context.document.getSelection();

    range.load("text");

    await context.sync();

    if (!range.text.trim()) {
      showFlashMessage("No text selected.");
      return;
    }

    const annotation: TagAnnotation = {
      tagId: selectedDatapoint.reference,
      valueChosen: selectedEnumerations,
      dimensions: selectedDimensions
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

    contentControl.title = selectedDatapoint.reference;
    contentControl.tag = serializeAnnotation(annotation);
    contentControl.appearance =
      Word.ContentControlAppearance.boundingBox;
    
    switch (selectedDatapoint.datatype) {
      case "narrative":
        contentControl.color = '#5ab4ec';
        break;
      case "string":
        contentControl.color = '#8bcff7';
        break;
      case "boolean":
        contentControl.color = '#f8bc6e';
        break;
      case "enumeration":
        contentControl.color = '#dd8741';
        break;
      case "enumerationset":
        contentControl.color = '#df6807';
        break;
      case "integer":
      case "decimal":
      case "percent":
        contentControl.color = '#d4a404';
        break;
      case "monetary":
        contentControl.color = '#f0412a';
        break;
      case "mass":
        contentControl.color = '#d92af0';
        break;
      case "volume":
        contentControl.color = '#2d1ce6';
        break;
      case "energy":
        contentControl.color = '#86db24';
        break;
      case "ghgemissions":
        contentControl.color = '#10c246';
        break;
      case "year":
      case "date":
        contentControl.color = '#2ae3f0';
        break;
    }

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
      showFlashMessage("The current selection is not tagged.");
      return;
    }

    const annotation =
      parseAnnotation(control.tag);

    if (!annotation) {
      showFlashMessage("Content control found, but it is not a taxonomy annotation.");
      return;
    }

    const dp = taxonomy.datapoints.find(
      dp => dp.reference === annotation.tagId
    );

    if (!dp) {
      showFlashMessage(`Unknown taxonomy ID: ${annotation.tagId}`);
      return;
    }

    loadAnnotationIntoUI(
      annotation,
      dp
    );

    // console.log(annotation)

    renderInspectionDetails(details, {
      title: dp.name,
      type: dp.datatype,
      text: control.text,
      value: annotation.valueChosen,
      defined_dims: dp.dimensions,
      dimensions: annotation.dimensions ? annotation.dimensions : undefined
    });
  });
}

function renderInspectionDetails(
  container: HTMLElement,
  details: {
    title: string;
    type: Datatype;
    text: string;
    value?: string | string[];
    defined_dims?: Dimension[];
    dimensions?: DimensionMemberSelection[];
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
  appendDescriptorValueLine(typeLine, "Type:", details.type, "descriptor-text--type");

  const textLine = document.createElement("div");
  appendDescriptorValueLine(textLine, "Text:", details.text, "descriptor-text--text");

  body.appendChild(typeLine);
  body.appendChild(textLine);

  if (details.value !== undefined) {

    const labelS = extractLabelsFromValues(details.value);

    const valueLine = document.createElement("div");
    switch (details.type) {
      case "boolean":
        appendDescriptorValueLine(valueLine, "Choice:", details.value, "descriptor-text--choice");
        break;
      case "enumeration":
        appendDescriptorValueLine(valueLine, "Choice:", labelS, "descriptor-text--choice");
        break;
      case "enumerationset":
        valueLine.classList.add("inspection-details-value--list");

        const choicesLabel = document.createElement("div");
        choicesLabel.className = "inspection-details-list-label";
        appendDescriptorSpan(choicesLabel, "Choices", "descriptor-text--choices");

        const list = document.createElement("ul");
        list.className = "inspection-details-list";

        if (typeof labelS === "string") {
          const li = document.createElement("li");
          li.textContent = labelS;
          list.appendChild(li);
        } else {
          labelS.forEach(label => {
          const li = document.createElement("li");
          li.textContent = label;
          list.appendChild(li);
        });
        }

        valueLine.append(choicesLabel, list);
        break;
      case "date":
        appendDescriptorValueLine(valueLine, "Date selected:", details.value, "descriptor-text--date");
        break;
      case "year":
        appendDescriptorValueLine(valueLine, "Year selected:", details.value, "descriptor-text--date");
        break;
      default:
        valueLine.textContent = `There is a problem.`;
    }
    body.appendChild(valueLine);
  }

  if (details.dimensions !== undefined && details.dimensions.length > 0 && details.defined_dims) {
    const dimensionsDiv = document.createElement("div");

    const ids = details.defined_dims.map(dim => dim.dimension_id);
    const titles = details.defined_dims.map(dim => dim.dimension_title);
    const dims: Record<string, string> = Object.fromEntries(
      ids.map((key, index) => [key, titles[index]])
    );

    details.dimensions.forEach(dim => {
      const dimensionLine = document.createElement("p");

      const title = dims[dim.id];

      if (dim.mode == "fixed") {
        appendDescriptorValueLine(dimensionLine, "Dimension:", title, "descriptor-text--dimension");
        dimensionLine.appendChild(document.createElement("br"));
        appendDescriptorValueLine(dimensionLine, "Dimension member selected:", dim.memberValue ?? "", "descriptor-text--member");
      } else {
        appendDescriptorValueLine(dimensionLine, "Dimension:", title, "descriptor-text--dimension");
        dimensionLine.appendChild(document.createElement("br"));
        appendDescriptorValueLine(dimensionLine, "Dimension member typed:", dim.memberLabel ?? "", "descriptor-text--member");
      }

      dimensionsDiv.appendChild(dimensionLine);
    });
    body.appendChild(dimensionsDiv);
  }

  box.appendChild(header);
  box.appendChild(body);
  container.appendChild(box);
}

function clearInspectionDetails(container: HTMLElement) {
  container.innerHTML = "";
}

function extractLabelsFromValues(values: string | string[]): string | string[] {
  if (typeof values === "string") {
    return valueLabelPairs[values];
  }
  if (Array.isArray(values)) {
    return values.map(val => valueLabelPairs[val]);
  }
  throw new Error(`Unexpected type for values: ${typeof values}`);
}

function appendDescriptorSpan(
  container: HTMLElement,
  descriptor: string,
  className?: string
): HTMLSpanElement {
  const span = document.createElement("span");
  span.className = className ? `descriptor-text ${className}` : "descriptor-text";
  span.textContent = `${descriptor}`;
  container.appendChild(span);
  return span;
}

function appendDescriptorValueLine(
  container: HTMLElement,
  descriptor: string,
  value: string | string[],
  className?: string
) {
  appendDescriptorSpan(container, descriptor, className);

  const valueSpan = document.createElement("span");
  valueSpan.className = "value-text";
  valueSpan.textContent = typeof value === "string" ? value : value.join(", ");
  container.appendChild(document.createTextNode(" "));
  container.appendChild(valueSpan);
}

function loadAnnotationIntoUI(
  annotation: TagAnnotation,
  dp: Datapoint
) {
  const select = getTagPickerSelect();

  if (!select) {
    return;
  }

  select.value = dp.reference;
  select.dispatchEvent(new Event("change", { bubbles: true }));

  // Value
  if (annotation.valueChosen !== undefined) {

    const values = Array.isArray(annotation.valueChosen)
      ? annotation.valueChosen
      : [annotation.valueChosen];
    
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

  // Dimensions
  const annotated_dims = annotation.dimensions;
  if (annotated_dims) {
    annotated_dims.forEach(dim => {
      if (dim.mode === "fixed") {
        const radio = document.querySelector(
          `input[name="dimension-member_${dim.id}"][value="${dim.memberValue}"]`
        ) as HTMLInputElement | null;

        if (radio) {
          radio.checked = true;
          return;
        }

        const dropdown = document.getElementById(
          `dimension-member_${dim.id}`
        ) as HTMLSelectElement | null;

        if (dropdown && dim.memberValue) {
          dropdown.value = dim.memberValue;
        }
      }

      if (dim.mode === "typed") {
        const input =
          document.getElementById(
            `dimension-member-input_${dim.id}`
          ) as HTMLInputElement | null;

        if (input) {
          input.value = dim.memberLabel ?? "";
        }
      }
    });
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
      showFlashMessage("No content control selected to remove.")
      return;
    }

    // Keep the text, remove only the content control.
    control.delete(true);

    await context.sync();
  });
}

// the update function for enums here controls also the update for dimensions
function updateValueSelector() {
  const dp = getSelectedDatapoint();
  const container = document.getElementById("enum-container");

  if (!container) {
    return;
  }

  container.innerHTML = "";

  if (!dp) {
    updateDimensionSelector();
    updateOptionalDimensionSelector();
    return;
  }

  switch (dp.datatype) {
    case "narrative":
    case "string":
      // No additional value is required.
      break;

    case "boolean":
    case "enumeration":
      renderSingleChoice(dp, container);
      break;

    case "enumerationset":
      renderMultipleChoice(dp, container);
      break;

    case "date":
      renderDateInput(container);
      break;

    case "year":
      renderYearInput(container);
      break;

    case "integer":
    case "decimal":
    case "percent":
    case "monetary":
    case "mass":
    case "volume":
    case "energy":
    case "ghgemissions":
      renderNumberHelper(dp.datatype, container);
      break;
    }

  updateDimensionSelector();
  updateOptionalDimensionSelector();
  updateDefaultDimensionSelector();
}

function capitalizeFirstLetter(val: string) {
    return String(val).charAt(0).toUpperCase() + String(val).slice(1);
}

function appendDimension(container: HTMLElement, dimTitle: string, type: string) {
  const heading = document.createElement("div");
  heading.className = `choice-heading choice-heading--${type}`;
  const descriptor = document.createElement("span");
  descriptor.className = `descriptor-text descriptor-text--${type}`;
  descriptor.textContent = `${capitalizeFirstLetter(type)} dimension:`;
  const title = document.createElement("span");
  title.className = "value-text";
  title.textContent = ` ${dimTitle}`;
  heading.appendChild(descriptor);
  heading.appendChild(title);
  container.appendChild(heading);
}

function updateDimensionSelector() {
  const dp = getSelectedDatapoint();
  const container = document.getElementById("dimensions-container");

  if (!container) {
    return;
  }

  container.innerHTML = "";

  if (!dp?.dimensions) {
    return;
  }

  const dimensions = Array.isArray(dp.dimensions)
    ? dp.dimensions.filter(dimension => dimension.required === true && dimension.dimension_mode !== "default")
    : [];

  dimensions.forEach((dimension) => {
    appendDimension(container, dimension.dimension_title, "required");

    if (dimension.dimension_mode === "fixed") {
      renderExplicitDimension(dimension, container);
    } else {
      renderTypedDimension(dimension.dimension_id, container);
    }
  });
}

function updateOptionalDimensionSelector() {
  const dp = getSelectedDatapoint();
  const container = document.getElementById("dimensions-container");

  if (!container) {
    return;
  }

  // Do not clear the container here — required dimensions are already rendered

  if (!dp?.dimensions) {
    return;
  }

  const dimensions = Array.isArray(dp.dimensions)
    ? dp.dimensions.filter(dimension => dimension.required === false && dimension.dimension_mode !== "default")
    : [];

  dimensions.forEach((dimension) => {
    appendDimension(container, dimension.dimension_title, "optional");

    if (dimension.dimension_mode === "fixed") {
      renderExplicitDimension(dimension, container);
    } else {
      renderTypedDimension(dimension.dimension_id, container);
    }
  });
}

export function getDefaultChoices(): [string, string, string, string] {
  const default_targetYearsCategory = "current_year";
  const default_targetYear = "2026";
  const default_previousReportingErrors = "currently_stated";
  const default_materialityEUdatapoints = "material";

  return [default_targetYearsCategory, default_targetYear, default_previousReportingErrors, default_materialityEUdatapoints]
}

function updateDefaultDimensionSelector() {
  const dp = getSelectedDatapoint();
  const container = document.getElementById("dimensions-container");

  if (!container) {
    return;
  }

  if (!dp?.dimensions) {
    return;
  }

  const dimensions = Array.isArray(dp.dimensions)
    ? dp.dimensions.filter(dimension => dimension.dimension_mode === "default")
    : [];

  if (dimensions.length === 0) {
    return;
  }

  // Create collapsible wrapper for default dimensions
  const wrapper = document.createElement("div");
  wrapper.className = "default-dimensions-wrapper";

  const toggleButton = document.createElement("button");
  toggleButton.type = "button";
  toggleButton.className = "default-dimensions-toggle";
  toggleButton.textContent = "Defaults ▼";
  toggleButton.setAttribute("aria-expanded", "false");

  const content = document.createElement("div");
  content.className = "default-dimensions-content";
  content.style.display = "none";

  toggleButton.onclick = () => {
    const isExpanded = content.style.display !== "none";
    content.style.display = isExpanded ? "none" : "block";
    toggleButton.setAttribute("aria-expanded", isExpanded ? "false" : "true");
    toggleButton.textContent = isExpanded ? "Defaults ▼" : "Defaults ▲";
  };

  dimensions.forEach((dimension) => {
    if (dimension.dimension_id === "dim44") {
      // Special handling for years dimension
      renderYearPickerDimension(dimension, content);
    } else {
      // Regular horizontal radio buttons with smaller font
      renderDefaultDimension(dimension, content);
    }
  });

  wrapper.appendChild(toggleButton);
  wrapper.appendChild(content);
  container.appendChild(wrapper);
}

function renderDefaultDimension(
  dimension: Dimension,
  container: HTMLElement
) {
  const members = dimension.dimension_members;
  if (!members || members.length === 0) {
    return;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "default-dimension-item";

  const label = document.createElement("div");
  label.className = "default-dimension-label";
  label.textContent = dimension.dimension_title + ":";

  const optionsContainer = document.createElement("div");
  optionsContainer.className = "default-dimension-options";

  members.forEach((member) => {
    const radioWrapper = document.createElement("label");
    radioWrapper.className = "default-dimension-radio";

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = `dimension-member_${dimension.dimension_id}`;
    radio.value = member.value;

    // Check default values
    if (getDefaultChoices().includes(member.value)) {
      radio.checked = true;
    }

    radioWrapper.appendChild(radio);
    radioWrapper.appendChild(document.createTextNode(member.label));
    optionsContainer.appendChild(radioWrapper);
  });

  wrapper.appendChild(label);
  wrapper.appendChild(optionsContainer);
  container.appendChild(wrapper);
}

function renderYearPickerDimension(
  dimension: Dimension,
  container: HTMLElement
) {
  const wrapper = document.createElement("div");
  wrapper.className = "default-dimension-item";

  const label = document.createElement("div");
  label.className = "default-dimension-label";
  label.textContent = dimension.dimension_title + ":";

  const inputWrapper = document.createElement("div");
  inputWrapper.className = "default-dimension-input-wrapper";

  const input = document.createElement("input");
  input.type = "number";
  input.id = `dimension-member-input_${dimension.dimension_id}`;
  input.min = "1900";
  input.max = "2100";
  input.step = "1";
  input.value = "2026"
  input.className = "default-dimension-year-input";
  input.placeholder = "Enter year";
  input.value = new Date().getFullYear().toString(); // Default to current year

  inputWrapper.appendChild(input);
  wrapper.appendChild(label);
  wrapper.appendChild(inputWrapper);
  container.appendChild(wrapper);
}

function renderSingleChoice(
  dp: Datapoint,
  container: HTMLElement
) {
  const label = document.createElement("div");
  label.className = "choice-heading";
  appendDescriptorSpan(label, dp.datatype === "boolean" ? "Boolean choice:" : "Enumeration choice:");
  container.appendChild(label);

  if (dp.datatype === "boolean") {
    if (!dp.enum) {
      dp = {
        ...dp,
        enum: {
          enum_id: "enum_bool",
          enum_members: [
            { value: "True", label: "Yes" },
            { value: "False", label: "No" }
          ]
        }
      };
    }
  } else if (!dp.enum) {
    return;
  }

  const enum_choices = dp.enum?.enum_members;
  if (!enum_choices) {
    return;
  }

  if (dp.datatype === "enumeration" && enum_choices.length > 4) {
    renderDropdown(enum_choices, container);
    return;
  }

  enum_choices.forEach((choice) => {
    const wrapper = document.createElement("div");
    wrapper.className = "choice-row";

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "tag-value";
    radio.value = choice.value;
    radio.id = `value-${choice.value}`;

    const valueLabel = document.createElement("label");
    valueLabel.htmlFor = radio.id;
    valueLabel.textContent = choice.label;

    wrapper.appendChild(radio);
    wrapper.appendChild(valueLabel);

    container.appendChild(wrapper);
  });
}

function renderDropdown(
  choices: ValueLabel[],
  container: HTMLElement,
  dimension_id?: string,
) {
  const select = document.createElement("select");

  if (dimension_id) {
    select.id = `dimension-member_${dimension_id}`
  } else {
    select.id = "tag-value-select";
  }

  const placeholder = document.createElement("option");
  placeholder.value = "";
  const text_content = dimension_id ? "-- Select an enumeration choice --" : "-- Select a member --";
  placeholder.textContent = text_content;

  select.appendChild(placeholder);

  choices.forEach((choice) => {
    const option = document.createElement("option");

    option.value = choice.value;
    option.textContent = choice.label;

    select.appendChild(option);
  });

  container.appendChild(select);
}

function renderMultipleChoice(
  dp: Datapoint,
  container: HTMLElement
) {
  const enum_choices = dp.enum?.enum_members;
  if (!enum_choices) {
    return;
  }

  const label = document.createElement("div");
  label.className = "choice-heading";
  appendDescriptorSpan(label, "Multiple choices:");
  container.appendChild(label);

  enum_choices.forEach((choice) => {
    const wrapper = document.createElement("div");
    wrapper.className = "choice-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "tag-values";
    checkbox.value = choice.value;
    checkbox.id = `value-${choice.value}`;

    const valueLabel = document.createElement("label");
    valueLabel.htmlFor = checkbox.id;
    valueLabel.textContent = choice.label;

    wrapper.appendChild(checkbox);
    wrapper.appendChild(valueLabel);

    container.appendChild(wrapper);
  });
}

function renderDateInput(container: HTMLElement) {
  const row = document.createElement("div");
  row.className = "field-row field-row--compact";

  const label = document.createElement("label");
  appendDescriptorSpan(label, "Date:");

  const input = document.createElement("input");
  input.type = "date";
  input.id = "tag-value-date";
  input.className = "field-input";

  label.appendChild(document.createTextNode(" "));
  row.appendChild(label);
  row.appendChild(input);
  container.appendChild(row);
}

function renderYearInput(container: HTMLElement) {
  const row = document.createElement("div");
  row.className = "field-row field-row--compact";

  const label = document.createElement("label");
  appendDescriptorSpan(label, "Year:");

  const input = document.createElement("input");
  input.type = "number";
  input.id = "tag-value-year";
  input.min = "1";
  input.max = "9999";
  input.step = "1";
  input.className = "field-input";

  label.appendChild(document.createTextNode(" "));
  row.appendChild(label);
  row.appendChild(input);
  container.appendChild(row);
}

function renderNumberHelper(
  type: Datatype,
  container: HTMLElement
) {
  const row = document.createElement("div");
  row.className = "field-row field-row--compact";

  const helper = document.createElement("p");
  const classHelper = "descriptor-text--helper"
  helper.style.margin = "0px 0 8px";

  switch (type) {
    case "integer":
      appendDescriptorSpan(helper, "Integer type", classHelper);
      break;

    case "decimal":
      appendDescriptorSpan(helper, "Decimal type", classHelper);
      break;

    case "percent":
      appendDescriptorSpan(helper, "Percentage type", classHelper);
      break;

    case "monetary":
      appendDescriptorSpan(helper, "Monetary type", classHelper);
      break;

    case "mass":
      appendDescriptorSpan(helper, "Mass type", classHelper);
      break;

    case "volume":
      appendDescriptorSpan(helper, "Volume type", classHelper);
      break;

    case "energy":
      appendDescriptorSpan(helper, "Energy type", classHelper);
      break;
    
    case "ghgemissions":
      appendDescriptorSpan(helper, "GHG emissions type", classHelper);
      break;
  }

  row.appendChild(helper);
  container.appendChild(row);
}

function renderExplicitDimension(
  dimension: Dimension,
  container: HTMLElement,
  is_default: boolean = false
) {
  const members = dimension.dimension_members;
  if (!members) {
    return;
  }

  if (members.length < 5 || is_default) {
    members.forEach((member) => {
      const wrapper = document.createElement("div");
      wrapper.className = "choice-row";

      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `dimension-member_${dimension.dimension_id}`;
      radio.value = member.value;
      radio.id = `dimension-${member.value}`;

      const label = document.createElement("label");
      label.htmlFor = radio.id;
      label.textContent = member.label;

      // when there are default dimensions, these values are checked by default
      if (member.value == "current_year" || member.value == "_2026" || member.value == "currently_stated") {
        radio.checked = true;
      }

      wrapper.appendChild(radio);
      wrapper.appendChild(label);

      container.appendChild(wrapper);
    });
  } else {
    renderDropdown(members, container, dimension.dimension_id);
  }
}

function renderTypedDimension(
  dimension_id: string,
  container: HTMLElement
) {
  const row = document.createElement("div");
  row.className = "field-row field-row--compact";

  const label = document.createElement("label");
  appendDescriptorSpan(label, "Member to type");

  const input = document.createElement("input");
  input.type = "text";
  input.id = `dimension-member-input_${dimension_id}`;
  input.className = "field-input";

  label.appendChild(document.createTextNode(" "));
  row.appendChild(label);
  row.appendChild(input);
  container.appendChild(row);
}