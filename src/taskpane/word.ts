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

let flashMessageTimer: number | undefined;


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

  taxonomy.datapoints.forEach((dp) => {
    const option = document.createElement("option");

    option.value = dp.reference;
    option.textContent = dp.name;

    select.appendChild(option);
  });
}

function getSelectedDatapoint(): Datapoint | undefined {
  const select = document.getElementById(
    "tag-select"
  ) as HTMLSelectElement;

  return taxonomy.datapoints.find(dp => dp.reference === select.value);
}

function showFlashMessage(message: string) {
  const flashMessage = document.getElementById("flash-message");

  if (!flashMessage) {
    console.log(message);
    return;
  }

  flashMessage.textContent = message;
  flashMessage.classList.add("is-visible");

  if (flashMessageTimer !== undefined) {
    window.clearTimeout(flashMessageTimer);
  }

  flashMessageTimer = window.setTimeout(() => {
    flashMessage.textContent = "";
    flashMessage.classList.remove("is-visible");
    flashMessageTimer = undefined;
  }, 4000);
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

    case "integer":
    case "decimal":
    case "percent":
    case "monetary":
    case "mass":
    case "volume":
    case "energy":
    case "ghgemissions": {
      const input = document.getElementById(
        "tag-value-number"
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
    if (dim.dimension_mode === "fixed") {

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

      if (!selectedValue) {
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
    ["integer", "decimal", "percent", "monetary", "mass", "volume", "energy", "ghgemissions", "year", "date"].includes(selectedDatapoint.datatype)
  ) &&
  (
    selectedEnumerations === undefined
  )) {
    showFlashMessage("One number/date must be picked.");
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
  typeLine.textContent = `Type: ${details.type}`;

  const textLine = document.createElement("div");
  textLine.textContent = `Text: ${details.text}`;

  body.appendChild(typeLine);
  body.appendChild(textLine);

  if (details.value !== undefined) {

    const labelS = extractLabelsFromValues(details.value);

    const valueLine = document.createElement("div");
    switch (details.type) {
      case "boolean":
        valueLine.textContent = `Choice: ${details.value}`;
        break;
      case "enumeration":
        valueLine.textContent = `Choice: ${labelS}`;
        break;
      case "enumerationset":
        valueLine.classList.add("inspection-details-value--list");

        const choicesLabel = document.createElement("div");
        choicesLabel.className = "inspection-details-list-label";
        choicesLabel.textContent = "Choices:";

        const list = document.createElement("ul");
        list.className = "inspection-details-list";

        labelS.forEach(label => {
          const li = document.createElement("li");
          li.textContent = label;
          list.appendChild(li);
        });

        valueLine.append(choicesLabel, list);
        break;
      case "date":
        valueLine.textContent = `Date selected: ${details.value}`;
        break;
      case "year":
        valueLine.textContent = `Year selected: ${details.value}`;
        break;
      case "integer":
      case "decimal":
      case "percent":
      case "monetary":
      case "mass":
      case "volume":
      case "energy":
      case "ghgemissions":
        valueLine.textContent = `Number inputted: ${details.value}`;
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
        dimensionLine.innerHTML = `Dimension: ${title}<br>Dimension member selected: ${dim.memberValue}`;
      } else {
        dimensionLine.innerHTML = `Dimension: ${title}<br>Dimension member typed: ${dim.memberLabel}`;
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

function loadAnnotationIntoUI(
  annotation: TagAnnotation,
  dp: Datapoint
) {
  const select = document.getElementById(
    "tag-select"
  ) as HTMLSelectElement;

  select.value = dp.reference;

  updateValueSelector();

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
      renderNumberInput(dp.datatype, container);
      break;
    }

  updateDimensionSelector();
  updateOptionalDimensionSelector();
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
    ? dp.dimensions.filter(dimension => dimension.required === true)
    : [];

  dimensions.forEach((dimension) => {
    const heading = document.createElement("div");
    heading.textContent = `Required dimension: ${dimension.dimension_title}`;
    container.appendChild(heading);

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
    ? dp.dimensions.filter(dimension => dimension.required === false)
    : [];

  dimensions.forEach((dimension) => {
    const heading = document.createElement("div");
    heading.textContent = `Optional dimension: ${dimension.dimension_title}`;
    container.appendChild(heading);

    if (dimension.dimension_mode === "fixed") {
      renderExplicitDimension(dimension, container);
    } else {
      renderTypedDimension(dimension.dimension_id, container);
    }
  });
}

function renderSingleChoice(
  dp: Datapoint,
  container: HTMLElement
) {
  const label = document.createElement("div");
  label.textContent = dp.datatype === "boolean" ? "Boolean choice:" : "Enumeration choice:";
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
  label.textContent = "Multiple choices:";
  container.appendChild(label);

  enum_choices.forEach((choice) => {
    const wrapper = document.createElement("div");

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
  type: Datatype,
  container: HTMLElement
) {
  const label = document.createElement("label");

  switch (type) {
    case "integer":
      label.textContent = "Integer:";
      break;

    case "decimal":
      label.textContent = "Decimal:";
      break;

    case "percent":
      label.textContent = "Percentage:";
      break;

    case "monetary":
      label.textContent = "Monetary amount:";
      break;

    case "mass":
      label.textContent = "Mass (in kg):";
      break;

    case "volume":
      label.textContent = "Volume (in L3):";
      break;

    case "energy":
      label.textContent = "Energy (in MWh):";
      break;
    
    case "ghgemissions":
      label.textContent = "GHG emissions (in tCo2eq):";
      break;
  }

  const input = document.createElement("input");
  input.type = "number";
  input.id = "tag-value-number";

  if (type === "integer") {
    input.step = "1";
  } else {
    input.step = "any";
  }

  container.appendChild(label);
  container.appendChild(input);
}

function renderExplicitDimension(
  dimension: Dimension,
  container: HTMLElement
) {
  const members = dimension.dimension_members;
  if (!members) {
    return;
  }

  if (members.length < 5) {
    members.forEach((member) => {
    const wrapper = document.createElement("div");

    const radio = document.createElement("input");

    radio.type = "radio";
    radio.name = `dimension-member_${dimension.dimension_id}`;
    radio.value = member.value;
    radio.id = `dimension-${member.value}`;

    const label = document.createElement("label");

    label.htmlFor = radio.id;
    label.textContent = member.label;

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
  const label = document.createElement("label");

  label.textContent = "Dimension member:";

  const input = document.createElement("input");

  input.type = "text";
  input.id = `dimension-member-input_${dimension_id}`;

  container.appendChild(label);
  container.appendChild(input);
}