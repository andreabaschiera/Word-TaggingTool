import * as React from "react";
import { createRoot } from "react-dom/client";

import "./word";
import { taxonomy } from "./datamodel";
import { TagPicker } from "./tagpicker";

function syncSelectedTag(reference: string) {
	const select = document.getElementById("tag-picker-select") as HTMLSelectElement | null;

	if (!select) {
		return;
	}

	select.value = reference;
	select.dispatchEvent(new Event("change", { bubbles: true }));
}

Office.onReady((info) => {
	if (info.host !== Office.HostType.Word) {
		return;
	}

	const mountNode = document.getElementById("tag-picker-root");

	if (!mountNode) {
		throw new Error("Missing tag picker root element.");
	}

	const root = createRoot(mountNode);

	root.render(
		React.createElement(TagPicker, {
			concepts: taxonomy.datapoints,
			onSelect: syncSelectedTag,
		})
	);
});
