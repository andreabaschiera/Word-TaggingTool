export function arraysEqualUnordered(arr1: Array<any>, arr2: Array<any>) {
    if (!Array.isArray(arr1) || !Array.isArray(arr2)) {
        throw new TypeError("Both arguments must be arrays.");
    }
    if (arr1.length !== arr2.length) return false;

    // Create sorted copies (to avoid mutating originals)
    const sorted1 = [...arr1].sort();
    const sorted2 = [...arr2].sort();

    // Compare element by element
    for (let i = 0; i < sorted1.length; i++) {
        if (sorted1[i] !== sorted2[i]) return false;
    }
    return true;
}