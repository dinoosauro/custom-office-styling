export default function UpdateGenericProperties(source: any, dest: any, skipKeys: string[] = []) {
    for (const key in source) {
        if (key.startsWith("_")) continue;
        try {
            if (skipKeys.indexOf(key) !== -1) continue;
            if (typeof source[key] !== "object" && source[key] !== dest[key]) {
                dest[key] = source[key];
            }
        } catch(ex) {
            console.warn(key, ex);
        }
    }
}