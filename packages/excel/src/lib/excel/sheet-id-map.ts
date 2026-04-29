/* global Office */

const SETTINGS_KEY_MAP = "openexcel-sheet-id-map";
const SETTINGS_KEY_COUNTER = "openexcel-sheet-id-counter";

interface SheetIdMap {
  [guid: string]: number;
}

let cachedMap: SheetIdMap | null = null;
let cachedCounter: number | null = null;
let isDirty = false;

async function loadFromSettings(): Promise<void> {
  console.log("[sheet-id-map] Loading from settings...");
  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      console.warn("[sheet-id-map] loadFromSettings timed out after 5s");
      cachedMap = cachedMap || {};
      cachedCounter = cachedCounter || 0;
      resolve();
    }, 5000);

    Office.context.document.settings.refreshAsync((result) => {
      clearTimeout(timeout);
      console.log("[sheet-id-map] loadFromSettings refreshAsync callback", result.status);
      cachedMap = Office.context.document.settings.get(SETTINGS_KEY_MAP) || {};
      cachedCounter =
        Office.context.document.settings.get(SETTINGS_KEY_COUNTER) || 0;
      resolve();
    });
  });
}

async function saveToSettings(): Promise<void> {
  if (!isDirty) return;

  console.log("[sheet-id-map] Saving to settings...");
  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      console.warn("[sheet-id-map] saveToSettings timed out after 5s");
      isDirty = false;
      resolve();
    }, 5000);

    Office.context.document.settings.set(SETTINGS_KEY_MAP, cachedMap);
    Office.context.document.settings.set(SETTINGS_KEY_COUNTER, cachedCounter);
    Office.context.document.settings.saveAsync((result) => {
      clearTimeout(timeout);
      console.log("[sheet-id-map] saveToSettings saveAsync callback", result.status);
      isDirty = false;
      resolve();
    });
  });
}

export async function getStableSheetId(guid: string): Promise<number> {
  if (cachedMap === null) {
    await loadFromSettings();
  }

  if (cachedMap![guid]) {
    return cachedMap![guid];
  }

  cachedCounter = (cachedCounter || 0) + 1;
  cachedMap![guid] = cachedCounter;
  isDirty = true;
  await saveToSettings();

  return cachedCounter;
}

export function getExistingSheetId(guid: string): number | null {
  if (cachedMap === null) return null;
  return cachedMap[guid] || null;
}

export async function getAllSheetIds(): Promise<SheetIdMap> {
  if (cachedMap === null) {
    await loadFromSettings();
  }
  return { ...cachedMap! };
}

export async function clearSheetIds(): Promise<void> {
  cachedMap = {};
  cachedCounter = 0;
  isDirty = true;
  await saveToSettings();
}

export async function preloadSheetIds(
  worksheets: Excel.Worksheet[],
): Promise<Map<string, number>> {
  if (cachedMap === null) {
    await loadFromSettings();
  }

  const result = new Map<string, number>();
  let needsSave = false;

  for (const sheet of worksheets) {
    const guid = sheet.id;

    if (cachedMap![guid]) {
      result.set(guid, cachedMap![guid]);
    } else {
      cachedCounter = (cachedCounter || 0) + 1;
      cachedMap![guid] = cachedCounter;
      result.set(guid, cachedCounter);
      needsSave = true;
    }
  }

  if (needsSave) {
    isDirty = true;
    await saveToSettings();
  }

  return result;
}
