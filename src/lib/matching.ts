export interface ColumnMapping {
  /** Zero-based index of the spreadsheet header row. */
  headerRow: number;
  idColumn: number;
  koColumn: number;
  enColumn: number;
  zhColumn: number;
}

export interface ImageIdentifier {
  id: string;
  name: string;
  originalStem: string;
}

export interface IdentifiedImage extends ImageIdentifier {
  imageIndex: number;
  imageName: string;
}

export interface IdentifiedRow {
  id: string;
  rowIndex: number;
  originalId: string;
}

export interface MatchedRecord extends IdentifiedRow {
  imageIndex: number;
  imageName: string;
}

export interface MappingAnalysis {
  /** Unmatched records are reported separately and do not invalidate a mapping. */
  valid: boolean;
  errors: string[];
  mappingErrors: string[];
  rowIds: IdentifiedRow[];
  images: IdentifiedImage[];
  /** Only this display list is sorted. Spreadsheet rows are never reordered. */
  sortedImages: IdentifiedImage[];
  matches: MatchedRecord[];
  duplicateRowIds: { id: string; rowIndices: number[]; values: string[] }[];
  duplicateImageIds: { id: string; imageIndices: number[]; imageNames: string[] }[];
  invalidRows: { rowIndex: number; value: string; reason: string }[];
  invalidImages: { imageIndex: number; imageName: string; reason: string }[];
  unmatchedRows: IdentifiedRow[];
  unmatchedImages: IdentifiedImage[];
}

/**
 * IDs are decimal strings, never floating-point numbers. Do not trim or extract
 * digits from text: an accidental space or a different naming scheme needs review.
 */
export function normalizeIdentifier(value: string): string | null {
  if (value.length === 0 || /[^0-9]/.test(value)) return null;
  return value.replace(/^0+(?=[0-9])/, "");
}

/** Accept a single filename, with a numeric stem and a supported image suffix. */
export function parseImageName(name: string): ImageIdentifier | null {
  const match = /^([0-9]+)\.(png|jpe?g|webp|avif|gif)$/i.exec(name);
  if (!match || match[0] !== name) return null;
  return {
    id: normalizeIdentifier(match[1])!,
    name,
    originalStem: match[1],
  };
}

/** Compare canonical decimal IDs without losing precision for very large IDs. */
export function compareCanonicalIds(a: string, b: string): number {
  return a.length - b.length || (a < b ? -1 : a > b ? 1 : 0);
}

const columnKeys = ["idColumn", "koColumn", "enColumn", "zhColumn"] as const;

function validateMapping(rows: string[][], mapping: ColumnMapping): string[] {
  const errors: string[] = [];
  if (
    !Number.isInteger(mapping.headerRow) ||
    mapping.headerRow < 0 ||
    mapping.headerRow >= rows.length
  ) {
    return ["머리글 행을 올바르게 선택해 주세요."];
  }

  const width = rows[mapping.headerRow].length;
  for (const key of columnKeys) {
    const index = mapping[key];
    if (!Number.isInteger(index) || index < 0 || index >= width) {
      errors.push(`${key}: 유효한 열을 선택해 주세요.`);
    }
  }
  if (new Set(columnKeys.map((key) => mapping[key])).size !== columnKeys.length) {
    errors.push("순번·한국어·영어·중국어 열은 각각 다른 열이어야 합니다.");
  }
  return errors;
}

function groupById<T extends { id: string }>(items: T[]): Map<string, T[]> {
  const groups = new Map<string, T[]>();
  for (const item of items) {
    const group = groups.get(item.id);
    if (group) group.push(item);
    else groups.set(item.id, [item]);
  }
  return groups;
}

/**
 * Match only explicitly chosen spreadsheet IDs to numeric image filenames.
 * Indices always refer to the original input; this function never mutates either
 * input or falls back to row position, description text, or fuzzy matching.
 */
export function analyzeMapping(
  rows: string[][],
  mapping: ColumnMapping,
  imageNames: string[],
): MappingAnalysis {
  const mappingErrors = validateMapping(rows, mapping);
  const errors = [...mappingErrors];
  const rowIds: IdentifiedRow[] = [];
  const images: IdentifiedImage[] = [];
  const invalidRows: MappingAnalysis["invalidRows"] = [];
  const invalidImages: MappingAnalysis["invalidImages"] = [];

  if (mappingErrors.length === 0) {
    for (let rowIndex = mapping.headerRow + 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      // Empty rows remain in the original spreadsheet, but are not records.
      if (row.every((cell) => cell === "")) continue;
      const originalId = row[mapping.idColumn] ?? "";
      const id = normalizeIdentifier(originalId);
      if (id === null) {
        invalidRows.push({
          rowIndex,
          value: originalId,
          reason: "순번은 공백 없이 숫자로만 입력되어야 합니다.",
        });
      } else {
        rowIds.push({ id, rowIndex, originalId });
      }
    }
  }

  imageNames.forEach((imageName, imageIndex) => {
    const parsed = parseImageName(imageName);
    if (!parsed) {
      invalidImages.push({
        imageIndex,
        imageName,
        reason: "숫자 이름의 PNG, JPG, JPEG, WEBP, AVIF, GIF 이미지만 지원합니다.",
      });
    } else {
      images.push({ ...parsed, imageIndex, imageName });
    }
  });

  const rowGroups = groupById(rowIds);
  const imageGroups = groupById(images);
  const duplicateRowIds = [...rowGroups.entries()]
    .filter(([, group]) => group.length > 1)
    .map(([id, group]) => ({
      id,
      rowIndices: group.map((row) => row.rowIndex),
      values: group.map((row) => row.originalId),
    }));
  const duplicateImageIds = [...imageGroups.entries()]
    .filter(([, group]) => group.length > 1)
    .map(([id, group]) => ({
      id,
      imageIndices: group.map((image) => image.imageIndex),
      imageNames: group.map((image) => image.imageName),
    }));

  const matches: MatchedRecord[] = [];
  const unmatchedRows: IdentifiedRow[] = [];
  for (const row of rowIds) {
    const matchingImages = imageGroups.get(row.id);
    if (!matchingImages) unmatchedRows.push(row);
    else if (rowGroups.get(row.id)!.length === 1 && matchingImages.length === 1) {
      matches.push({
        ...row,
        imageIndex: matchingImages[0].imageIndex,
        imageName: matchingImages[0].imageName,
      });
    }
  }
  const unmatchedImages = images.filter((image) => !rowGroups.has(image.id));

  if (invalidRows.length) errors.push(`숫자 순번이 없거나 올바르지 않은 행: ${invalidRows.length}개`);
  if (invalidImages.length) errors.push(`파일명을 확인해야 하는 이미지: ${invalidImages.length}개`);
  if (duplicateRowIds.length) errors.push(`중복된 표 순번: ${duplicateRowIds.map((item) => item.id).join(", ")}`);
  if (duplicateImageIds.length) errors.push(`중복된 이미지 순번: ${duplicateImageIds.map((item) => item.id).join(", ")}`);

  return {
    valid: errors.length === 0,
    errors,
    mappingErrors,
    rowIds,
    images,
    sortedImages: [...images].sort(
      (a, b) => compareCanonicalIds(a.id, b.id) || a.imageIndex - b.imageIndex,
    ),
    matches,
    duplicateRowIds,
    duplicateImageIds,
    invalidRows,
    invalidImages,
    unmatchedRows,
    unmatchedImages,
  };
}

function normalizedHeader(header: string): string {
  return header.trim().toLowerCase().replace(/[\s_()\-]+/g, "");
}

const headerAliases: Record<(typeof columnKeys)[number], string[]> = {
  idColumn: ["id", "순번", "번호", "이미지번호", "imageid", "imagenumber", "no", "number"],
  koColumn: [
    "ko", "kor", "korean", "한국어", "한글", "한국어설명", "한글설명",
    "descriptionko", "kodescription", "descriptionkorean", "설명한국어", "설명한글",
  ],
  enColumn: [
    "en", "eng", "english", "영어", "영어설명", "descriptionen", "endescription", "descriptionenglish",
  ],
  zhColumn: [
    "zh", "zhcn", "zhhans", "chinesesimplified", "simplifiedchinese", "중국어간체",
    "중국어", "간체", "简体中文", "简体", "中文简体", "中文", "descriptionzh",
    "zhdescription", "descriptionzhcn", "descriptionzhhans", "중국어간체설명",
  ],
};

/** Suggestions only: a user must confirm the selected columns before editing. */
export function suggestMapping(headers: string[]): Partial<Omit<ColumnMapping, "headerRow">> {
  const result: Partial<Omit<ColumnMapping, "headerRow">> = {};
  for (const key of columnKeys) {
    const candidates = headers.flatMap((header, index) =>
      headerAliases[key].includes(normalizedHeader(header)) ? [index] : [],
    );
    if (candidates.length === 1) result[key] = candidates[0];
  }
  return result;
}
