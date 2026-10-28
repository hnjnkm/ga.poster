import { describe, expect, it } from "vitest";
import {
  analyzeMapping,
  compareCanonicalIds,
  normalizeIdentifier,
  parseImageName,
  suggestMapping,
  type ColumnMapping,
} from "./matching";

const mapping: ColumnMapping = {
  headerRow: 0,
  idColumn: 0,
  koColumn: 1,
  enColumn: 2,
  zhColumn: 3,
};
const headers = ["순번", "한글", "English", "简体中文"];

describe("explicit identifiers", () => {
  it("normalizes only ASCII decimal digits without losing original filenames", () => {
    expect(normalizeIdentifier("0001002")).toBe("1002");
    expect(normalizeIdentifier("000")).toBe("0");
    expect(parseImageName("01.PNG")).toEqual({ id: "1", name: "01.PNG", originalStem: "01" });
    for (const value of [" 01", "01 ", "01\n", "01\r", "01\t", "1.0", "1e3", "+1", "image01", "１", ""]) {
      expect(normalizeIdentifier(value)).toBeNull();
    }
  });

  it("rejects unsupported extensions, paths, and names with extra text", () => {
    for (const name of ["01.svg", "poster01.png", "01 copy.png", "01.png.exe", "a/01.png", " 01.png", "01.png ", "01.png\n"]) {
      expect(parseImageName(name)).toBeNull();
    }
    for (const name of ["0.gif", "002.webp", "1002.avif", "3.jpeg", "4.jpg"]) {
      expect(parseImageName(name)).not.toBeNull();
    }
  });

  it("compares arbitrarily large numeric IDs without converting to Number", () => {
    const values = ["100000000000000000000", "9007199254740993", "9007199254740992", "2"];
    expect(values.sort(compareCanonicalIds)).toEqual([
      "2", "9007199254740992", "9007199254740993", "100000000000000000000",
    ]);
  });
});

describe("mapping integrity", () => {
  it("preserves spreadsheet order and content while sorting only image display order", () => {
    const rows = [headers, ["1002", "원문\n그대로", "Original", "原文"], ["02", "둘", "Two", "二"], ["01", "하나", "One", "一"]];
    const original = rows.map((row) => [...row]);
    const names = ["01.png", "1002.png", "02.png"];
    const result = analyzeMapping(rows, mapping, names);
    expect(result.valid).toBe(true);
    expect(result.matches.map((match) => [match.id, match.rowIndex, match.imageIndex])).toEqual([
      ["1002", 1, 1], ["2", 2, 2], ["1", 3, 0],
    ]);
    expect(result.sortedImages.map((image) => image.id)).toEqual(["1", "2", "1002"]);
    expect(result.matches[1].originalId).toBe("02");
    expect(rows).toEqual(original);
    expect(names).toEqual(["01.png", "1002.png", "02.png"]);
  });

  it("blocks canonical duplicate spreadsheet IDs and does not choose one", () => {
    const result = analyzeMapping([headers, ["01", "", "", ""], ["1", "", "", ""]], mapping, ["1.png"]);
    expect(result.valid).toBe(false);
    expect(result.duplicateRowIds).toEqual([{ id: "1", rowIndices: [1, 2], values: ["01", "1"] }]);
    expect(result.matches).toEqual([]);
  });

  it("blocks canonical duplicate image IDs across formats", () => {
    const result = analyzeMapping([headers, ["01", "", "", ""]], mapping, ["1.png", "001.jpg"]);
    expect(result.valid).toBe(false);
    expect(result.duplicateImageIds).toEqual([{ id: "1", imageIndices: [0, 1], imageNames: ["1.png", "001.jpg"] }]);
    expect(result.matches).toEqual([]);
  });

  it("reports missing and malformed IDs instead of falling back to row order", () => {
    const result = analyzeMapping([
      headers,
      ["", "image one", "1", "一"],
      [" 2", "둘", "Two", "二"],
      ["poster3", "셋", "Three", "三"],
      ["", "", "", ""],
    ], mapping, ["1.png", "2.png", "3.png"]);
    expect(result.invalidRows.map((row) => row.rowIndex)).toEqual([1, 2, 3]);
    expect(result.matches).toEqual([]);
    expect(result.unmatchedImages).toHaveLength(3);
    expect(result.valid).toBe(false);
  });

  it("reports unmatched records without fuzzy text or positional matches", () => {
    const result = analyzeMapping([headers, ["01", "2.png", "Two", "二"], ["1002", "", "", ""]], mapping, ["2.png", "1002.png", "poster1.png"]);
    expect(result.matches.map((match) => match.id)).toEqual(["1002"]);
    expect(result.unmatchedRows.map((row) => row.id)).toEqual(["1"]);
    expect(result.unmatchedImages.map((image) => image.id)).toEqual(["2"]);
    expect(result.invalidImages.map((image) => image.imageName)).toEqual(["poster1.png"]);
  });

  it("allows a valid subset of uploaded images without changing other rows", () => {
    const result = analyzeMapping([headers, ["1", "기존", "Existing", "现有"], ["2", "", "", ""]], mapping, ["2.png"]);
    expect(result.valid).toBe(true);
    expect(result.unmatchedRows.map((row) => row.id)).toEqual(["1"]);
  });

  it("requires four distinct valid column indices and an explicit header row", () => {
    expect(analyzeMapping([headers], { ...mapping, enColumn: 1 }, []).mappingErrors).not.toHaveLength(0);
    expect(analyzeMapping([headers], { ...mapping, idColumn: -1 }, []).mappingErrors).not.toHaveLength(0);
    expect(analyzeMapping([headers], { ...mapping, zhColumn: 4 }, []).mappingErrors).not.toHaveLength(0);
    expect(analyzeMapping([headers], { ...mapping, headerRow: 1 }, []).mappingErrors).not.toHaveLength(0);
    expect(analyzeMapping([], mapping, []).mappingErrors).not.toHaveLength(0);
  });

  it("uses original row indices when the header follows a preamble", () => {
    const result = analyzeMapping([["설명"], headers, ["0001", "", "", ""]], { ...mapping, headerRow: 1 }, ["1.png"]);
    expect(result.matches[0]).toMatchObject({ rowIndex: 2, id: "1", originalId: "0001" });
  });
});

describe("column suggestions", () => {
  it("suggests only unique exact normalized aliases", () => {
    expect(suggestMapping(["Image ID", "Description (Korean)", "EN", "zh-CN"])).toEqual({
      idColumn: 0, koColumn: 1, enColumn: 2, zhColumn: 3,
    });
  });

  it("does not pick an arbitrary first column, partial label, or ambiguous language", () => {
    expect(suggestMapping(["Name", "English title", "Other"])).toEqual({});
    expect(suggestMapping(["ID", "KO", "한국어", "EN", "ZH"])).toEqual({ idColumn: 0, enColumn: 3, zhColumn: 4 });
  });
});
