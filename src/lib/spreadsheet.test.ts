// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { assertEditableCell, exportSpreadsheet, findCellIssue, loadSpreadsheet, type CellEdit } from './spreadsheet'

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${ns}"><dimension ref="A1:F4"/><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>순번</t></is></c><c r="B1" t="inlineStr"><is><t>한국어</t></is></c><c r="C1" t="inlineStr"><is><t>English</t></is></c><c r="D1" t="inlineStr"><is><t>简体中文</t></is></c><c r="E1" t="inlineStr"><is><t>Formula</t></is></c><c r="F1" t="inlineStr"><is><t>Other</t></is></c></row>
<row r="2" customFormat="1"><c r="A2" t="s"><v>0</v></c><c r="B2" s="4" t="inlineStr"><is><t xml:space="preserve"> 원문 &amp; 문장\n둘째 줄 </t></is></c><c r="C2" s="8"><v>0012</v></c><c r="E2"><f>1+1</f><v>2</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>0002</t></is></c><c r="B3" t="inlineStr"><is><t>병합</t></is></c></row>
<row r="4"/>
</sheetData><mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells><pageMargins left="0.7" right="0.7"/></worksheet>`

function xlsx(xml = worksheet): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    'xl/workbook.xml': strToU8(`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Webflow 가져오기" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
    'xl/sharedStrings.xml': strToU8(`<sst xmlns="${ns}"><si><t>0001</t></si></sst>`),
    'xl/worksheets/sheet1.xml': strToU8(xml),
    'xl/styles.xml': strToU8('<styles original="keep this byte-for-byte"/>'),
    'custom/opaque.bin': new Uint8Array([0, 255, 128, 7, 3]),
  })
}

describe('CSV 원본 보존', () => {
  it('keeps leading zeros, multiline quotes, BOM, all original bytes and row order', () => {
    const original = strToU8('\uFEFF순번,한국어,English,简体中文,외부 ID\r\n"0001","한글\r\n""인용""",old,"中文",=original\r\n0002,,,未修改,9\r\n')
    const doc = loadSpreadsheet('archive.csv', original)
    expect(doc.sheets[0].rows[1]).toEqual(['0001', '한글\r\n"인용"', 'old', '中文', '=original'])
    expect(Array.from(exportSpreadsheet(doc, []))).toEqual(Array.from(original))
    expect(Array.from(exportSpreadsheet(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: '한글\r\n"인용"' }]))).toEqual(Array.from(original))
    const exported = exportSpreadsheet(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 2, value: 'New, "quoted"\nline' }])
    expect(strFromU8(exported)).toBe(strFromU8(original).replace(',old,', ',"New, ""quoted""\nline",'))
    expect(loadSpreadsheet('roundtrip.csv', exported).sheets[0].rows[1][2]).toBe('New, "quoted"\nline')
  })

  it('preserves existing quote style and mixed newline sequences while editing', () => {
    const doc = loadSpreadsheet('a.csv', strToU8('id;ko;en\r\n01;"원문";keep\n02;둘;keep\r'))
    expect(strFromU8(exportSpreadsheet(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: '수정' }])))
      .toBe('id;ko;en\r\n01;"수정";keep\n02;둘;keep\r')
  })

  it('adds only the necessary delimiters for a genuinely absent trailing field', () => {
    const doc = loadSpreadsheet('a.csv', strToU8('id,ko,en,zh\n01,original\n'))
    const edits: CellEdit[] = [
      { sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: 'changed' },
      { sheetId: 'csv-0', rowIndex: 1, columnIndex: 3, value: '简体中文' },
    ]
    expect(strFromU8(exportSpreadsheet(doc, edits))).toBe('id,ko,en,zh\n01,changed,,简体中文\n')
  })

  it('handles a final empty field without introducing a phantom row', () => {
    expect(loadSpreadsheet('a.csv', strToU8('id,ko\n001,')).sheets[0].rows).toEqual([['id', 'ko'], ['001', '']])
    expect(loadSpreadsheet('a.csv', strToU8('id,ko\n001,\n')).sheets[0].rows).toEqual([['id', 'ko'], ['001', '']])
  })

  it('detects tabs and preserves them on export', () => {
    const doc = loadSpreadsheet('a.csv', strToU8('id\tko\ten\n001\t원문\tEnglish'))
    expect(strFromU8(exportSpreadsheet(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: '수정\t글' }])))
      .toBe('id\tko\ten\n001\t"수정\t글"\tEnglish')
  })

  it('rejects invalid UTF-8, malformed quotes, and ambiguous duplicate edits', () => {
    expect(() => loadSpreadsheet('a.csv', new Uint8Array([0xff, 0xfe, 0x41, 0]))).toThrow('UTF-8')
    expect(() => loadSpreadsheet('a.csv', strToU8('id,ko\n01,"unclosed'))).toThrow('큰따옴표')
    expect(() => loadSpreadsheet('a.csv', strToU8('id,ko\n01,"closed"extra'))).toThrow('큰따옴표')
    const doc = loadSpreadsheet('a.csv', strToU8('id,ko\n01,원문'))
    const edit = { sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: 'changed' }
    expect(() => exportSpreadsheet(doc, [edit, edit])).toThrow('중복')
    expect(() => exportSpreadsheet(doc, [{ ...edit, columnIndex: 2 }])).toThrow('열 범위')
    expect(() => exportSpreadsheet(doc, [{ ...edit, rowIndex: 2 }])).toThrow('행 범위')
    expect(() => exportSpreadsheet(doc, [{ ...edit, value: '\uD800' }])).toThrow('유니코드')
  })
})

describe('XLSX 선택 셀만 변경', () => {
  it('reads shared/inline strings, raw numeric identifiers and explicit row positions', () => {
    const original = xlsx()
    const doc = loadSpreadsheet('archive.xlsx', original)
    expect(doc.sheets[0].name).toBe('Webflow 가져오기')
    expect(doc.sheets[0].rows[1]).toEqual(['0001', ' 원문 & 문장\n둘째 줄 ', '0012', '', '2'])
    expect(doc.sheets[0].rows[3]).toEqual([])
    expect(exportSpreadsheet(doc, [])).toEqual(original)
  })

  it('keeps every unrelated ZIP entry payload and XML outside the selected cell byte-identical', () => {
    const original = xlsx()
    const doc = loadSpreadsheet('archive.xlsx', original)
    const value = ' 새 <설명> & "인용"\nEnglish 中文 '
    const output = exportSpreadsheet(doc, [{ sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 1, value }])
    const before = unzipSync(original)
    const after = unzipSync(output)
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort())
    for (const path of Object.keys(before).filter(path => path !== 'xl/worksheets/sheet1.xml')) expect(after[path]).toEqual(before[path])
    const changedXml = strFromU8(after['xl/worksheets/sheet1.xml'])
    const selected = /<c\b[^>]*\br="B2"[^>]*>[\s\S]*?<\/c>/
    expect(changedXml.replace(selected, '__SELECTED__')).toBe(worksheet.replace(selected, '__SELECTED__'))
    expect(changedXml).toContain('s="4"')
    expect(changedXml).toContain('&lt;설명&gt; &amp;')
    const reread = loadSpreadsheet('archive.xlsx', output)
    expect(reread.sheets[0].rows[1][1]).toBe(value)
    expect(reread.sheets[0].rows[1][0]).toBe('0001')
    expect(reread.sheets[0].rows[1][2]).toBe('0012')
  })

  it('inserts a missing text cell in coordinate order without altering an adjacent edited cell', () => {
    const doc = loadSpreadsheet('archive.xlsx', xlsx())
    const output = exportSpreadsheet(doc, [
      { sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 3, value: '中文' },
      { sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 2, value: '00123' },
      { sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 5, value: 'last' },
    ])
    expect(loadSpreadsheet('archive.xlsx', output).sheets[0].rows[1]).toEqual(['0001', ' 원문 & 문장\n둘째 줄 ', '00123', '中文', '2', 'last'])
    const xml = strFromU8(unzipSync(output)['xl/worksheets/sheet1.xml'])
    expect(xml.indexOf('r="C2"')).toBeLessThan(xml.indexOf('r="D2"'))
    expect(xml.indexOf('r="D2"')).toBeLessThan(xml.indexOf('r="E2"'))
  })

  it('allows multiple cells in an explicit empty row', () => {
    const doc = loadSpreadsheet('archive.xlsx', xlsx())
    const output = exportSpreadsheet(doc, [
      { sheetId: 'xlsx-0', rowIndex: 3, columnIndex: 2, value: 'third' },
      { sheetId: 'xlsx-0', rowIndex: 3, columnIndex: 0, value: '0003' },
    ])
    expect(loadSpreadsheet('archive.xlsx', output).sheets[0].rows[3]).toEqual(['0003', '', 'third'])
  })

  it('handles namespace prefixes, inferred cell addresses and insertion before an edited neighbor', () => {
    const xml = `<m:worksheet xmlns:m="${ns}"><m:sheetData><m:row r="1"><m:c r="A1" t="inlineStr"><m:is><m:t>ID</m:t></m:is></m:c><m:c r="D1" t="inlineStr"><m:is><m:t>zh</m:t></m:is></m:c></m:row><m:row r="2"><m:c t="inlineStr"><m:is><m:t>0001</m:t></m:is></m:c><m:c r="C2" t="inlineStr"><m:is><m:t>old</m:t></m:is></m:c></m:row></m:sheetData></m:worksheet>`
    const doc = loadSpreadsheet('prefix.xlsx', xlsx(xml))
    const output = exportSpreadsheet(doc, [
      { sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 1, value: '한국어' },
      { sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 2, value: 'English' },
    ])
    expect(loadSpreadsheet('prefix.xlsx', output).sheets[0].rows[1]).toEqual(['0001', '한국어', 'English'])
    const patched = strFromU8(unzipSync(output)['xl/worksheets/sheet1.xml'])
    expect(patched).toContain('<m:c t="inlineStr"><m:is><m:t>0001</m:t></m:is></m:c>')
  })

  it('preserves literal OOXML escape patterns, actual CR and control characters', () => {
    const doc = loadSpreadsheet('archive.xlsx', xlsx())
    const value = 'literal _x0041_ / literal _x005F_ / CR\r / control\u0001 / emoji 🎨'
    const output = exportSpreadsheet(doc, [{ sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 1, value }])
    expect(loadSpreadsheet('archive.xlsx', output).sheets[0].rows[1][1]).toBe(value)
  })

  it('rejects formulas, all cells in merged ranges and invalid coordinates', () => {
    const doc = loadSpreadsheet('archive.xlsx', xlsx())
    expect(findCellIssue(doc, 'xlsx-0', 1, 4)).toContain('수식')
    expect(findCellIssue(doc, 'xlsx-0', 2, 1)).toContain('병합')
    expect(findCellIssue(doc, 'xlsx-0', 2, 2)).toContain('병합')
    expect(() => assertEditableCell(doc, 'xlsx-0', 1, 4)).toThrow('수식')
    expect(() => exportSpreadsheet(doc, [{ sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 4, value: 'changed' }])).toThrow('수식')
    expect(() => exportSpreadsheet(doc, [{ sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 1, value: 'a'.repeat(32_768) }])).toThrow('32,767')
    expect(findCellIssue(doc, 'missing', 1, 1)).toContain('워크시트')
    expect(findCellIssue(doc, 'xlsx-0', -1, 1)).toContain('행 범위')
    expect(findCellIssue(doc, 'xlsx-0', 1, NaN)).toContain('열 범위')
  })

  it('retains sparse row indices and refuses to invent absent rows', () => {
    const xml = `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>0002</t></is></c></row></sheetData></worksheet>`
    const doc = loadSpreadsheet('sparse.xlsx', xlsx(xml))
    expect(doc.sheets[0].rows).toEqual([['ID'], [], ['0002']])
    expect(findCellIssue(doc, 'xlsx-0', 1, 0)).toContain('존재하지 않는')
  })

  it('rejects oversized worksheet coordinates, duplicate cells and XML doctypes', () => {
    expect(() => loadSpreadsheet('huge.xlsx', xlsx(`<worksheet xmlns="${ns}"><sheetData><row r="20001"/></sheetData></worksheet>`))).toThrow('20,000')
    expect(() => loadSpreadsheet('duplicate.xlsx', xlsx(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"/><c r="A1"/></row></sheetData></worksheet>`))).toThrow('중복')
    expect(() => loadSpreadsheet('unsafe.xlsx', xlsx(`<!DOCTYPE worksheet><worksheet xmlns="${ns}"><sheetData/></worksheet>`))).toThrow('DOCTYPE')
  })
})
