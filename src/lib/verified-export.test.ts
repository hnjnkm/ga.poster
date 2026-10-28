// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { strToU8, unzipSync, zipSync } from 'fflate'
import * as spreadsheet from './spreadsheet'
import { verifiedExport } from './verified-export'

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
function workbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Archive" sheetId="1" r:id="rId1"/><sheet name="Keep" sheetId="2" r:id="rId2"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'),
    'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c><c r="D1" t="inlineStr"><is><t>中文</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>0001</t></is></c></row></sheetData></worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>unchanged</t></is></c></row></sheetData></worksheet>`),
  })
}

afterEach(() => vi.restoreAllMocks())

describe('verifiedExport 전체 데이터 재검증', () => {
  it('verifies 1000 unsorted IDs without sorting or changing untouched fields', () => {
    const lines = ['ID,한국어,English,简体中文,Untouched']
    for (let index = 0; index < 1000; index++) lines.push(`${String((index * 37) % 1000).padStart(4, '0')},한글 ${index},English ${index},中文 ${index},keep ${index}`)
    const original = strToU8(lines.join('\r\n') + '\r\n')
    const doc = spreadsheet.loadSpreadsheet('archive.csv', original)
    const edits = Array.from({ length: 1000 }, (_, index) => ({ sheetId: 'csv-0', rowIndex: index + 1, columnIndex: 1, value: `수정 ${index}` }))
    const output = verifiedExport(doc, edits)
    const reread = spreadsheet.loadSpreadsheet('archive.csv', output)
    expect(reread.sheets[0].rows.map(row => row[0])).toEqual(doc.sheets[0].rows.map(row => row[0]))
    expect(reread.sheets[0].rows[1000]).toEqual(['0963', '수정 999', 'English 999', '中文 999', 'keep 999'])
    expect(Array.from(verifiedExport(doc, []))).toEqual(Array.from(original))
  })

  it('permits filling absent CSV trailing cells and verifies all other values', () => {
    const doc = spreadsheet.loadSpreadsheet('archive.csv', strToU8('id,ko,en,zh\n01,original\n'))
    const output = verifiedExport(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 3, value: '中文' }])
    expect(spreadsheet.loadSpreadsheet('archive.csv', output).sheets[0].rows[1]).toEqual(['01', 'original', '', '中文'])
  })

  it('verifies every XLSX sheet while allowing a missing text cell to be filled', () => {
    const doc = spreadsheet.loadSpreadsheet('archive.xlsx', workbook())
    const output = verifiedExport(doc, [{ sheetId: 'xlsx-0', rowIndex: 1, columnIndex: 3, value: '简体中文 & <原文>' }])
    const reread = spreadsheet.loadSpreadsheet('archive.xlsx', output)
    expect(reread.sheets[0].rows[1]).toEqual(['0001', '', '', '简体中文 & <原文>'])
    expect(reread.sheets[1].rows).toEqual([['unchanged']])
  })

  it('rejects duplicate edits before returning a file', () => {
    const doc = spreadsheet.loadSpreadsheet('archive.csv', strToU8('id,ko\n01,original'))
    const edit = { sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: 'new' }
    expect(() => verifiedExport(doc, [edit, edit])).toThrow('중복')
  })

  it.each([
    ['changed untouched value', 'id,ko,en\n01,new,CORRUPTED'],
    ['changed ID', 'id,ko,en\n02,new,keep'],
    ['missing requested edit', 'id,ko,en\n01,original,keep'],
    ['added row', 'id,ko,en\n01,new,keep\n02,extra,extra'],
    ['added blank column', 'id,ko,en\n01,new,keep,'],
    ['removed original trailing empty field', 'id,ko,en\n01,new'],
  ])('blocks an exporter regression: %s', (_label, corrupt) => {
    const doc = spreadsheet.loadSpreadsheet('archive.csv', strToU8('id,ko,en\n01,original,keep'))
    vi.spyOn(spreadsheet, 'exportSpreadsheet').mockReturnValueOnce(strToU8(corrupt))
    expect(() => verifiedExport(doc, [{ sheetId: 'csv-0', rowIndex: 1, columnIndex: 1, value: 'new' }])).toThrow('내보내기 검증')
  })

  it('blocks corruption on an entirely unedited XLSX worksheet', () => {
    const original = workbook()
    const doc = spreadsheet.loadSpreadsheet('archive.xlsx', original)
    const entries = unzipSync(original)
    entries['xl/worksheets/sheet2.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>CORRUPTED</t></is></c></row></sheetData></worksheet>`)
    vi.spyOn(spreadsheet, 'exportSpreadsheet').mockReturnValueOnce(zipSync(entries))
    expect(() => verifiedExport(doc, [])).toThrow('값이 일치하지 않습니다')
  })
})
