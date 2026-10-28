import { exportSpreadsheet, loadSpreadsheet, type CellEdit, type SpreadsheetDocument } from './spreadsheet'

/** Refuse to hand out a file unless rereading it reproduces the complete intended matrix. */
export function verifiedExport(doc: SpreadsheetDocument, edits: CellEdit[]): Uint8Array {
  const bytes = exportSpreadsheet(doc, edits)
  const reread = loadSpreadsheet(doc.name, bytes)
  const expectedEdits = new Map<string, Map<number, Map<number, string>>>()
  for (const edit of edits) {
    let sheet = expectedEdits.get(edit.sheetId)
    if (!sheet) { sheet = new Map(); expectedEdits.set(edit.sheetId, sheet) }
    let row = sheet.get(edit.rowIndex)
    if (!row) { row = new Map(); sheet.set(edit.rowIndex, row) }
    row.set(edit.columnIndex, edit.value)
  }

  function mismatch(detail: string): never {
    throw new Error(`내보내기 검증에 실패했습니다. 파일 다운로드를 중단했습니다: ${detail}`)
  }

  if (reread.kind !== doc.kind || reread.sheets.length !== doc.sheets.length) mismatch('파일 형식 또는 워크시트 수가 원본과 다릅니다.')
  for (let sheetIndex = 0; sheetIndex < doc.sheets.length; sheetIndex++) {
    const original = doc.sheets[sheetIndex]
    const actual = reread.sheets[sheetIndex]
    if (original.id !== actual.id || original.name !== actual.name) mismatch('워크시트 ID, 이름 또는 순서가 바뀌었습니다.')
    if (original.rows.length !== actual.rows.length) mismatch(`${original.name}: 행 수가 바뀌었습니다.`)
    if (original.maxColumns !== actual.maxColumns) mismatch(`${original.name}: 원본 열 범위가 바뀌었습니다.`)
    const sheetEdits = expectedEdits.get(original.id)
    for (let rowIndex = 0; rowIndex < original.rows.length; rowIndex++) {
      const before = original.rows[rowIndex]
      const after = actual.rows[rowIndex]
      const rowEdits = sheetEdits?.get(rowIndex)
      let expectedLength = before.length
      if (rowEdits) for (const columnIndex of rowEdits.keys()) expectedLength = Math.max(expectedLength, columnIndex + 1)
      // Missing trailing fields can legitimately be materialized by an edit. No other growth is permitted.
      if (after.length < before.length || after.length > expectedLength) mismatch(`${original.name}, ${rowIndex + 1}행: 허용하지 않은 셀 구조 변경입니다.`)
      for (let columnIndex = 0; columnIndex < expectedLength; columnIndex++) {
        const expected = rowEdits?.has(columnIndex) ? rowEdits.get(columnIndex)! : before[columnIndex] ?? ''
        if ((after[columnIndex] ?? '') !== expected) mismatch(`${original.name}, ${rowIndex + 1}행 ${columnIndex + 1}열: 값이 일치하지 않습니다.`)
      }
    }
  }
  return bytes
}
