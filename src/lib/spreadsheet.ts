import { unzipSync, zipSync } from 'fflate'

/** Coordinates always refer to the original worksheet, including its header. */
export interface CellEdit {
  sheetId: string
  rowIndex: number
  columnIndex: number
  value: string
}

export interface SpreadsheetSheet {
  id: string
  name: string
  rows: string[][]
  maxColumns: number
}

export interface SpreadsheetDocument {
  kind: 'csv' | 'xlsx'
  name: string
  sheets: SpreadsheetSheet[]
}

const MAX_ROWS = 20_000
const MAX_COLUMNS = 256
const MAX_FILE_BYTES = 25 * 1024 * 1024
const MAX_UNZIPPED_BYTES = 100 * 1024 * 1024
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const encoder = new TextEncoder()

interface Span { start: number; end: number }
interface CsvToken extends Span { value: string; quoted: boolean }
interface CsvRow { tokens: CsvToken[]; end: number }
interface CsvSource {
  kind: 'csv'
  bytes: Uint8Array
  text: string
  delimiter: string
  rows: CsvRow[]
  sheet: SpreadsheetSheet
}
interface XmlSpan extends Span { closeStart: number; selfClosing: boolean }
interface RowSpan extends XmlSpan { cells: XmlSpan[] }
interface Merge { top: number; left: number; bottom: number; right: number }
interface XlsxSheetSource {
  sheet: SpreadsheetSheet
  path: string
  xml: string
  document: XMLDocument
  rowElements: Map<number, Element>
  rowSpans: Map<number, RowSpan>
  cells: Map<string, { element: Element; span: XmlSpan }>
  formulas: Set<string>
  merges: Merge[]
}
interface XlsxSource {
  kind: 'xlsx'
  bytes: Uint8Array
  entries: Record<string, Uint8Array>
  sheets: Map<string, XlsxSheetSource>
}
type Source = CsvSource | XlsxSource
const sources = new WeakMap<SpreadsheetDocument, Source>()

function fail(message: string): never { throw new Error(message) }

function decodeUtf8(bytes: Uint8Array, label: string): string {
  try {
    // Keep the BOM as part of the source so an edited CSV retains it too.
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes)
  } catch {
    return fail(`${label}: UTF-8 파일만 지원합니다. 원본을 UTF-8로 내보내 주세요.`)
  }
}

function parseXml(text: string, label: string): XMLDocument {
  if (/<!DOCTYPE/i.test(text)) fail(`${label}: DOCTYPE이 포함된 XML은 지원하지 않습니다.`)
  const declaration = text.match(/^\uFEFF?\s*<\?xml[^?]*encoding\s*=\s*["']([^"']+)/i)
  if (declaration && !/^utf-?8$/i.test(declaration[1])) {
    fail(`${label}: UTF-8 XML만 지원합니다.`)
  }
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  if (doc.documentElement.localName === 'parsererror' || doc.getElementsByTagName('parsererror').length) {
    fail(`${label}: XML 형식을 읽을 수 없습니다.`)
  }
  return doc
}

function children(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter(child => child.localName === localName && child.namespaceURI === element.namespaceURI)
}

function freezeSheet(sheet: SpreadsheetSheet): SpreadsheetSheet {
  for (const row of sheet.rows) Object.freeze(row)
  Object.freeze(sheet.rows)
  return Object.freeze(sheet)
}

function register(doc: SpreadsheetDocument, source: Source): SpreadsheetDocument {
  Object.freeze(doc.sheets)
  Object.freeze(doc)
  sources.set(doc, source)
  return doc
}

function detectDelimiter(text: string): string {
  const counts = new Map([ [',', 0], ['\t', 0], [';', 0] ])
  let quoted = false
  for (let index = text.charCodeAt(0) === 0xfeff ? 1 : 0; index < text.length; index++) {
    const char = text[index]
    if (char === '"') {
      if (quoted && text[index + 1] === '"') index++
      else quoted = !quoted
    } else if (!quoted && (char === '\r' || char === '\n')) break
    else if (!quoted && counts.has(char)) counts.set(char, counts.get(char)! + 1)
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0]
}

function parseCsv(text: string, delimiter: string): CsvRow[] {
  const rows: CsvRow[] = []
  let index = text.charCodeAt(0) === 0xfeff ? 1 : 0
  if (index === text.length) fail('CSV에 데이터가 없습니다.')
  while (index < text.length) {
    const tokens: CsvToken[] = []
    let rowEnded = false
    while (!rowEnded) {
      const start = index
      const quoted = text[index] === '"'
      let value = ''
      if (quoted) {
        index++
        let closed = false
        while (index < text.length) {
          const char = text[index++]
          if (char === '"') {
            if (text[index] === '"') { value += '"'; index++ }
            else { closed = true; break }
          } else value += char
        }
        if (!closed) fail('CSV의 큰따옴표가 닫히지 않았습니다.')
        if (index < text.length && ![delimiter, '\r', '\n'].includes(text[index])) {
          fail('CSV에서 닫는 큰따옴표 뒤에 잘못된 문자가 있습니다.')
        }
      } else {
        while (index < text.length && ![delimiter, '\r', '\n'].includes(text[index])) {
          if (text[index] === '"') fail('CSV의 큰따옴표 형식이 올바르지 않습니다.')
          value += text[index++]
        }
      }
      tokens.push({ start, end: index, value, quoted })
      if (tokens.length > MAX_COLUMNS) fail(`열은 최대 ${MAX_COLUMNS}개까지 지원합니다.`)
      if (text[index] === delimiter) { index++; continue }
      const end = index
      if (text[index] === '\r') { index++; if (text[index] === '\n') index++ }
      else if (text[index] === '\n') index++
      rows.push({ tokens, end })
      if (rows.length > MAX_ROWS) fail(`행은 최대 ${MAX_ROWS.toLocaleString()}개까지 지원합니다.`)
      rowEnded = true
    }
  }
  return rows
}

function loadCsv(name: string, bytes: Uint8Array): SpreadsheetDocument {
  const text = decodeUtf8(bytes, 'CSV')
  const delimiter = detectDelimiter(text)
  const rawRows = parseCsv(text, delimiter)
  const sheet = freezeSheet({
    id: 'csv-0', name: name.replace(/\.csv$/i, ''),
    rows: rawRows.map(row => row.tokens.map(token => token.value)),
    maxColumns: Math.max(...rawRows.map(row => row.tokens.length)),
  })
  return register({ kind: 'csv', name, sheets: [sheet] }, { kind: 'csv', bytes, text, delimiter, rows: rawRows, sheet })
}

function coordinate(reference: string): { row: number; column: number } {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(reference)
  if (!match) fail(`잘못된 셀 주소입니다: ${reference}`)
  let column = 0
  for (const char of match[1].toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row > 1_048_576 || column > 16_384) fail(`셀 주소 범위를 벗어났습니다: ${reference}`)
  return { row: row - 1, column: column - 1 }
}

export function columnLabel(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) { value--; label = String.fromCharCode(65 + value % 26) + label; value = Math.floor(value / 26) }
  return label
}

function cellKey(row: number, column: number): string { return `${columnLabel(column)}${row + 1}` }

function decodeExcelString(value: string): string {
  return value.replace(/_x([\da-f]{4})_/gi, (_, code: string) => String.fromCharCode(parseInt(code, 16)))
}

function encodeExcelString(value: string): string {
  return value.replace(/_x[\da-f]{4}_/gi, match => `_x005F_${match.slice(1)}`)
    .replace(/[\u0000-\u0008\u000b-\u001f\ufffe\uffff]/g, char => `_x${char.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`)
}

function excelText(element: Element): string {
  const texts = Array.from(element.getElementsByTagNameNS(element.namespaceURI, 't'))
    .filter(text => {
      let parent = text.parentElement
      while (parent && parent !== element) {
        if (parent.localName === 'rPh') return false
        parent = parent.parentElement
      }
      return true
    })
  return decodeExcelString(texts.map(text => text.textContent ?? '').join(''))
}

function resolvePath(base: string, target: string): string {
  const parts: string[] = target.startsWith('/') ? [] : base.split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (!part || part === '.') continue
    if (part === '..') { if (!parts.length) fail('XLSX 관계 경로가 올바르지 않습니다.'); parts.pop() }
    else parts.push(part)
  }
  return parts.join('/')
}

/** Locate only worksheet row/cell elements; replacements leave all other XML bytes alone. */
function worksheetSpans(xml: string): RowSpan[] {
  const rows: RowSpan[] = []
  const stack: { name: string; span: XmlSpan; row?: RowSpan }[] = []
  let position = 0
  while ((position = xml.indexOf('<', position)) !== -1) {
    const start = position
    if (xml.startsWith('<!--', start)) { position = xml.indexOf('-->', start + 4) + 3; continue }
    if (xml.startsWith('<![CDATA[', start)) { position = xml.indexOf(']]>', start + 9) + 3; continue }
    if (xml.startsWith('<?', start)) { position = xml.indexOf('?>', start + 2) + 2; continue }
    let quote = ''
    position++
    while (position < xml.length) {
      const char = xml[position++]
      if (quote) { if (char === quote) quote = '' }
      else if (char === '"' || char === "'") quote = char
      else if (char === '>') break
    }
    const tag = xml.slice(start, position)
    const match = /^<\/?([^\s/>]+)/.exec(tag)
    if (!match) continue
    const name = match[1].split(':').pop()!
    if (tag.startsWith('</')) {
      const open = stack.pop()
      if (open) { open.span.end = position; open.span.closeStart = start }
      continue
    }
    const span: XmlSpan = { start, end: position, closeStart: position, selfClosing: /\/\s*>$/.test(tag) }
    const parent = stack[stack.length - 1]
    let row: RowSpan | undefined
    if (name === 'row' && parent?.name === 'sheetData' && stack.length === 2) {
      row = { ...span, cells: [] }
      rows.push(row)
    } else if (name === 'c' && parent?.row) parent.row.cells.push(span)
    if (!span.selfClosing) stack.push({ name, span: row ?? span, row })
  }
  return rows
}

function loadXlsx(name: string, bytes: Uint8Array): SpreadsheetDocument {
  let entries: Record<string, Uint8Array>
  let total = 0
  try {
    entries = unzipSync(bytes, { filter: file => {
      total += file.originalSize
      if (total > MAX_UNZIPPED_BYTES) fail('압축 해제 크기가 100 MB를 초과합니다.')
      return true
    } })
  } catch (error) {
    fail(`XLSX 압축 파일을 읽을 수 없습니다: ${error instanceof Error ? error.message : '알 수 없는 오류'}`)
  }
  if (Object.values(entries).reduce((size, entry) => size + entry.length, 0) > MAX_UNZIPPED_BYTES) {
    fail('압축 해제 크기가 100 MB를 초과합니다.')
  }
  function readXml(path: string): { text: string; document: XMLDocument } {
    const entry = entries[path]
    if (!entry) fail(`XLSX 필수 파일이 없습니다: ${path}`)
    const text = decodeUtf8(entry, path)
    return { text, document: parseXml(text, path) }
  }
  const workbookPath = 'xl/workbook.xml'
  const workbook = readXml(workbookPath).document
  const relations = readXml('xl/_rels/workbook.xml.rels').document
  const relationMap = new Map<string, Element>()
  for (const relation of Array.from(relations.documentElement.children)) {
    const id = relation.getAttribute('Id')
    if (!id || relationMap.has(id)) fail('XLSX 관계 ID가 없거나 중복되어 있습니다.')
    relationMap.set(id, relation)
  }
  const sharedStrings = entries['xl/sharedStrings.xml']
    ? children(readXml('xl/sharedStrings.xml').document.documentElement, 'si').map(excelText) : []
  const sheets: SpreadsheetSheet[] = []
  const sheetSources = new Map<string, XlsxSheetSource>()
  const worksheetPaths = new Set<string>()
  const workbookSheets = children(workbook.documentElement, 'sheets')[0]
  if (!workbookSheets) fail('XLSX에 워크시트가 없습니다.')
  for (const [sheetIndex, entry] of children(workbookSheets, 'sheet').entries()) {
    if (sheets.length >= 100) fail('워크시트는 최대 100개까지 지원합니다.')
    const relationId = Array.from(entry.attributes).find(attribute => attribute.localName === 'id')?.value
    const relation = relationMap.get(relationId ?? '')
    if (!relation || relation.getAttribute('TargetMode') === 'External') fail('XLSX 워크시트 연결을 읽을 수 없습니다.')
    if (!relation.getAttribute('Type')?.endsWith('/worksheet')) fail('일반 워크시트만 지원합니다.')
    const path = resolvePath(workbookPath, relation.getAttribute('Target') ?? '')
    if (worksheetPaths.has(path)) fail('여러 워크시트가 동일한 원본 파일을 가리키고 있습니다.')
    worksheetPaths.add(path)
    const xml = readXml(path)
    const sheetData = children(xml.document.documentElement, 'sheetData')[0]
    if (!sheetData) fail('워크시트의 sheetData가 없습니다.')
    const spans = worksheetSpans(xml.text)
    const rowElements = new Map<number, Element>()
    const rowSpans = new Map<number, RowSpan>()
    const cells: XlsxSheetSource['cells'] = new Map()
    const formulas = new Set<string>()
    const rows: string[][] = []
    let maxColumns = 0
    let previousRow = -1
    const xmlRows = children(sheetData, 'row')
    if (spans.length !== xmlRows.length) fail('워크시트 행 구조가 올바르지 않습니다.')
    for (const [rowOffset, row] of xmlRows.entries()) {
      const rowNumber = row.getAttribute('r')
      const rowIndex = rowNumber ? Number(rowNumber) - 1 : previousRow + 1
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= MAX_ROWS) fail(`행은 최대 ${MAX_ROWS.toLocaleString()}개까지 지원합니다.`)
      if (rowElements.has(rowIndex) || rowIndex <= previousRow) fail('워크시트의 행 주소가 중복되거나 순서가 잘못되었습니다.')
      previousRow = rowIndex
      rowElements.set(rowIndex, row)
      rowSpans.set(rowIndex, spans[rowOffset])
      while (rows.length <= rowIndex) rows.push([])
      let previousColumn = -1
      const xmlCells = children(row, 'c')
      if (xmlCells.length !== spans[rowOffset].cells.length) fail('워크시트 셀 구조가 올바르지 않습니다.')
      for (const [cellOffset, cell] of xmlCells.entries()) {
        const address = cell.getAttribute('r')
        const at = address ? coordinate(address) : { row: rowIndex, column: previousColumn + 1 }
        if (at.row !== rowIndex || at.column <= previousColumn) fail('셀 주소가 중복되거나 행과 일치하지 않습니다.')
        if (at.column >= MAX_COLUMNS) fail(`열은 최대 ${MAX_COLUMNS}개까지 지원합니다.`)
        previousColumn = at.column
        const key = cellKey(rowIndex, at.column)
        cells.set(key, { element: cell, span: spans[rowOffset].cells[cellOffset] })
        if (children(cell, 'f').length) formulas.add(key)
        const type = cell.getAttribute('t')
        const rawValue = children(cell, 'v')[0]?.textContent ?? ''
        let value = rawValue
        if (type === 's') {
          const sharedIndex = Number(rawValue)
          if (!/^\d+$/.test(rawValue) || !Number.isSafeInteger(sharedIndex) || sharedStrings[sharedIndex] === undefined) fail('공유 문자열 참조가 올바르지 않습니다.')
          value = sharedStrings[sharedIndex]
        } else if (type === 'inlineStr') value = children(cell, 'is')[0] ? excelText(children(cell, 'is')[0]) : ''
        else if (type === 'str') value = decodeExcelString(rawValue)
        else if (type === 'b') value = rawValue === '1' ? 'TRUE' : rawValue === '0' ? 'FALSE' : rawValue
        while (rows[rowIndex].length <= at.column) rows[rowIndex].push('')
        rows[rowIndex][at.column] = value
        maxColumns = Math.max(maxColumns, at.column + 1)
      }
    }
    const merges: Merge[] = []
    const mergeCells = children(xml.document.documentElement, 'mergeCells')[0]
    if (mergeCells) for (const merge of children(mergeCells, 'mergeCell')) {
      const ref = merge.getAttribute('ref') ?? ''
      const pair = ref.split(':')
      if (pair.length > 2) fail('병합 셀 주소가 올바르지 않습니다.')
      const start = coordinate(pair[0])
      const end = coordinate(pair[1] ?? pair[0])
      if (end.row < start.row || end.column < start.column) fail('병합 셀 주소가 올바르지 않습니다.')
      merges.push({ top: start.row, left: start.column, bottom: end.row, right: end.column })
    }
    const sheet = freezeSheet({ id: `xlsx-${sheetIndex}`, name: entry.getAttribute('name') ?? `Sheet ${sheetIndex + 1}`, rows, maxColumns })
    sheets.push(sheet)
    sheetSources.set(sheet.id, { sheet, path, xml: xml.text, document: xml.document, rowElements, rowSpans, cells, formulas, merges })
  }
  if (!sheets.length) fail('XLSX에 워크시트가 없습니다.')
  return register({ kind: 'xlsx', name, sheets }, { kind: 'xlsx', bytes, entries, sheets: sheetSources })
}

export function loadSpreadsheet(name: string, bytes: Uint8Array): SpreadsheetDocument {
  if (!bytes.length) fail('비어 있는 파일입니다.')
  if (bytes.length > MAX_FILE_BYTES) fail('파일은 최대 25 MB까지 지원합니다.')
  const original = new Uint8Array(bytes)
  if (/\.csv$/i.test(name)) return loadCsv(name, original)
  if (/\.xlsx$/i.test(name)) return loadXlsx(name, original)
  return fail('UTF-8 CSV 또는 XLSX 파일을 선택해 주세요. XLS 파일은 지원하지 않습니다.')
}

export function findCellIssue(doc: SpreadsheetDocument, sheetId: string, rowIndex: number, columnIndex: number): string | null {
  const source = sources.get(doc)
  if (!source) return '이 도구에서 읽은 원본 파일이 아닙니다.'
  const sheet = doc.sheets.find(item => item.id === sheetId)
  if (!sheet) return '워크시트를 찾을 수 없습니다.'
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= sheet.rows.length) return '원본 행 범위를 벗어났습니다.'
  if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= sheet.maxColumns) return '원본 열 범위를 벗어났습니다.'
  if (source.kind === 'xlsx') {
    const data = source.sheets.get(sheetId)!
    if (!data.rowElements.has(rowIndex)) return '원본에 존재하지 않는 행은 편집할 수 없습니다.'
    if (data.formulas.has(cellKey(rowIndex, columnIndex))) return '수식 셀은 원본 보호를 위해 편집할 수 없습니다.'
    if (data.merges.some(merge => rowIndex >= merge.top && rowIndex <= merge.bottom && columnIndex >= merge.left && columnIndex <= merge.right)) {
      return '병합 셀은 원본 보호를 위해 편집할 수 없습니다.'
    }
  }
  return null
}

export function assertEditableCell(doc: SpreadsheetDocument, sheetId: string, rowIndex: number, columnIndex: number): void {
  const issue = findCellIssue(doc, sheetId, rowIndex, columnIndex)
  if (issue) fail(issue)
}

function applyPatches(text: string, patches: (Span & { text: string })[]): string {
  patches.sort((left, right) => right.start - left.start)
  let boundary = text.length
  for (const patch of patches) {
    if (patch.end > boundary) fail('동일한 원본 영역에 편집이 겹쳤습니다.')
    text = text.slice(0, patch.start) + patch.text + text.slice(patch.end)
    boundary = patch.start
  }
  return text
}

function csvToken(value: string, delimiter: string, quoted: boolean): string {
  return quoted || value.includes(delimiter) || /["\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value
}

function exportCsv(source: CsvSource, edits: CellEdit[]): Uint8Array {
  const patches: (Span & { text: string })[] = []
  const byRow = new Map<number, CellEdit[]>()
  for (const edit of edits) {
    const row = byRow.get(edit.rowIndex) ?? []
    row.push(edit)
    byRow.set(edit.rowIndex, row)
  }
  for (const [rowIndex, rowEdits] of byRow) {
    const row = source.rows[rowIndex]
    const extra = rowEdits.filter(edit => edit.columnIndex >= row.tokens.length)
    for (const edit of rowEdits.filter(edit => edit.columnIndex < row.tokens.length)) {
      const token = row.tokens[edit.columnIndex]
      patches.push({ start: token.start, end: token.end, text: csvToken(edit.value, source.delimiter, token.quoted) })
    }
    if (extra.length) {
      const last = Math.max(...extra.map(edit => edit.columnIndex))
      const values = Array.from({ length: last - row.tokens.length + 1 }, () => '')
      for (const edit of extra) values[edit.columnIndex - row.tokens.length] = csvToken(edit.value, source.delimiter, false)
      patches.push({ start: row.end, end: row.end, text: source.delimiter + values.join(source.delimiter) })
    }
  }
  return encoder.encode(applyPatches(source.text, patches))
}

function textCell(document: XMLDocument, original: Element | undefined, row: Element, edit: CellEdit): Element {
  const namespace = row.namespaceURI
  const prefix = row.prefix ? `${row.prefix}:` : ''
  const cell = original ? original.cloneNode(true) as Element : document.createElementNS(namespace, `${prefix}c`)
  if (!cell.hasAttribute('r')) cell.setAttribute('r', cellKey(edit.rowIndex, edit.columnIndex))
  cell.setAttribute('t', 'inlineStr')
  for (const child of Array.from(cell.children)) {
    if (child.namespaceURI === namespace && ['v', 'is', 'f'].includes(child.localName)) cell.removeChild(child)
  }
  const inline = document.createElementNS(namespace, `${prefix}is`)
  const text = document.createElementNS(namespace, `${prefix}t`)
  text.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  text.textContent = encodeExcelString(edit.value)
  inline.appendChild(text)
  cell.insertBefore(inline, cell.firstChild)
  return cell
}

function exportXlsx(source: XlsxSource, edits: CellEdit[]): Uint8Array {
  const output = { ...source.entries }
  const bySheet = new Map<string, CellEdit[]>()
  for (const edit of edits) {
    const group = bySheet.get(edit.sheetId) ?? []
    group.push(edit)
    bySheet.set(edit.sheetId, group)
  }
  const serializer = new XMLSerializer()
  for (const [sheetId, sheetEdits] of bySheet) {
    const data = source.sheets.get(sheetId)!
    const patches: (Span & { text: string })[] = []
    const insertions = new Map<number, { column: number; xml: string }[]>()
    const emptyRows = new Map<number, CellEdit[]>()
    for (const edit of sheetEdits) {
      const key = cellKey(edit.rowIndex, edit.columnIndex)
      const existing = data.cells.get(key)
      const row = data.rowElements.get(edit.rowIndex)!
      const xml = serializer.serializeToString(textCell(data.document, existing?.element, row, edit))
      if (existing) patches.push({ ...existing.span, text: xml })
      else {
        const rowSpan = data.rowSpans.get(edit.rowIndex)!
        if (rowSpan.selfClosing) {
          const group = emptyRows.get(edit.rowIndex) ?? []
          group.push(edit)
          emptyRows.set(edit.rowIndex, group)
        } else {
          let position = rowSpan.closeStart
          for (let column = edit.columnIndex + 1; column < data.sheet.maxColumns; column++) {
            const nextCell = data.cells.get(cellKey(edit.rowIndex, column))
            if (nextCell) { position = nextCell.span.start; break }
          }
          if (position === undefined) fail('빈 셀의 삽입 위치를 찾을 수 없습니다.')
          const group = insertions.get(position) ?? []
          group.push({ column: edit.columnIndex, xml })
          insertions.set(position, group)
        }
      }
    }
    for (const [position, group] of insertions) patches.push({ start: position, end: position, text: group.sort((a, b) => a.column - b.column).map(item => item.xml).join('') })
    for (const [rowIndex, group] of emptyRows) {
      const row = data.rowElements.get(rowIndex)!.cloneNode(true) as Element
      for (const edit of group.sort((a, b) => a.columnIndex - b.columnIndex)) row.appendChild(textCell(data.document, undefined, row, edit))
      patches.push({ ...data.rowSpans.get(rowIndex)!, text: serializer.serializeToString(row) })
    }
    output[data.path] = encoder.encode(applyPatches(data.xml, patches))
  }
  return zipSync(output, { level: 6 })
}

export function exportSpreadsheet(doc: SpreadsheetDocument, edits: CellEdit[]): Uint8Array {
  const source = sources.get(doc)
  if (!source) fail('이 도구에서 읽은 원본 파일이 아닙니다.')
  const seen = new Set<string>()
  const changed: CellEdit[] = []
  for (const edit of edits) {
    assertEditableCell(doc, edit.sheetId, edit.rowIndex, edit.columnIndex)
    const key = `${edit.sheetId}:${edit.rowIndex}:${edit.columnIndex}`
    if (seen.has(key)) fail('동일한 셀에 편집이 중복되어 있습니다.')
    seen.add(key)
    if (typeof edit.value !== 'string') fail('셀 값은 문자열이어야 합니다.')
    if (/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(edit.value)) fail('셀에 유효하지 않은 유니코드 문자가 있습니다.')
    if (source.kind === 'xlsx' && edit.value.length > 32_767) fail('XLSX 셀은 최대 32,767자까지 지원합니다.')
    const sheet = doc.sheets.find(item => item.id === edit.sheetId)!
    if ((sheet.rows[edit.rowIndex][edit.columnIndex] ?? '') !== edit.value) changed.push(edit)
  }
  if (!changed.length) return new Uint8Array(source.bytes)
  return source.kind === 'csv' ? exportCsv(source, changed) : exportXlsx(source, changed)
}
