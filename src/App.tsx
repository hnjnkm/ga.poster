import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent } from 'react';
import { ArrowDownToLine, ArrowLeft, ArrowRight, Check, CheckCheck, ChevronDown, CircleHelp, FileSpreadsheet, FolderOpen, ImagePlus, Languages, Link2, LoaderCircle, LockKeyhole, Plus, RotateCcw, Save, Search, ShieldCheck, Sparkles, Trash2, Upload, X } from 'lucide-react';
import { analyzeMapping, compareCanonicalIds, parseImageName, suggestMapping, type ColumnMapping } from './lib/matching';
import { assertEditableCell, loadSpreadsheet, type CellEdit, type SpreadsheetDocument } from './lib/spreadsheet';
import ArchiveCard from './components/ArchiveCard';
import { verifiedExport } from './lib/verified-export';
import { readSession, saveSession, type SavedSession } from './lib/session';

const emptyMapping: ColumnMapping = { headerRow: 0, idColumn: -1, koColumn: -1, enColumn: -1, zhColumn: -1 };
const languageFields = [
  { key: 'koColumn', label: '한국어', code: 'KO', lang: 'ko' },
  { key: 'enColumn', label: 'English', code: 'EN', lang: 'en' },
  { key: 'zhColumn', label: '简体中文', code: 'ZH', lang: 'zh-Hans' },
] as const;
const errorMessage = (error: unknown) => error instanceof Error ? error.message : '작업을 완료하지 못했습니다.';
const cellKey = (row: number, column: number) => `${row}:${column}`;

export default function App() {
  const [workbook, setWorkbook] = useState<{ name: string; bytes: Uint8Array } | null>(null);
  const [document, setDocument] = useState<SpreadsheetDocument | null>(null);
  const [images, setImages] = useState<File[]>([]);
  const [sheetId, setSheetId] = useState('');
  const [mapping, setMapping] = useState<ColumnMapping>(emptyMapping);
  const [confirmed, setConfirmed] = useState(false);
  const [edits, setEdits] = useState<Map<string, CellEdit>>(new Map());
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState('all');
  const [notice, setNotice] = useState<{ text: string; error?: boolean } | null>(null);
  const [saved, setSaved] = useState<SavedSession>();
  const [busy, setBusy] = useState('');
  const [reviewing, setReviewing] = useState(false);
  const [showIssues, setShowIssues] = useState(false);
  const [showHelp, setShowHelp] = useState(false);
  const [sessionDirty, setSessionDirty] = useState(false);
  const [lastSaved, setLastSaved] = useState('');
  const imageInput = useRef<HTMLInputElement>(null);
  const spreadsheetInput = useRef<HTMLInputElement>(null);

  const sheet = document?.sheets.find(item => item.id === sheetId);
  const analysis = useMemo(() => sheet && confirmed ? analyzeMapping(sheet.rows, mapping, images.map(file => file.name)) : null, [sheet, confirmed, mapping, images]);
  const matchByImage = useMemo(() => new Map(analysis?.matches.map(match => [match.imageIndex, match]) ?? []), [analysis]);
  const modifiedRows = useMemo(() => new Set([...edits.values()].map(edit => edit.rowIndex)), [edits]);
  const orderedImages = useMemo(() => images.map((file, imageIndex) => ({ file, imageIndex, parsed: parseImageName(file.name) }))
    .sort((a, b) => a.parsed && b.parsed ? compareCanonicalIds(a.parsed.id, b.parsed.id) || a.imageIndex - b.imageIndex : a.parsed ? -1 : b.parsed ? 1 : a.imageIndex - b.imageIndex), [images]);
  const visibleImages = orderedImages.filter(item => {
    const match = matchByImage.get(item.imageIndex);
    return item.file.name.toLowerCase().includes(search.toLowerCase()) && (filter === 'all' || (filter === 'edited' ? match && modifiedRows.has(match.rowIndex) : !match));
  });
  const canExport = Boolean(document && analysis?.valid);
  const matchedCount = analysis?.matches.length ?? 0;

  useEffect(() => { readSession().then(setSaved).catch(() => setNotice({ text: '브라우저의 로컬 저장 공간을 사용할 수 없습니다. 작업 파일은 다운로드할 수 있습니다.', error: true })); }, []);
  useEffect(() => {
    const listener = (event: BeforeUnloadEvent) => { if (sessionDirty) { event.preventDefault(); event.returnValue = ''; } };
    window.addEventListener('beforeunload', listener);
    return () => window.removeEventListener('beforeunload', listener);
  }, [sessionDirty]);
  useEffect(() => {
    if (!reviewing) return;
    const handler = (event: KeyboardEvent) => { if (event.key === 'Escape') setReviewing(false); };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [reviewing]);

  function setMappingFromHeaders(rows: string[][], headerRow = 0) {
    setMapping({ ...emptyMapping, headerRow, ...suggestMapping(rows[headerRow] ?? []) });
  }

  async function importSpreadsheet(file?: File) {
    if (!file || busy) return;
    if (document && !window.confirm('새 설명 파일을 열면 현재 파일의 편집 내용과 열 연결이 교체됩니다. 필요한 작업을 먼저 저장하거나 내려받았나요?')) return;
    setBusy('설명 파일을 읽고 있습니다');
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const next = loadSpreadsheet(file.name, bytes);
      if (!next.sheets.length) throw new Error('읽을 수 있는 시트가 없습니다.');
      setWorkbook({ name: file.name, bytes });
      setDocument(next);
      setSheetId(next.sheets[0].id);
      setMappingFromHeaders(next.sheets[0].rows);
      setConfirmed(false);
      setEdits(new Map());
      setSessionDirty(true);
      setNotice({ text: '설명 파일을 불러왔습니다. 순번과 언어별 열을 확인한 뒤 연결해 주세요.' });
    } catch (error) { setNotice({ text: errorMessage(error), error: true }); }
    finally { setBusy(''); }
  }

  function addImages(files: FileList | File[] | null) {
    if (!files?.length || busy) return;
    const added = Array.from(files);
    setImages(current => [...current, ...added]);
    setSessionDirty(true);
    setNotice({ text: `${added.length}개의 이미지를 추가했습니다. 같은 순번이 중복되면 내보내기를 차단합니다.` });
  }

  const editCell = useCallback((row: number, column: number, value: string) => {
    if (!document || !sheet || busy) return;
    try { assertEditableCell(document, sheetId, row, column); }
    catch (error) { setNotice({ text: errorMessage(error), error: true }); return; }
    setEdits(current => {
      const next = new Map(current);
      if (value === (sheet.rows[row]?.[column] ?? '')) next.delete(cellKey(row, column));
      else next.set(cellKey(row, column), { sheetId, rowIndex: row, columnIndex: column, value });
      return next;
    });
    setSessionDirty(true);
  }, [document, sheet, sheetId, busy]);

  const removeImage = useCallback((index: number, name: string) => {
    if (busy) return;
    if (window.confirm(`${name} 이미지를 작업 목록에서 제거할까요? 원본 파일과 설명은 유지됩니다.`)) {
      setImages(current => current.filter((_, currentIndex) => currentIndex !== index));
      setSessionDirty(true);
    }
  }, [busy]);

  async function saveWork() {
    setBusy('브라우저에 작업을 저장하고 있습니다');
    try {
      const session: SavedSession = { version: 1, savedAt: new Date().toISOString(), workbook, images, sheetId, mapping, confirmed, edits: [...edits.values()] };
      await saveSession(session);
      setSaved(session);
      setLastSaved(session.savedAt);
      setSessionDirty(false);
      setNotice({ text: '이 브라우저에 이미지, 원본 파일, 편집 내용을 저장했습니다. 다른 기기에는 동기화되지 않습니다.' });
    } catch (error) { setNotice({ text: `작업 저장 실패: ${errorMessage(error)}. 다운로드하거나 브라우저 저장 공간을 확인해 주세요.`, error: true }); }
    finally { setBusy(''); }
  }

  function restoreWork() {
    if (busy || !saved || !window.confirm('저장된 작업을 복원합니다. 현재 화면의 저장하지 않은 작업이 있다면 교체됩니다.')) return;
    try {
      if (saved.version !== 1) throw new Error('지원하지 않는 작업 저장 버전입니다.');
      const next = saved.workbook ? loadSpreadsheet(saved.workbook.name, new Uint8Array(saved.workbook.bytes)) : null;
      if (next && !next.sheets.some(item => item.id === saved.sheetId)) throw new Error('저장된 시트 정보를 찾을 수 없습니다.');
      for (const edit of saved.edits) { if (!next) throw new Error('편집 내용의 원본 파일이 없습니다.'); assertEditableCell(next, edit.sheetId, edit.rowIndex, edit.columnIndex); }
      setWorkbook(saved.workbook);
      setDocument(next);
      setImages(saved.images);
      setSheetId(saved.sheetId);
      setMapping(saved.mapping);
      setConfirmed(saved.confirmed);
      setEdits(new Map(saved.edits.map(edit => [cellKey(edit.rowIndex, edit.columnIndex), edit])));
      setLastSaved(saved.savedAt);
      setSessionDirty(false);
      setNotice({ text: '저장한 작업을 복원했습니다.' });
    } catch (error) { setNotice({ text: errorMessage(error), error: true }); }
  }

  async function downloadWork() {
    if (!document || !analysis?.valid) return;
    setBusy('수정 내용과 원본 행 순서를 검증하고 있습니다');
    try {
      await new Promise(resolve => setTimeout(resolve, 0));
      const changes = [...edits.values()];
      const bytes = verifiedExport(document, changes);
      const extension = document.kind === 'xlsx' ? 'xlsx' : 'csv';
      const blob = new Blob([new Uint8Array(bytes).buffer], { type: extension === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = window.document.createElement('a');
      link.href = url;
      link.download = document.name.replace(/\.[^.]+$/, '') + '-edited.' + extension;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      setReviewing(false);
      setNotice({ text: `원본 대조를 통과했습니다. ${changes.length}개 셀을 반영한 ${extension.toUpperCase()} 파일을 다운로드했습니다.` });
    } catch (error) { setNotice({ text: errorMessage(error), error: true }); }
    finally { setBusy(''); }
  }

  function resetEdits() {
    if (window.confirm('현재 파일의 수정 내용을 모두 원본으로 되돌릴까요? 이미지와 원본 파일은 유지됩니다.')) {
      setEdits(new Map()); setSessionDirty(true); setReviewing(false);
    }
  }

  const changeFile = (event: ChangeEvent<HTMLInputElement>) => { void importSpreadsheet(event.target.files?.[0]); event.target.value = ''; };

  return <div className="app-shell" aria-busy={Boolean(busy)}>
    <header className="topbar">
      <a className="wordmark" href="#" aria-label="Graphic Archive 작업 공간"><span className="brand-symbol">g<span>a</span></span><span>GRAPHIC<br/>ARCHIVE<span className="brand-dot">®</span></span></a>
      <div className="workspace-label"><span className="divider"/> WORKSPACE <span className="version">01</span></div>
      <div className="header-actions"><button className="icon-button" aria-label="사용 방법" onClick={() => setShowHelp(!showHelp)}><CircleHelp size={19}/></button><button className="button subtle" onClick={saveWork} disabled={Boolean(busy) || (!images.length && !document)}><Save size={16}/> 작업 저장</button><button className="button primary" onClick={() => setReviewing(true)} disabled={!canExport || Boolean(busy)}><ArrowDownToLine size={16}/> 내보내기 <span className="button-count">{edits.size}</span></button></div>
    </header>

    {showHelp && <section className="help-bar"><strong>작업 순서</strong><span>① 이미지와 기존 CSV / XLSX 파일 업로드</span><ArrowRight size={14}/><span>② 순번·언어 열 연결 확인</span><ArrowRight size={14}/><span>③ 설명 편집 후 검토·다운로드</span><button className="icon-button" aria-label="사용 방법 닫기" onClick={() => setShowHelp(false)}><X size={16}/></button></section>}
    {notice && <div className={`notice ${notice.error ? 'error' : ''}`} role={notice.error ? 'alert' : 'status'}><span>{notice.error ? <CircleHelp size={16}/> : <Check size={16}/>} {notice.text}</span><button className="icon-button" aria-label="알림 닫기" onClick={() => setNotice(null)}><X size={15}/></button></div>}
    {busy && <div className="busy-bar" role="status"><LoaderCircle className="spin" size={16}/>{busy}</div>}

    <div className="workspace-layout" inert={Boolean(busy)}>
      <aside className="sidebar">
        <div className="sidebar-heading"><span className="eyebrow">YOUR COLLECTION</span><h1>아카이브 작업실<span className="heading-dot"/></h1><p>이미지와 세 언어의 설명을<br/>한곳에서 정리하세요.</p></div>
        <section className="upload-section"><div className="section-label"><span>01</span> 이미지 불러오기</div><button className="upload-box" onClick={() => imageInput.current?.click()} onDragOver={event => event.preventDefault()} onDrop={event => { event.preventDefault(); addImages(event.dataTransfer.files); }}><div className="upload-icon"><ImagePlus size={23} strokeWidth={1.4}/></div><strong>이미지를 여기에 놓아주세요</strong><span>또는 클릭하여 파일 선택</span><small>01.png, 02.jpg, 1002.png · 여러 장 선택</small></button><input ref={imageInput} type="file" accept=".png,.jpg,.jpeg,.webp,.avif,.gif" multiple hidden aria-label="이미지 업로드" onChange={event => { addImages(event.target.files); event.target.value = ''; }}/>{images.length > 0 && <div className="file-status"><CheckCheck size={14}/><span>{images.length.toLocaleString()}개 이미지</span><button onClick={() => imageInput.current?.click()} className="text-button"><Plus size={13}/>추가</button></div>}</section>

        <section className="upload-section"><div className="section-label"><span>02</span> 설명 파일 불러오기</div><button className={`spreadsheet-upload ${document ? 'has-file' : ''}`} onClick={() => spreadsheetInput.current?.click()}><FileSpreadsheet size={23} strokeWidth={1.5}/><span><strong>{document ? document.name : '기존 엑셀 · CSV 파일'}</strong><small>{document ? `${document.kind.toUpperCase()} · 원본 보관 중` : 'Webflow에서 사용하는 양식 그대로'}</small></span>{document ? <Check size={16}/> : <Plus size={16}/>}</button><input ref={spreadsheetInput} type="file" accept=".csv,.xlsx" hidden aria-label="설명 파일 업로드" onChange={changeFile}/><p className="small-note">UTF-8 CSV · XLSX 지원. 실제 양식을 바꾸거나 새 열을 만들지 않습니다.</p></section>

        {document && sheet && <section className="mapping-section"><div className="section-label"><span>03</span> 열 연결 확인{confirmed && <Check size={14}/>}</div><label className="field-label">워크시트<select aria-label="워크시트" disabled={edits.size > 0} value={sheetId} onChange={event => { const selected = document.sheets.find(item => item.id === event.target.value)!; setSheetId(selected.id); setMappingFromHeaders(selected.rows); setConfirmed(false); setSessionDirty(true); }} >{document.sheets.map(item => <option key={item.id} value={item.id}>{item.name}</option>)}</select></label><label className="field-label">제목 행 번호<input aria-label="제목 행 번호" type="number" min="1" max={Math.max(1, sheet.rows.length)} disabled={edits.size > 0} value={mapping.headerRow + 1} onChange={event => { const row = Math.max(0, Math.min(sheet.rows.length - 1, Number(event.target.value) - 1)); setMappingFromHeaders(sheet.rows, row); setConfirmed(false); setSessionDirty(true); }}/></label>
          {([['idColumn', '순번'], ...languageFields.map(item => [item.key, item.label])] as [keyof Omit<ColumnMapping, 'headerRow'>, string][]).map(([key, label]) => <label className="field-label mapping-row" key={key}><span>{label}</span><select aria-label={`${label} 열`} value={mapping[key]} disabled={edits.size > 0} onChange={event => { setMapping(current => ({ ...current, [key]: Number(event.target.value) })); setConfirmed(false); setSessionDirty(true); }}><option value={-1}>열 선택</option>{(sheet.rows[mapping.headerRow] ?? []).map((header, index) => <option key={index} value={index}>{index + 1}. {header || '(제목 없음)'}</option>)}</select></label>)}
          <button className={`button ${confirmed ? 'confirmed' : 'primary'} full`} onClick={() => { setConfirmed(true); setSessionDirty(true); setShowIssues(true); }} disabled={Object.entries(mapping).some(([, value]) => value < 0) || edits.size > 0}>{confirmed ? <ShieldCheck size={15}/> : <Link2 size={15}/>} {confirmed ? '열 연결 확인됨' : '이 열로 연결하기'}</button><p className="small-note">자동 제안된 열도 직접 확인해 주세요. 순번 외의 정보로는 연결하지 않습니다.</p>
        </section>}

        <section className="ai-note"><div><Sparkles size={16}/><strong>AI 설명 작성</strong><span>준비 예정</span></div><p>작성 지침을 받은 뒤 연결합니다.<br/>현재는 기존 설명을 불러와 편집할 수 있어요.</p></section>
        <div className="sidebar-bottom"><LockKeyhole size={14}/><p>이미지와 파일은 이 브라우저에서 처리됩니다. 서버로 전송하지 않습니다.</p></div>
        {saved && <button className="restore-button" onClick={restoreWork}><FolderOpen size={15}/><span>저장한 작업 복원<small>{new Date(saved.savedAt).toLocaleString('ko-KR')}</small></span><ArrowRight size={14}/></button>}
      </aside>

      <main className="main-panel">
        <div className="collection-heading"><div><span className="eyebrow">COLLECTION EDITOR</span><h2>이미지와 설명<span className="muted-number">{images.length.toString().padStart(2, '0')}</span></h2></div><div className="local-status"><span/> {sessionDirty ? '저장하지 않은 작업' : lastSaved ? '브라우저에 저장됨' : '로컬 작업 공간'}</div></div>
        <div className="stats-row"><div><span>전체 이미지</span><strong>{images.length.toLocaleString()}<small>장</small></strong></div><div><span>연결된 이미지</span><strong>{matchedCount.toLocaleString()}<small>장</small></strong></div><div><span>수정한 설명</span><strong>{edits.size.toLocaleString()}<small>셀</small></strong></div><button className="integrity-stat" onClick={() => setShowIssues(!showIssues)}><ShieldCheck size={23} strokeWidth={1.4}/><span><strong>{analysis ? analysis.valid ? '연결 검사 통과' : '연결 확인 필요' : '원본을 기준으로 연결'}</strong><small>{analysis ? `이미지 없는 행 ${analysis.unmatchedRows.length} · 미연결 이미지 ${analysis.unmatchedImages.length}` : '행 순서와 다른 열은 그대로'}</small></span><ChevronDown size={15}/></button></div>

        {showIssues && analysis && <section className={`validation-panel ${analysis.valid ? '' : 'invalid'}`}><div className="validation-title"><ShieldCheck size={17}/><strong>{analysis.valid ? '중복·형식 검사를 통과했습니다' : '다음 항목을 확인해 주세요'}</strong><button className="icon-button" aria-label="연결 검사 닫기" onClick={() => setShowIssues(false)}><X size={15}/></button></div>{analysis.errors.map((issue, index) => <p key={index}>{issue}</p>)}<p>이미지가 없는 설명 행 {analysis.unmatchedRows.length}개는 원본 그대로 유지합니다. 미연결 이미지 {analysis.unmatchedImages.length}개에는 설명을 입력할 수 없습니다.</p>{analysis.invalidRows.map(item => <p key={`row-${item.rowIndex}`}>원본 {item.rowIndex + 1}행 · 순번: {JSON.stringify(item.value)} — {item.reason}</p>)}
          {analysis.invalidImages.map(item => <p key={`image-${item.imageIndex}`}>{item.imageName} — {item.reason}</p>)}
          {analysis.duplicateRowIds.map(item => <p key={`duplicate-${item.id}`}>순번 {item.id} 중복: 원본 {item.rowIndices.map(index => index + 1).join(', ')}행</p>)}
          {analysis.duplicateImageIds.map(item => <p key={`duplicate-image-${item.id}`}>중복 이미지: {item.imageNames.join(', ')}</p>)}
          <p>화면의 이미지는 순번으로 정렬되며, 다운로드 파일의 행 순서는 원본과 같습니다.</p></section>}

        <div className="editor-toolbar"><div className="tabs" role="group" aria-label="이미지 필터">{[['all', '전체'], ['edited', '수정됨'], ['unmatched', '미연결']].map(([key, label]) => <button key={key} className={filter === key ? 'active' : ''} onClick={() => setFilter(key)}>{label}{key === 'all' && <span>{images.length}</span>}</button>)}</div><div className="toolbar-right"><span className="sort-label">순번 오름차순</span><label className="search-box"><Search size={15}/><input placeholder="파일명 검색" aria-label="파일명 검색" value={search} onChange={event => setSearch(event.target.value)}/></label></div></div>

        {images.length === 0 ? <section className="empty-state"><div className="empty-art" aria-hidden="true"><div className="art-card art-card-back"><span>EN</span><i/><i/><i/></div><div className="art-card art-card-front"><span>GA — 001</span><div className="poster-mark">g<span>a</span></div><small>THE VISUAL ARCHIVE</small></div><div className="art-caption"><Languages size={16}/><span>KO</span><span>EN</span><span>ZH</span></div></div><span className="eyebrow">A PLACE FOR EVERY IMAGE</span><h3>첫 이미지를 불러와 보세요.</h3><p>순번이 적힌 이미지와 설명 파일을 연결하면<br/>세 언어의 설명을 나란히 확인할 수 있습니다.</p><button className="button primary" onClick={() => imageInput.current?.click()}><Upload size={16}/> 이미지 불러오기</button><div className="empty-footnote"><ShieldCheck size={14}/>원본 파일은 유지하고, 수정 내용만 반영합니다.</div></section> : visibleImages.length === 0 ? <section className="empty-results"><Search size={24}/><h3>해당하는 이미지가 없습니다.</h3><p>검색어나 필터를 변경해 주세요.</p><button className="text-button" onClick={() => { setSearch(''); setFilter('all'); }}>전체 이미지 보기 <ArrowRight size={14}/></button></section> : <div className="image-list">{visibleImages.map(({ file, imageIndex, parsed }) => {
          const match = matchByImage.get(imageIndex);
          const row = match?.rowIndex;
          return <ArchiveCard key={`${imageIndex}-${file.name}-${file.lastModified}`} file={file} imageIndex={imageIndex} parsed={parsed} match={match} document={document} sheet={sheet} sheetId={sheetId} mapping={mapping}
            koEdit={row === undefined ? undefined : edits.get(cellKey(row, mapping.koColumn))}
            enEdit={row === undefined ? undefined : edits.get(cellKey(row, mapping.enColumn))}
            zhEdit={row === undefined ? undefined : edits.get(cellKey(row, mapping.zhColumn))}
            onEdit={editCell} onRemove={removeImage}/>;
        })}</div>}
        <footer className="workspace-footer"><span>GRAPHIC ARCHIVE <span>WORKSPACE</span></span><p><ShieldCheck size={13}/>원본 행 순서 유지 · 순번으로 연결</p></footer>
      </main>
    </div>

    {reviewing && <div className="modal-backdrop" onClick={event => { if (event.target === event.currentTarget && !busy) setReviewing(false); }}><section className="export-modal" role="dialog" aria-modal="true" aria-labelledby="export-title"><div className="modal-heading"><div><span className="eyebrow">REVIEW & EXPORT</span><h2 id="export-title">변경 내용을 확인하세요.</h2></div><button className="icon-button" aria-label="내보내기 검토 닫기" onClick={() => setReviewing(false)} disabled={Boolean(busy)}><X size={21}/></button></div><div className="export-summary"><span><FileSpreadsheet size={20}/>{document?.name}</span><strong>{edits.size}개 셀 변경</strong></div><p className="export-explanation">원본의 행 순서와 다른 열을 유지합니다. 다운로드 전에 파일을 다시 읽어 모든 셀을 원본 및 수정 내용과 대조합니다.</p><div className="change-list">{edits.size === 0 ? <div className="no-changes"><ShieldCheck size={22}/><p>변경된 셀이 없습니다.<br/>원본과 같은 내용을 다운로드합니다.</p></div> : [...edits.values()].sort((a, b) => a.rowIndex - b.rowIndex || a.columnIndex - b.columnIndex).map(edit => <div className="change-item" key={cellKey(edit.rowIndex, edit.columnIndex)}><div><strong>원본 {edit.rowIndex + 1}행</strong><span>{sheet?.rows[mapping.headerRow]?.[edit.columnIndex] ?? `열 ${edit.columnIndex + 1}`}</span><span>순번 {sheet?.rows[edit.rowIndex]?.[mapping.idColumn]}</span></div><div className="change-values"><p><span>원본</span>{sheet?.rows[edit.rowIndex]?.[edit.columnIndex] || <em>(빈 셀)</em>}</p><ArrowDownToLine size={12}/><p><span>수정</span>{edit.value || <em>(빈 셀)</em>}</p></div></div>)}</div><p className="format-note">가져온 파일과 같은 형식으로 저장합니다. Webflow의 CSV 업로드를 사용한다면 해당 CSV 양식을 불러와 작업해 주세요.</p><div className="modal-actions">{edits.size > 0 && <button className="button subtle" onClick={resetEdits} disabled={Boolean(busy)}><RotateCcw size={15}/>전체 수정 취소</button>}<button className="button primary" onClick={downloadWork} disabled={!canExport || Boolean(busy)}>{busy ? <LoaderCircle className="spin" size={15}/> : <ArrowDownToLine size={15}/>}검증 후 다운로드</button></div></section></div>}
  </div>;
}
