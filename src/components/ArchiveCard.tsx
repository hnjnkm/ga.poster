import { memo, useEffect, useState } from 'react';
import { Languages, Link2, RotateCcw, Trash2 } from 'lucide-react';
import { findCellIssue, type CellEdit, type SpreadsheetDocument, type SpreadsheetSheet } from '../lib/spreadsheet';
import type { ColumnMapping, ImageIdentifier, MatchedRecord } from '../lib/matching';

const languages = [
  { key: 'koColumn', label: '한국어', code: 'KO', lang: 'ko', editKey: 'koEdit' },
  { key: 'enColumn', label: 'English', code: 'EN', lang: 'en', editKey: 'enEdit' },
  { key: 'zhColumn', label: '简体中文', code: 'ZH', lang: 'zh-Hans', editKey: 'zhEdit' },
] as const;

function ImagePreview({ file }: { file: File }) {
  const [url, setUrl] = useState('');
  useEffect(() => {
    const next = URL.createObjectURL(file);
    setUrl(next);
    return () => URL.revokeObjectURL(next);
  }, [file]);
  return <img src={url || undefined} alt={file.name} loading="lazy" decoding="async" />;
}

interface Props {
  file: File;
  imageIndex: number;
  parsed: ImageIdentifier | null;
  match: MatchedRecord | undefined;
  document: SpreadsheetDocument | null;
  sheet: SpreadsheetSheet | undefined;
  sheetId: string;
  mapping: ColumnMapping;
  koEdit?: CellEdit;
  enEdit?: CellEdit;
  zhEdit?: CellEdit;
  onEdit: (row: number, column: number, value: string) => void;
  onRemove: (index: number, name: string) => void;
}

export default memo(function ArchiveCard(props: Props) {
  const { file, imageIndex, parsed, match, document, sheet, sheetId, mapping, onEdit, onRemove } = props;
  const rowIndex = match?.rowIndex;
  const edited = props.koEdit || props.enEdit || props.zhEdit;
  return <article className="archive-card">
    <div className="image-column">
      <div className="image-meta"><span className="image-number">{parsed?.originalStem ?? '—'}</span><span className={`connection-tag ${match ? 'connected' : ''}`}>{match ? <><Link2 size={11}/>연결됨</> : '미연결'}</span></div>
      <div className="image-frame"><ImagePreview file={file}/></div>
      <div className="image-file"><span title={file.name}>{file.name}</span><small>{(file.size / 1024 / 1024).toFixed(2)} MB</small><button className="icon-button" aria-label={`${file.name} 이미지 제거`} onClick={() => onRemove(imageIndex, file.name)}><Trash2 size={13}/></button></div>
    </div>
    <div className="description-column">
      <div className="description-heading"><span><Languages size={15}/> 이미지 설명</span>{rowIndex !== undefined ? <small>원본 {rowIndex + 1}행 {edited && <b>· 수정됨</b>}</small> : <small>순번 연결 필요</small>}</div>
      {languages.map(language => {
        const column = mapping[language.key];
        const edit = props[language.editKey];
        const value = rowIndex !== undefined ? edit?.value ?? sheet?.rows[rowIndex]?.[column] ?? '' : '';
        const issue = document && rowIndex !== undefined ? findCellIssue(document, sheetId, rowIndex, column) : null;
        return <div className={`language-field ${edit ? 'is-edited' : ''}`} key={language.key}>
          <div className="language-label"><label htmlFor={`desc-${imageIndex}-${language.code}`}><span>{language.code}</span>{language.label}</label><span>{edit && <button className="inline-reset" aria-label={`${file.name} ${language.label} 원본 복원`} onClick={() => onEdit(rowIndex!, column, sheet!.rows[rowIndex!]?.[column] ?? '')}><RotateCcw size={11}/>원본</button>}{value.length.toLocaleString()} 자</span></div>
          <textarea id={`desc-${imageIndex}-${language.code}`} aria-label={`${file.name} ${language.label} 설명`} lang={language.lang} spellCheck={false} value={value} disabled={rowIndex === undefined || Boolean(issue)} placeholder={issue ?? (rowIndex !== undefined ? '설명을 입력하세요.' : document ? '열 연결과 순번을 확인해 주세요.' : '기존 설명 파일을 먼저 불러와 주세요.')} onChange={event => onEdit(rowIndex!, column, event.target.value)}/>
          {issue && <small className="cell-warning">{issue}</small>}
        </div>;
      })}
    </div>
  </article>;
});
