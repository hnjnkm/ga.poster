import { test, expect, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { strToU8, strFromU8, zipSync, unzipSync } from 'fflate';

const pixel = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a2ioAAAAASUVORK5CYII=', 'base64');
const originalCsv = '\uFEFF순번,한국어,English,简体中文,Slug,Item ID\r\n02,"둘째 줄\r\n원문",Second,第二,keep-b,webflow-b\r\n1002,천이,One thousand,一千零二,keep-c,webflow-c\r\n01,첫 이미지,First,第一,keep-a,webflow-a\r\n04,이미지 없음,Unmatched,没有图片,keep-d,webflow-d\r\n';
async function uploadImages(page: Page, names = ['1002.png', '01.png', '02.png']) {
  await page.getByLabel('이미지 업로드', { exact: true }).setInputFiles(names.map(name => ({ name, mimeType: 'image/png', buffer: pixel })));
}
async function uploadCsv(page: Page, text = originalCsv) {
  await page.getByLabel('설명 파일 업로드', { exact: true }).setInputFiles({ name: 'archive.csv', mimeType: 'text/csv', buffer: Buffer.from(text) });
  await page.getByRole('button', { name: '이 열로 연결하기', exact: true }).click();
}
async function exportFile(page: Page) {
  await page.getByRole('button', { name: /^내보내기/ }).click();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: '검증 후 다운로드', exact: true }).click();
  const result = await download;
  return { name: result.suggestedFilename(), bytes: await readFile((await result.path())!) };
}

test('imports, matches by ID, edits three languages and preserves original order and unrelated data', async ({ page }) => {
  const runtimeErrors: string[] = [];
  page.on('pageerror', error => runtimeErrors.push(error.message));
  await page.goto('./');
  await expect(page.getByRole('heading', { name: '첫 이미지를 불러와 보세요.' })).toBeVisible();
  await page.screenshot({ path: '/tmp/graphic-archive-empty.png' });
  await uploadImages(page);
  await uploadCsv(page);
  await expect(page.getByLabel('01.png 한국어 설명', { exact: true })).toHaveValue('첫 이미지');
  await expect(page.getByLabel('02.png 한국어 설명', { exact: true })).toHaveValue('둘째 줄\n원문');
  expect(await page.locator('.archive-card .image-number').allTextContents()).toEqual(['01', '02', '1002']);
  await page.getByLabel('01.png 한국어 설명', { exact: true }).fill('수정한 설명');
  await page.getByLabel('01.png English 설명', { exact: true }).fill('Updated, "English"\nwith a new line');
  await page.getByLabel('01.png 简体中文 설명', { exact: true }).fill('更新的简体中文');
  await page.screenshot({ path: '/tmp/graphic-archive-editor.png' });
  const output = await exportFile(page);
  expect(output.name).toBe('archive-edited.csv');
  expect(output.bytes.toString('utf8')).toBe(originalCsv.replace('01,첫 이미지,First,第一,', '01,수정한 설명,"Updated, ""English""\nwith a new line",更新的简体中文,'));
  expect(runtimeErrors).toEqual([]);
});

test('blocks canonical duplicate image IDs and unlocks after deliberate removal', async ({ page }) => {
  await page.goto('./');
  await uploadImages(page, ['01.png', '1.png']);
  await uploadCsv(page, '\uFEFF순번,한국어,English,简体中文\r\n01,원문,Original,原文\r\n');
  await expect(page.getByRole('button', { name: /^내보내기/ })).toBeDisabled();
  await expect(page.getByText('중복된 이미지 순번: 1', { exact: true })).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await page.getByLabel('1.png 이미지 제거', { exact: true }).click();
  await expect(page.getByRole('button', { name: /^내보내기/ })).toBeEnabled();
  expect((await exportFile(page)).bytes.toString('utf8')).toBe('\uFEFF순번,한국어,English,简体中文\r\n01,원문,Original,原文\r\n');
});

test('saves and restores original files, mapping and unsaved-to-workbook edits after reload', async ({ page }) => {
  await page.goto('./');
  await uploadImages(page, ['01.png']);
  await uploadCsv(page);
  await page.getByLabel('01.png 한국어 설명', { exact: true }).fill('저장된 수정');
  await page.getByRole('button', { name: '작업 저장', exact: true }).click();
  await expect(page.getByText('브라우저에 저장됨', { exact: true })).toBeVisible();
  await page.reload();
  await expect(page.getByRole('button', { name: /저장한 작업 복원/ })).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await page.getByRole('button', { name: /저장한 작업 복원/ }).click();
  await expect(page.getByLabel('01.png 한국어 설명', { exact: true })).toHaveValue('저장된 수정');
  await expect(page.locator('.archive-card img')).toHaveJSProperty('naturalWidth', 1);
  expect((await exportFile(page)).bytes.toString('utf8')).toBe(originalCsv.replace('01,첫 이미지,', '01,저장된 수정,'));
});

test('XLSX edits preserve other ZIP entries and cell content', async ({ page }) => {
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const entry = (ref: string, value: string) => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
  const sheet = `<worksheet xmlns="${ns}"><sheetData><row r="1">${['순번','한국어','English','简体中文','Slug'].map((value,index) => entry(`${String.fromCharCode(65+index)}1`,value)).join('')}</row><row r="2">${['0001','원본','Original','原文','fixed-slug'].map((value,index) => entry(`${String.fromCharCode(65+index)}2`,value)).join('')}</row></sheetData></worksheet>`;
  const original = zipSync({
    'xl/workbook.xml': strToU8(`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
    'xl/worksheets/sheet1.xml': strToU8(sheet),
    'xl/styles.xml': strToU8('<original-styles/>'),
    'custom.bin': new Uint8Array([17,32,255,0]),
  });
  await page.goto('./');
  await uploadImages(page, ['01.png']);
  await page.getByLabel('설명 파일 업로드', { exact: true }).setInputFiles({ name: 'archive.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(original) });
  await page.getByRole('button', { name: '이 열로 연결하기', exact: true }).click();
  await expect(page.getByLabel('01.png 한국어 설명', { exact: true })).toHaveValue('원본');
  await page.getByLabel('01.png 한국어 설명', { exact: true }).fill('한글 & English <中文>');
  const output = await exportFile(page);
  const before = unzipSync(original);
  const after = unzipSync(new Uint8Array(output.bytes));
  for (const key of Object.keys(before).filter(key => !key.endsWith('sheet1.xml'))) expect(after[key]).toEqual(before[key]);
  const xml = strFromU8(after['xl/worksheets/sheet1.xml']);
  expect(xml).toContain('한글 &amp; English &lt;中文&gt;');
  expect(xml).toContain('fixed-slug');
  expect(xml).toContain('0001');
});

test('supports 1000 records without changing original descending row order', async ({ page }) => {
  const rows = Array.from({ length: 1000 }, (_, index) => 1000 - index).map(id => `${id},원문${id},English${id},中文${id},slug-${id}`);
  const csv = '순번,한국어,English,简体中文,Slug\n' + rows.join('\n');
  await page.goto('./');
  await uploadImages(page, Array.from({ length: 1000 }, (_, index) => `${index + 1}.png`));
  await uploadCsv(page, csv);
  await expect(page.locator('.archive-card')).toHaveCount(1000);
  await page.getByLabel('파일명 검색', { exact: true }).fill('1000.');
  await expect(page.locator('.archive-card')).toHaveCount(1);
  await page.getByLabel('1000.png 한국어 설명', { exact: true }).fill('수정1000');
  const output = await exportFile(page);
  expect(output.bytes.toString('utf8')).toBe(csv.replace('1000,원문1000,', '1000,수정1000,'));
});

test('mobile layout fits the viewport', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('./');
  await uploadImages(page, ['01.png']);
  await uploadCsv(page);
  await expect(page.getByLabel('01.png 한국어 설명', { exact: true })).toHaveValue('첫 이미지');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
  await page.screenshot({ path: '/tmp/graphic-archive-mobile.png', fullPage: true });
});
