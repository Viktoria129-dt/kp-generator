"use client";

import { ChangeEvent, useMemo, useRef, useState } from "react";
import PizZip from "pizzip";
import * as XLSX from "xlsx";

type Cell = string | number;
const sample: Cell[][] = [["№", "Номенклатура", "Кол-во", "Ед.", "Цена", "Сумма"], [1, "Наименование товара или услуги", 1, "шт", "0,00", "0,00"]];
const xmlEscape = (value: Cell) => String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
const parsePastedTable = (value: string): Cell[][] => value.trim().split(/\r?\n/).filter(Boolean).map((row) => row.split("\t").map((cell) => cell.trim()));

function parseMoney(value: Cell): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const normalized = String(value).replace(/[\s\u00a0₽]/g, "").replace(/,(?=\d{1,2}$)/, ".").replace(/[^\d.-]/g, "");
  const result = Number(normalized);
  return Number.isFinite(result) ? result : 0;
}

function plural(value: number, forms: [string, string, string]) {
  const n = Math.abs(value) % 100;
  const n1 = n % 10;
  if (n > 10 && n < 20) return forms[2];
  if (n1 === 1) return forms[0];
  if (n1 >= 2 && n1 <= 4) return forms[1];
  return forms[2];
}

function integerToWords(value: number) {
  if (value === 0) return "ноль";
  const ones = [["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"], ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]];
  const teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];
  const groups = [null, ["тысяча", "тысячи", "тысяч"], ["миллион", "миллиона", "миллионов"], ["миллиард", "миллиарда", "миллиардов"]] as const;
  const result: string[] = [];
  let n = Math.floor(value), group = 0;
  while (n > 0) {
    const part = n % 1000;
    if (part) {
      const words: string[] = [];
      words.push(hundreds[Math.floor(part / 100)]);
      const lastTwo = part % 100;
      if (lastTwo >= 10 && lastTwo < 20) words.push(teens[lastTwo - 10]);
      else { words.push(tens[Math.floor(lastTwo / 10)]); words.push(ones[group === 1 ? 1 : 0][lastTwo % 10]); }
      if (group > 0 && groups[group]) words.push(plural(part, [...groups[group]] as [string, string, string]));
      result.unshift(...words.filter(Boolean));
    }
    n = Math.floor(n / 1000); group++;
  }
  return result.join(" ");
}

function formatMoney(value: number) { return value.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).replaceAll(" ", " "); }
function moneyInWords(value: number) { const rubles = Math.floor(value); const kopecks = Math.round((value - rubles) * 100); const words = integerToWords(rubles); return `${words.charAt(0).toUpperCase()}${words.slice(1)} ${plural(rubles, ["рубль", "рубля", "рублей"])} ${String(kopecks).padStart(2, "0")} ${plural(kopecks, ["копейка", "копейки", "копеек"])}`; }

function tableXml(rows: Cell[][]) {
  const columns = Math.max(...rows.map((row) => row.length), 1);
  const totalWidth = 9360;
  const widths = columns === 6 ? [540, 4950, 750, 520, 1240, 1360] : Array.from({ length: columns }, () => Math.floor(totalWidth / columns));
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join("");
  const body = rows.map((row, rowIndex) => {
    const cells = Array.from({ length: columns }, (_, index) => {
      const alignment = rowIndex > 0 && index === 1 ? "left" : "center";
      return `<w:tc><w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="${alignment}"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">${xmlEscape(row[index] ?? "")}</w:t></w:r></w:p></w:tc>`;
    }).join("");
    return `<w:tr><w:trPr><w:trHeight w:val="${rowIndex === 0 ? 620 : 520}" w:hRule="atLeast"/></w:trPr>${cells}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function paragraph(text: string, options: { bold?: boolean; center?: boolean; size?: number; after?: number } = {}) {
  const { bold = false, center = false, size = 22, after = 120 } = options;
  return `<w:p><w:pPr><w:spacing w:after="${after}"/>${center ? '<w:jc w:val="center"/>' : ""}</w:pPr><w:r><w:rPr>${bold ? "<w:b/>" : ""}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}

function alignedLine(left: string, right: string, options: { bold?: boolean; before?: number; after?: number } = {}) {
  const { bold = false, before = 0, after = 120 } = options;
  const style = `${bold ? "<w:b/>" : ""}<w:sz w:val="22"/><w:szCs w:val="22"/>`;
  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:before="${before}" w:after="${after}"/></w:pPr><w:r><w:rPr>${style}</w:rPr><w:t xml:space="preserve">${xmlEscape(left)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr>${style}</w:rPr><w:t xml:space="preserve">${xmlEscape(right)}</w:t></w:r></w:p>`;
}

function signatureWithCenteredStamp(left: string, right: string) {
  const raisedText = '<w:b/><w:position w:val="90"/><w:sz w:val="22"/><w:szCs w:val="22"/>';
  const drawing = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2221200" cy="2087300"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="999" name="Печать"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="stamp.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdStamp"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2221200" cy="2087300"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  return `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="3300"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:before="2400" w:after="0"/></w:pPr><w:r><w:rPr>${raisedText}</w:rPr><w:t xml:space="preserve">${xmlEscape(left)}</w:t></w:r><w:r><w:tab/></w:r><w:r>${drawing}</w:r><w:r><w:tab/></w:r><w:r><w:rPr>${raisedText}</w:rPr><w:t xml:space="preserve">${xmlEscape(right)}</w:t></w:r></w:p>`;
}

export default function Home() {
  const [number, setNumber] = useState("125-1/1");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [city, setCity] = useState("г. Москва");
  const [title, setTitle] = useState("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ");
  const [includeObject, setIncludeObject] = useState(false);
  const [objectName, setObjectName] = useState("");
  const [delivery, setDelivery] = useState("Доставка до города Омск входит в стоимость светильников и осуществляется Поставщиком.");
  const [directorTitle, setDirectorTitle] = useState("Генеральный директор");
  const directorName = "Матыцин А. Ю.";
  const [rows, setRows] = useState<Cell[][]>(sample);
  const [template, setTemplate] = useState<File | null>(null);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteValue, setPasteValue] = useState("");
  const [status, setStatus] = useState("");
  const tableInput = useRef<HTMLInputElement>(null);
  const templateInput = useRef<HTMLInputElement>(null);
  const formattedDate = useMemo(() => { if (!date) return ""; const [year, month, day] = date.split("-"); return `${day}.${month}.${year}`; }, [date]);
  const calculation = useMemo(() => {
    const header = rows[0] ?? [];
    const detected = header.findIndex((cell) => String(cell).trim().toLowerCase().includes("сумм"));
    const sumColumn = detected >= 0 ? detected : Math.max(header.length - 1, 0);
    const totalValue = Math.round(rows.slice(1).reduce((sum, row) => sum + parseMoney(row[sumColumn] ?? 0), 0) * 100) / 100;
    const vat = Math.round((totalValue * 22 / 122) * 100) / 100;
    const totalText = `Итого: ${formatMoney(totalValue)} рублей (${moneyInWords(totalValue)}), в том числе НДС 22% - ${formatMoney(vat)} рублей.`;
    return { totalValue, vat, totalText };
  }, [rows]);

  async function importSpreadsheet(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]; if (!file) return;
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json<Cell[]>(sheet, { header: 1, defval: "" });
    const cleaned = data.filter((row) => row.some((cell) => String(cell).trim()));
    if (cleaned.length) setRows(cleaned);
    setStatus(`Таблица загружена: ${file.name}`);
  }

  function applyPastedTable() { const parsed = parsePastedTable(pasteValue); if (!parsed.length) return; setRows(parsed); setPasteOpen(false); setStatus(`Вставлено строк: ${parsed.length}`); }
  function updateCell(rowIndex: number, columnIndex: number, value: string) { setRows((current) => current.map((row, ri) => ri === rowIndex ? Array.from({length: Math.max(row.length, columnIndex + 1)}, (_, ci) => ci === columnIndex ? value : row[ci] ?? "") : row)); }

  async function generateDocument() {
    try {
      setStatus("Формируем документ…");
      const templateBuffer = template ? await template.arrayBuffer() : await fetch("/kp-template.docx").then((response) => { if (!response.ok) throw new Error("Не удалось загрузить встроенный бланк"); return response.arrayBuffer(); });
      const zip = new PizZip(templateBuffer);
      const stampBuffer = await fetch("/stamp.png").then((response) => { if (!response.ok) throw new Error("Не удалось загрузить печать"); return response.arrayBuffer(); });
      zip.file("word/media/stamp.png", stampBuffer);
      const relsFile = zip.file("word/_rels/document.xml.rels");
      if (!relsFile) throw new Error("В бланке не найдены связи документа");
      const rels = relsFile.asText();
      if (!rels.includes('Id="rIdStamp"')) zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdStamp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stamp.png"/></Relationships>'));
      const typesFile = zip.file("[Content_Types].xml");
      if (typesFile) { const types = typesFile.asText(); if (!types.includes('Extension="png"')) zip.file("[Content_Types].xml", types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')); }
      const documentFile = zip.file("word/document.xml"); if (!documentFile) throw new Error("В бланке не найден документ Word");
      let original = documentFile.asText();
      if (!original.includes("xmlns:a=")) original = original.replace("<w:document ", '<w:document xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ');
      const section = original.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0]; if (!section) throw new Error("В бланке отсутствуют параметры страницы");
      const content = [alignedLine(`Исх. №${number} от ${formattedDate} года`, city, { after: 260 }), paragraph(title, { bold: true, center: true, size: 22, after: includeObject && objectName.trim() ? 140 : 260 }), ...(includeObject && objectName.trim() ? [paragraph(`Объект: ${objectName.trim()}`, { after: 220 })] : []), tableXml(rows), paragraph("", { after: 180 }), paragraph(calculation.totalText, { bold: true, after: 320 }), paragraph(delivery, { after: 0 }), signatureWithCenteredStamp(directorTitle, directorName)].join("");
      zip.file("word/document.xml", original.replace(/<w:body>[\s\S]*?<\/w:body>/, `<w:body>${content}${section}</w:body>`));
      const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `КП_исх_${number.replaceAll("/", "-")}_${formattedDate}.docx`; link.click(); URL.revokeObjectURL(url);
      setStatus("Готово — документ скачан");
    } catch (error) { setStatus(error instanceof Error ? error.message : "Не удалось сформировать документ"); }
  }

  const columns = Math.max(...rows.map((row) => row.length), 1);
  return <main>
    <header className="topbar"><div className="brandMark">КП</div><div><strong>Генератор коммерческих предложений</strong><span>ООО «МосЭнергоСервис»</span></div><div className="privacy">Данные остаются на компьютере</div></header>
    <section className="workspace">
      <div className="intro"><div><p className="eyebrow">Новый документ</p><h1>Соберите КП за пару минут</h1><p>Заполните реквизиты, добавьте таблицу и скачайте готовый Word с фирменной шапкой и подвалом.</p></div><button className="primary" onClick={generateDocument}>Сформировать Word</button></div>
      <div className="grid">
        <aside className="panel controls">
          <div className="step"><span>1</span><div><h2>Бланк</h2><p>По умолчанию используется фирменный бланк из примера.</p></div></div>
          <input ref={templateInput} hidden type="file" accept=".docx" onChange={(e) => { setTemplate(e.target.files?.[0] ?? null); setStatus(e.target.files?.[0] ? `Бланк выбран: ${e.target.files[0].name}` : ""); }} />
          <button className="secondary" onClick={() => templateInput.current?.click()}>{template ? "Заменить выбранный бланк" : "Загрузить другой бланк"}</button>{template && <button className="textButton" onClick={() => setTemplate(null)}>Вернуть фирменный бланк</button>}
          <div className="divider" /><div className="step"><span>2</span><div><h2>Реквизиты</h2><p>Они попадут в начало документа.</p></div></div>
          <label>Исходящий номер<input value={number} onChange={(e) => setNumber(e.target.value)} /></label><div className="twoCols"><label>Дата<input type="date" value={date} onChange={(e) => setDate(e.target.value)} /></label><label>Город<input value={city} onChange={(e) => setCity(e.target.value)} /></label></div><label>Заголовок<input value={title} onChange={(e) => setTitle(e.target.value)} /></label>
          <label className="checkRow"><input type="checkbox" checked={includeObject} onChange={(e) => setIncludeObject(e.target.checked)} /><span>Добавить название объекта</span></label>
          {includeObject && <label>Название объекта<input value={objectName} onChange={(e) => setObjectName(e.target.value)} placeholder="Например: БЦ «Омега», корпус 2" /></label>}
          <div className="divider" /><div className="step"><span>3</span><div><h2>Таблица</h2><p>Excel, CSV или копирование прямо из таблицы.</p></div></div>
          <input ref={tableInput} hidden type="file" accept=".xlsx,.xls,.csv" onChange={importSpreadsheet} /><div className="buttonRow"><button className="secondary" onClick={() => tableInput.current?.click()}>Загрузить Excel / CSV</button><button className="secondary" onClick={() => setPasteOpen((v) => !v)}>Вставить из буфера</button></div>
          {pasteOpen && <div className="pasteBox"><textarea autoFocus placeholder="Скопируйте ячейки в Excel и вставьте сюда" value={pasteValue} onChange={(e) => setPasteValue(e.target.value)} /><button className="smallPrimary" onClick={applyPastedTable}>Применить</button></div>}
          <div className="divider" /><div className="step"><span>4</span><div><h2>Итог и подпись</h2><p>Сумма и НДС рассчитываются из столбца «Сумма».</p></div></div>
          <div className="calculation"><span>Итого автоматически</span><strong>{formatMoney(calculation.totalValue)} ₽</strong><small>в том числе НДС 22%: {formatMoney(calculation.vat)} ₽</small></div>
          <label>Условия доставки<textarea value={delivery} onChange={(e) => setDelivery(e.target.value)} /></label>
          <div className="twoCols"><label>Должность<input value={directorTitle} onChange={(e) => setDirectorTitle(e.target.value)} /></label><label>ФИО директора<input value={directorName} readOnly aria-readonly="true" /></label></div>
        </aside>
        <section className="panel previewPanel"><div className="previewHeader"><div><h2>Предпросмотр таблицы</h2><p>{rows.length} строк · {columns} столбцов</p></div><span className="editable">Можно редактировать</span></div><div className="tableWrap"><table><tbody>{rows.map((row, ri) => <tr key={ri}>{Array.from({ length: columns }, (_, ci) => <td key={ci} className={ri === 0 ? "headingCell" : ""}><input aria-label={`Строка ${ri + 1}, столбец ${ci + 1}`} value={String(row[ci] ?? "")} onChange={(e) => updateCell(ri, ci, e.target.value)} /></td>)}</tr>)}</tbody></table></div><div className="tableActions"><button className="textButton" onClick={() => setRows((r) => [...r, Array(columns).fill("")])}>+ Добавить строку</button><button className="textButton danger" disabled={rows.length <= 1} onClick={() => setRows((r) => r.slice(0, -1))}>Удалить последнюю</button></div></section>
      </div>
      <footer className="actionbar"><span className={status.includes("Не удалось") ? "error" : ""}>{status || "Документ ещё не сформирован"}</span><button className="primary" onClick={generateDocument}>Сформировать и скачать .docx</button></footer>
    </section>
  </main>;
}
