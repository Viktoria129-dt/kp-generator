"use client";

import { ChangeEvent, useMemo, useRef, useState } from "react";
import PizZip from "pizzip";
import * as XLSX from "xlsx";

type Cell = string | number;
const sample: Cell[][] = [["№", "Номенклатура", "Кол-во", "Ед.", "Цена", "Сумма"], [1, "Наименование товара или услуги", 1, "шт", "0,00", "0,00"]];
const verdanaFont = '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" w:eastAsia="Verdana" w:cs="Verdana"/>';
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
  const groups: Array<[string, string, string] | null> = [null, ["тысяча", "тысячи", "тысяч"], ["миллион", "миллиона", "миллионов"], ["миллиард", "миллиарда", "миллиардов"]];
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
      const groupForms = groups[group];
      if (group > 0 && groupForms) words.push(plural(part, groupForms));
      result.unshift(...words.filter(Boolean));
    }
    n = Math.floor(n / 1000); group++;
  }
  return result.join(" ");
}

function formatMoney(value: number) { return value.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).replaceAll(" ", " "); }
function moneyInWords(value: number) { const rubles = Math.floor(value); const kopecks = Math.round((value - rubles) * 100); const words = integerToWords(rubles); return `${words.charAt(0).toUpperCase()}${words.slice(1)} ${plural(rubles, ["рубль", "рубля", "рублей"])} ${String(kopecks).padStart(2, "0")} ${plural(kopecks, ["копейка", "копейки", "копеек"])}`; }

async function imageToPng(file: File): Promise<ArrayBuffer> {
  const image = new Image();
  const url = URL.createObjectURL(file);
  try {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Не удалось прочитать изображение печати"));
      image.src = url;
    });
    const source = document.createElement("canvas");
    source.width = image.naturalWidth;
    source.height = image.naturalHeight;
    const sourceContext = source.getContext("2d", { willReadFrequently: true });
    if (!sourceContext) throw new Error("Не удалось подготовить изображение печати");
    sourceContext.drawImage(image, 0, 0);
    const pixels = sourceContext.getImageData(0, 0, source.width, source.height).data;
    let left = source.width, top = source.height, right = -1, bottom = -1;
    for (let y = 0; y < source.height; y++) {
      for (let x = 0; x < source.width; x++) {
        const index = (y * source.width + x) * 4;
        const visible = pixels[index + 3] > 12 && (pixels[index] < 248 || pixels[index + 1] < 248 || pixels[index + 2] < 248);
        if (!visible) continue;
        left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x); bottom = Math.max(bottom, y);
      }
    }
    if (right < left || bottom < top) { left = 0; top = 0; right = source.width - 1; bottom = source.height - 1; }
    const contentWidth = right - left + 1;
    const contentHeight = bottom - top + 1;
    const padding = Math.ceil(Math.max(contentWidth, contentHeight) * 0.04);
    const side = Math.max(contentWidth, contentHeight) + padding * 2;
    const canvas = document.createElement("canvas");
    canvas.width = side;
    canvas.height = side;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Не удалось подготовить изображение печати");
    context.drawImage(source, left, top, contentWidth, contentHeight, (side - contentWidth) / 2, (side - contentHeight) / 2, contentWidth, contentHeight);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!blob) throw new Error("Не удалось преобразовать печать в PNG");
    return blob.arrayBuffer();
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function rasterToGrayscale(buffer: ArrayBuffer | Uint8Array, extension: string): Promise<ArrayBuffer> {
  const normalized = extension.toLowerCase();
  const mimeType = normalized === "jpg" || normalized === "jpeg" ? "image/jpeg" : normalized === "png" ? "image/png" : "";
  if (!mimeType) return buffer instanceof Uint8Array ? buffer.slice().buffer : buffer;

  const image = new Image();
  const bytes = buffer instanceof Uint8Array ? new Uint8Array(buffer) : new Uint8Array(buffer);
  const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));
  try {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Не удалось подготовить чёрно-белую версию бланка"));
      image.src = url;
    });
    const canvas = document.createElement("canvas");
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Не удалось подготовить чёрно-белую версию бланка");
    context.drawImage(image, 0, 0);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
    for (let index = 0; index < pixels.data.length; index += 4) {
      const gray = Math.round(pixels.data[index] * 0.299 + pixels.data[index + 1] * 0.587 + pixels.data[index + 2] * 0.114);
      pixels.data[index] = gray;
      pixels.data[index + 1] = gray;
      pixels.data[index + 2] = gray;
    }
    context.putImageData(pixels, 0, 0);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, mimeType, 0.94));
    if (!blob) throw new Error("Не удалось создать чёрно-белую версию бланка");
    return blob.arrayBuffer();
  } finally {
    URL.revokeObjectURL(url);
  }
}

function normalizeWordFormatting(zip: PizZip) {
  Object.keys(zip.files).filter((name) => /^word\/.*\.xml$/i.test(name)).forEach((name) => {
    const file = zip.file(name);
    if (!file) return;
    let xml = file.asText()
      .replace(/<w:rFonts\b[^>]*\/>/g, verdanaFont)
      .replace(/<w:sz w:val="[^"]*"\/>/g, '<w:sz w:val="20"/>')
      .replace(/<w:szCs w:val="[^"]*"\/>/g, '<w:szCs w:val="20"/>')
      .replace(/(<a:(?:defRPr|rPr|endParaRPr)\b[^>]*\bsz=")[^"]*(")/g, (_, before, after) => `${before}1000${after}`)
      .replace(/(<a:(?:latin|ea|cs)\b[^>]*\btypeface=")[^"]*(")/g, "$1Verdana$2");
    if (name === "word/styles.xml" && !xml.includes('w:ascii="Verdana"')) xml = xml.replace("<w:rPrDefault><w:rPr>", `<w:rPrDefault><w:rPr>${verdanaFont}<w:sz w:val="20"/><w:szCs w:val="20"/>`);
    zip.file(name, xml);
  });
}

function grayscaleHex(hex: string) {
  const value = hex.replace("#", "");
  const red = parseInt(value.slice(0, 2), 16);
  const green = parseInt(value.slice(2, 4), 16);
  const blue = parseInt(value.slice(4, 6), 16);
  const gray = Math.round(red * 0.299 + green * 0.587 + blue * 0.114).toString(16).padStart(2, "0").toUpperCase();
  return `${gray}${gray}${gray}`;
}

function grayscaleWordColors(zip: PizZip) {
  Object.keys(zip.files).filter((name) => /^word\/.*\.xml$/i.test(name)).forEach((name) => {
    const file = zip.file(name);
    if (!file) return;
    const xml = file.asText()
      .replace(/(<a:(?:srgbClr|sysClr)\b[^>]*\b(?:val|lastClr)=")([0-9a-f]{6})(")/gi, (_, before, color, after) => `${before}${grayscaleHex(color)}${after}`)
      .replace(/((?:w:color|w:fill|color|fill|strokecolor|fillcolor)=")#?([0-9a-f]{6})(")/gi, (_, before, color, after) => `${before}${grayscaleHex(color)}${after}`)
      .replace(/<w:color\b[^>]*\/>/gi, '<w:color w:val="000000"/>');
    zip.file(name, xml);
  });
}

function tableXml(rows: Cell[][]) {
  const columns = Math.max(...rows.map((row) => row.length), 1);
  const totalWidth = 9360;
  const widths = columns === 6 ? [540, 4950, 750, 520, 1240, 1360] : Array.from({ length: columns }, () => Math.floor(totalWidth / columns));
  const grid = widths.map((width) => `<w:gridCol w:w="${width}"/>`).join("");
  const body = rows.map((row, rowIndex) => {
    const cells = Array.from({ length: columns }, (_, index) => {
      const alignment = rowIndex > 0 && index === 1 ? "left" : "center";
      return `<w:tc><w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="${alignment}"/></w:pPr><w:r><w:rPr>${verdanaFont}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">${xmlEscape(row[index] ?? "")}</w:t></w:r></w:p></w:tc>`;
    }).join("");
    return `<w:tr><w:trPr><w:trHeight w:val="${rowIndex === 0 ? 620 : 520}" w:hRule="atLeast"/></w:trPr>${cells}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function paragraph(text: string, options: { bold?: boolean; center?: boolean; size?: number; after?: number } = {}) {
  const { bold = false, center = false, size = 20, after = 120 } = options;
  return `<w:p><w:pPr><w:spacing w:after="${after}"/>${center ? '<w:jc w:val="center"/>' : ""}</w:pPr><w:r><w:rPr>${verdanaFont}${bold ? "<w:b/>" : ""}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}

function alignedLine(left: string, right: string, options: { bold?: boolean; before?: number; after?: number } = {}) {
  const { bold = false, before = 0, after = 120 } = options;
  const style = `${verdanaFont}${bold ? "<w:b/>" : ""}<w:sz w:val="20"/><w:szCs w:val="20"/>`;
  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:before="${before}" w:after="${after}"/></w:pPr><w:r><w:rPr>${style}</w:rPr><w:t xml:space="preserve">${xmlEscape(left)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr>${style}</w:rPr><w:t xml:space="preserve">${xmlEscape(right)}</w:t></w:r></w:p>`;
}

function signatureWithCenteredStamp(left: string, right: string) {
  const textStyle = `${verdanaFont}<w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>`;
  const drawing = `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>520000</wp:posOffset></wp:positionV><wp:extent cx="1800000" cy="1800000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="999" name="Печать 5 см"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="stamp.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdStamp"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="1800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  return `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:before="1800" w:after="1800"/><w:keepLines/></w:pPr><w:r><w:rPr>${textStyle}</w:rPr><w:t xml:space="preserve">${xmlEscape(left)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr>${textStyle}</w:rPr><w:t xml:space="preserve">${xmlEscape(right)}</w:t></w:r><w:r>${drawing}</w:r></w:p>`;
}

export default function Home() {
  const [number, setNumber] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [city, setCity] = useState("");
  const [title, setTitle] = useState("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ");
  const [includeObject, setIncludeObject] = useState(false);
  const [objectName, setObjectName] = useState("");
  const [delivery, setDelivery] = useState("");
  const [directorTitle, setDirectorTitle] = useState("Генеральный директор");
  const [directorName, setDirectorName] = useState("");
  const [stamp, setStamp] = useState<File | null>(null);
  const [rows, setRows] = useState<Cell[][]>(sample);
  const [template, setTemplate] = useState<File | null>(null);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteValue, setPasteValue] = useState("");
  const [status, setStatus] = useState("");
  const tableInput = useRef<HTMLInputElement>(null);
  const templateInput = useRef<HTMLInputElement>(null);
  const stampInput = useRef<HTMLInputElement>(null);
  const formattedDate = useMemo(() => { if (!date) return ""; const [year, month, day] = date.split("-"); return `${day}.${month}.${year}`; }, [date]);
  const formattedCity = useMemo(() => { const value = city.trim().replace(/^г\.?\s*/i, ""); return value ? `г. ${value}` : ""; }, [city]);
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

  async function generateDocument(version: "color" | "grayscale") {
    try {
      if (!template) {
        window.alert("Сначала загрузите бланк Word");
        return;
      }
      setStatus(version === "color" ? "Формируем цветную версию…" : "Формируем чёрно-белую версию…");
      const templateBuffer = await template.arrayBuffer();
      const zip = new PizZip(templateBuffer);
      normalizeWordFormatting(zip);
      if (version === "grayscale" && stamp) {
        const stampBuffer = await imageToPng(stamp);
        zip.file("word/media/stamp.png", stampBuffer);
        const relsFile = zip.file("word/_rels/document.xml.rels");
        if (!relsFile) throw new Error("В бланке не найдены связи документа");
        const rels = relsFile.asText();
        if (!rels.includes('Id="rIdStamp"')) zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdStamp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stamp.png"/></Relationships>'));
        const typesFile = zip.file("[Content_Types].xml");
        if (typesFile) { const types = typesFile.asText(); if (!types.includes('Extension="png"')) zip.file("[Content_Types].xml", types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')); }
      }
      if (version === "grayscale") {
        const mediaFiles = Object.keys(zip.files).filter((name) => /^word\/media\/.*\.(png|jpe?g)$/i.test(name) && name !== "word/media/stamp.png");
        await Promise.all(mediaFiles.map(async (name) => {
          const media = zip.file(name);
          if (!media) return;
          const extension = name.split(".").pop() ?? "";
          zip.file(name, await rasterToGrayscale(media.asUint8Array(), extension));
        }));
        grayscaleWordColors(zip);
      }
      const documentFile = zip.file("word/document.xml"); if (!documentFile) throw new Error("В бланке не найден документ Word");
      let original = documentFile.asText();
      if (!original.includes("xmlns:a=")) original = original.replace("<w:document ", '<w:document xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ');
      const section = original.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)?.[0]; if (!section) throw new Error("В бланке отсутствуют параметры страницы");
      const signature = version === "grayscale" && stamp
        ? signatureWithCenteredStamp(directorTitle, directorName)
        : alignedLine(directorTitle, directorName, { before: 2400, after: 0 });
      const content = [alignedLine(`Исх. №${number} от ${formattedDate} года`, formattedCity, { after: 260 }), paragraph(title, { bold: true, center: true, size: 20, after: includeObject && objectName.trim() ? 140 : 260 }), ...(includeObject && objectName.trim() ? [paragraph(`Объект: ${objectName.trim()}`, { after: 220 })] : []), tableXml(rows), paragraph("", { after: 180 }), paragraph(calculation.totalText, { bold: true, after: 320 }), paragraph(delivery, { after: 0 }), signature].join("");
      zip.file("word/document.xml", original.replace(/<w:body>[\s\S]*?<\/w:body>/, `<w:body>${content}${section}</w:body>`));
      const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const suffix = version === "color" ? "цветное" : "черно-белое";
      const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `КП_исх_${number.replaceAll("/", "-")}_${formattedDate}_${suffix}.docx`; link.click(); URL.revokeObjectURL(url);
      setStatus(version === "color" ? "Готово — цветная версия скачана" : "Готово — чёрно-белая версия скачана");
    } catch (error) { setStatus(error instanceof Error ? error.message : "Не удалось сформировать документ"); }
  }

  const columns = Math.max(...rows.map((row) => row.length), 1);
  return <main>
    <section className="workspace">
      <div className="intro"><div><h1>Соберите КП за пару минут</h1><p>Загрузите цветной бланк, заполните реквизиты и скачайте цветную или чёрно-белую версию Word.</p></div><div className="documentActions"><button className="primary" onClick={() => generateDocument("color")}>Скачать цветную версию</button><button className="primary monochromeButton" onClick={() => generateDocument("grayscale")}>Скачать чёрно-белую версию</button></div></div>
      <div className="grid">
        <aside className="panel controls">
          <div className="step"><h2>Бланк</h2></div>
          <input ref={templateInput} hidden type="file" accept=".docx" onChange={(e) => { setTemplate(e.target.files?.[0] ?? null); setStatus(e.target.files?.[0] ? `Бланк выбран: ${e.target.files[0].name}` : ""); }} />
          <div className="buttonRow fullWidthActions"><button className="secondary" onClick={() => templateInput.current?.click()}>{template ? "Заменить бланк" : "Загрузить бланк"}</button>{template && <button className="textButton" onClick={() => { setTemplate(null); if (templateInput.current) templateInput.current.value = ""; }}>Удалить бланк</button>}</div>
          <div className="divider" /><div className="step"><h2>Реквизиты</h2></div>
          <label>Исходящий номер<input value={number} onChange={(e) => setNumber(e.target.value)} /></label><div className="twoCols"><label>Дата<input type="date" value={date} onChange={(e) => setDate(e.target.value)} /></label><label>Город<input value={city} onChange={(e) => setCity(e.target.value)} placeholder="Москва" /></label></div><label>Заголовок<input value={title} onChange={(e) => setTitle(e.target.value)} /></label>
          <label className="checkRow"><input type="checkbox" checked={includeObject} onChange={(e) => setIncludeObject(e.target.checked)} /><span>Добавить название объекта</span></label>
          {includeObject && <label>Название объекта<input value={objectName} onChange={(e) => setObjectName(e.target.value)} placeholder="Например: БЦ «Омега», корпус 2" /></label>}
          <div className="divider" /><div className="step"><h2>Таблица</h2></div>
          <input ref={tableInput} hidden type="file" accept=".xlsx,.xls,.csv" onChange={importSpreadsheet} /><div className="buttonRow tableImportActions"><button className="secondary" onClick={() => tableInput.current?.click()}>Загрузить Excel / CSV</button><button className="secondary" onClick={() => setPasteOpen((v) => !v)}>Вставить из буфера</button></div>
          {pasteOpen && <div className="pasteBox"><textarea autoFocus placeholder="Скопируйте ячейки в Excel и вставьте сюда" value={pasteValue} onChange={(e) => setPasteValue(e.target.value)} /><button className="smallPrimary" onClick={applyPastedTable}>Применить</button></div>}
          <div className="divider" /><div className="step"><h2>Условия и подпись</h2></div>
          <label>Условия доставки<textarea value={delivery} onChange={(e) => setDelivery(e.target.value)} /></label>
          <div className="twoCols"><label>Должность<input value={directorTitle} onChange={(e) => setDirectorTitle(e.target.value)} /></label><label>ФИО директора<input value={directorName} onChange={(e) => setDirectorName(e.target.value)} /></label></div>
          <div className="divider" />
          <div className="stampSection">
            <div className="step"><h2>Печать</h2></div>
          <input ref={stampInput} hidden type="file" accept="image/png,image/jpeg" onChange={(e) => { const file = e.target.files?.[0] ?? null; setStamp(file); setStatus(file ? `Печать выбрана: ${file.name}` : ""); }} />
          <div className="buttonRow fullWidthActions"><button className="secondary" onClick={() => stampInput.current?.click()}>{stamp ? "Заменить печать" : "Загрузить печать"}</button>{stamp && <button className="textButton" onClick={() => { setStamp(null); if (stampInput.current) stampInput.current.value = ""; setStatus("Печать удалена"); }}>Удалить печать</button>}</div>
          </div>
          <div className="uploadNotice">
            <strong>Важно</strong>
            <span>Загружая печать и бланк, пользователь подтверждает право на их использование.</span>
            <span>Все файлы обрабатываются только на устройстве пользователя и не отправляются на сервер.</span>
          </div>
        </aside>
        <section className="panel previewPanel"><div className="previewHeader"><div><h2>Предпросмотр таблицы</h2><p>{rows.length} строк · {columns} столбцов</p></div><span className="editable">Можно редактировать</span></div><div className="tableWrap"><table><tbody>{rows.map((row, ri) => <tr key={ri}>{Array.from({ length: columns }, (_, ci) => <td key={ci} className={ri === 0 ? "headingCell" : ""}><input aria-label={`Строка ${ri + 1}, столбец ${ci + 1}`} value={String(row[ci] ?? "")} onChange={(e) => updateCell(ri, ci, e.target.value)} /></td>)}</tr>)}</tbody></table></div></section>
      </div>
    </section>
  </main>;
}
