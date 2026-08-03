"use client";

import {
  ChangeEvent,
  DragEvent,
  useCallback,
  useRef,
  useState,
} from "react";

type ConvertedRow = {
  orderNumber: string | number;
  productCode: string | number;
  productName: string;
  recipientName: string;
  buyerName: string;
  quantity: number;
  itemName: string;
  recipientPhone: string;
  postalCode: string;
  address: string;
  buyerPhone: string;
  cartNumber: string | number;
  boxType: string;
};

type ConversionResult = {
  fileName: string;
  rowCount: number;
  preview: ConvertedRow[];
  blob: Blob;
};

type ShippingMergeResult = {
  fileName: string;
  matchedCount: number;
  totalCount: number;
  defaultCarrierCount: number;
  blob: Blob;
};

const OUTPUT_HEADERS = [
  "고객주문번호",
  "품목코드",
  "품목명",
  "받는분성명",
  "주문자성명",
  "내품수량",
  "내품명",
  "받는분전화번호",
  "받는분우편번호",
  "받는분주소(전체, 분할)",
  "주문자전화번호",
  "기타2",
  "박스타입",
];

const REQUIRED_HEADERS = [
  "주문번호",
  "상품번호",
  "상품명",
  "수령인명",
  "구매자명",
  "수량",
  "옵션",
  "수령인 휴대폰",
  "우편번호",
  "주소",
  "구매자 휴대폰",
  "장바구니번호(결제번호)",
];

const PREVIEW_COLUMNS = [
  { key: "orderNumber", label: "주문번호" },
  { key: "recipientName", label: "받는 분" },
  { key: "productName", label: "품목명" },
  { key: "quantity", label: "수량" },
] as const;

function fileStamp() {
  const now = new Date();
  return [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "_",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
  ].join("");
}

function textValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function outputValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return value as string | number;
}

function normalizedOrderNumber(value: unknown) {
  return textValue(value).replace(/\.0$/, "");
}

function headerIndex(headers: string[], candidates: string[], label: string) {
  const index = candidates
    .map((candidate) => headers.indexOf(candidate))
    .find((candidateIndex) => candidateIndex >= 0);
  if (index === undefined) {
    throw new Error(`${label} 열을 찾지 못했습니다.`);
  }
  return index;
}

function optionalHeaderIndex(headers: string[], candidates: string[]) {
  return candidates
    .map((candidate) => headers.indexOf(candidate))
    .find((candidateIndex) => candidateIndex >= 0);
}

function normalizeCarrier(value: unknown) {
  const carrier = textValue(value);
  const compact = carrier.replace(/\s+/g, "").toLocaleLowerCase("ko-KR");
  if (!carrier || /cj|씨제이|대한통운/.test(compact)) return "CJ택배";
  if (/롯데/.test(compact)) return "롯데택배";
  if (/한진/.test(compact)) return "한진택배";
  if (/로젠/.test(compact)) return "로젠택배";
  if (/우체국/.test(compact)) return "우체국택배";
  return carrier;
}

function xmlAttribute(tag: string, name: string) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return tag.match(new RegExp(`\\b${escapedName}="([^"]*)"`))?.[1] ?? "";
}

function firstWorksheetPath(files: Record<string, Uint8Array>) {
  const decoder = new TextDecoder();
  const workbookFile = files["xl/workbook.xml"];
  const relationshipsFile = files["xl/_rels/workbook.xml.rels"];
  if (!workbookFile || !relationshipsFile) {
    throw new Error("2번 파일의 엑셀 내부 구조를 읽지 못했습니다.");
  }

  const workbookXml = decoder.decode(workbookFile);
  const relationshipsXml = decoder.decode(relationshipsFile);
  const firstSheetTag =
    workbookXml.match(/<(?:\w+:)?sheet\b[^>]*>/)?.[0] ?? "";
  const relationshipId = xmlAttribute(firstSheetTag, "r:id");
  const relationshipTag = [...relationshipsXml.matchAll(/<Relationship\b[^>]*>/g)]
    .map((match) => match[0])
    .find((tag) => xmlAttribute(tag, "Id") === relationshipId);
  const target = relationshipTag ? xmlAttribute(relationshipTag, "Target") : "";
  if (!target) {
    throw new Error("2번 파일의 첫 번째 시트를 찾지 못했습니다.");
  }

  return target.startsWith("/")
    ? target.slice(1)
    : `xl/${target.replace(/^\.?\//, "")}`;
}

function columnNumber(columnLetters: string) {
  return [...columnLetters].reduce(
    (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
    0,
  );
}

function inlineStringCell(
  address: string,
  value: string,
  style = "",
  namespacePrefix = "",
) {
  const escapedValue = value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const preserveSpace = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : "";
  const styleAttribute = style ? ` s="${style}"` : "";
  return `<${namespacePrefix}c r="${address}"${styleAttribute} t="inlineStr"><${namespacePrefix}is><${namespacePrefix}t${preserveSpace}>${escapedValue}</${namespacePrefix}t></${namespacePrefix}is></${namespacePrefix}c>`;
}

function updateSheetCell(sheetXml: string, address: string, value: string) {
  const rowNumber = address.match(/\d+$/)?.[0];
  const columnLetters = address.match(/^[A-Z]+/)?.[0];
  if (!rowNumber || !columnLetters) {
    throw new Error("엑셀 셀 주소를 처리하지 못했습니다.");
  }

  const rowPattern = new RegExp(
    `<(?:\\w+:)?row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/(?:\\w+:)?row>`,
  );
  const rowMatch = sheetXml.match(rowPattern);
  if (!rowMatch) {
    throw new Error(`2번 파일에서 ${rowNumber}행을 찾지 못했습니다.`);
  }

  const originalRow = rowMatch[0];
  const namespacePrefix =
    originalRow.match(/^<(\w+:)?row\b/)?.[1] ?? "";
  const cellPattern = new RegExp(
    `<(?:\\w+:)?c\\b(?=[^>]*\\br="${address}")(?:[^>]*\\/\\s*>|[^>]*>[\\s\\S]*?<\\/(?:\\w+:)?c>)`,
  );
  const existingCell = originalRow.match(cellPattern)?.[0];
  const style = existingCell?.match(/\bs="([^"]*)"/)?.[1] ?? "";
  const replacementCell = inlineStringCell(
    address,
    value,
    style,
    namespacePrefix,
  );
  let updatedRow: string;

  if (existingCell) {
    updatedRow = originalRow.replace(cellPattern, replacementCell);
  } else {
    const closingRowIndex = originalRow.lastIndexOf(
      `</${namespacePrefix}row>`,
    );
    let insertAt = closingRowIndex;
    const targetColumn = columnNumber(columnLetters);
    const cellTags =
      originalRow.match(
        /<(?:\w+:)?c\b(?=[^>]*\br="[A-Z]+\d+")(?:[^>]*\/\s*>|[^>]*>[\s\S]*?<\/(?:\w+:)?c>)/g,
      ) ?? [];

    for (const cellTag of cellTags) {
      const existingAddress = xmlAttribute(cellTag, "r");
      const existingColumnLetters = existingAddress.match(/^[A-Z]+/)?.[0];
      if (
        existingColumnLetters &&
        columnNumber(existingColumnLetters) > targetColumn
      ) {
        insertAt = originalRow.indexOf(cellTag);
        break;
      }
    }
    updatedRow =
      originalRow.slice(0, insertAt) +
      replacementCell +
      originalRow.slice(insertAt);
  }

  return sheetXml.replace(rowPattern, updatedRow);
}

async function patchWorkbookCells(
  originalWorkbook: ArrayBuffer,
  updates: Array<{ address: string; value: string }>,
) {
  const { strFromU8, strToU8, unzipSync, zipSync } = await import("fflate");
  const files = unzipSync(new Uint8Array(originalWorkbook));
  const sheetPath = firstWorksheetPath(files);
  const sheetFile = files[sheetPath];
  if (!sheetFile) {
    throw new Error("2번 파일의 첫 번째 시트 내용을 찾지 못했습니다.");
  }

  let sheetXml = strFromU8(sheetFile);
  for (const update of updates) {
    sheetXml = updateSheetCell(sheetXml, update.address, update.value);
  }
  files[sheetPath] = strToU8(sheetXml);
  return zipSync(files, { level: 6 });
}

export async function fillShippingWorkbook(
  invoiceResultFile: File,
  gmarketFile: File,
): Promise<ShippingMergeResult> {
  const XLSXModule = await import("xlsx-js-style");
  const XLSX = XLSXModule.default ?? XLSXModule;
  const invoiceWorkbook = XLSX.read(await invoiceResultFile.arrayBuffer(), {
    type: "array",
    cellDates: true,
    raw: true,
  });
  const invoiceSheetName = invoiceWorkbook.SheetNames[0];
  if (!invoiceSheetName) throw new Error("CJ 출력확정 파일에 시트가 없습니다.");
  const invoiceRows = XLSX.utils.sheet_to_json<unknown[]>(
    invoiceWorkbook.Sheets[invoiceSheetName],
    { header: 1, defval: "", raw: true },
  );
  const invoiceHeaders = (invoiceRows[0] ?? []).map(textValue);
  const invoiceCartIndex = headerIndex(
    invoiceHeaders,
    ["기타2", "장바구니번호(결제번호)"],
    "기타2",
  );
  const invoiceTrackingIndex = headerIndex(
    invoiceHeaders,
    ["운송장번호", "송장번호", "운송장 번호", "등기번호"],
    "송장번호",
  );
  const invoiceCarrierIndex = optionalHeaderIndex(invoiceHeaders, [
    "택배사",
    "택배사명",
    "택배사명(발송방법)",
    "배송사",
  ]);

  const invoiceLookup = new Map<
    string,
    { carrier: string; trackingNumber: string; usedDefaultCarrier: boolean }
  >();
  for (const row of invoiceRows.slice(1)) {
    const trackingNumber = textValue(row[invoiceTrackingIndex]);
    if (!trackingNumber) continue;
    const key = normalizedOrderNumber(row[invoiceCartIndex]);
    if (!key) continue;
    const rawCarrier =
      invoiceCarrierIndex === undefined ? "" : row[invoiceCarrierIndex];
    const carrier = normalizeCarrier(rawCarrier);
    const usedDefaultCarrier = textValue(rawCarrier) === "";
    const existing = invoiceLookup.get(key);
    if (existing && existing.trackingNumber !== trackingNumber) {
      throw new Error(
        "같은 결제번호에 서로 다른 송장번호가 있어 CJ 송장 파일을 확인해 주세요.",
      );
    }
    invoiceLookup.set(key, { carrier, trackingNumber, usedDefaultCarrier });
  }
  if (!invoiceLookup.size) {
    throw new Error("CJ 출력확정 파일에서 운송장번호를 찾지 못했습니다.");
  }

  if (!gmarketFile.name.toLocaleLowerCase().endsWith(".xlsx")) {
    throw new Error("2번 G마켓 발송관리 파일은 .xlsx 형식으로 올려주세요.");
  }
  const gmarketBuffer = await gmarketFile.arrayBuffer();
  const gmarketWorkbook = XLSX.read(gmarketBuffer, {
    type: "array",
    cellDates: false,
    cellNF: true,
    raw: true,
  });
  const gmarketSheetName = gmarketWorkbook.SheetNames[0];
  if (!gmarketSheetName) throw new Error("G마켓 파일에 시트가 없습니다.");
  const gmarketSheet = gmarketWorkbook.Sheets[gmarketSheetName];
  const gmarketRows = XLSX.utils.sheet_to_json<unknown[]>(gmarketSheet, {
    header: 1,
    defval: "",
    raw: true,
  });
  const gmarketHeaders = (gmarketRows[0] ?? []).map(textValue);
  const gmarketCartIndex = headerIndex(
    gmarketHeaders,
    ["장바구니번호(결제번호)", "기타2"],
    "장바구니번호(결제번호)",
  );
  const gmarketCarrierIndex = headerIndex(
    gmarketHeaders,
    ["택배사명(발송방법)"],
    "택배사",
  );
  const gmarketTrackingIndex = headerIndex(
    gmarketHeaders,
    ["송장번호"],
    "송장번호",
  );

  let matchedCount = 0;
  let defaultCarrierCount = 0;
  const cellUpdates: Array<{ address: string; value: string }> = [];
  const dataRows = gmarketRows.slice(1).filter((row) =>
    row.some((cell) => textValue(cell) !== ""),
  );
  for (let rowIndex = 1; rowIndex < gmarketRows.length; rowIndex += 1) {
    const row = gmarketRows[rowIndex];
    if (!row.some((cell) => textValue(cell) !== "")) continue;
    const key = normalizedOrderNumber(row[gmarketCartIndex]);
    const match = invoiceLookup.get(key);
    if (!match) continue;

    const carrierAddress = XLSX.utils.encode_cell({
      r: rowIndex,
      c: gmarketCarrierIndex,
    });
    const trackingAddress = XLSX.utils.encode_cell({
      r: rowIndex,
      c: gmarketTrackingIndex,
    });
    cellUpdates.push(
      { address: carrierAddress, value: match.carrier },
      { address: trackingAddress, value: match.trackingNumber },
    );
    matchedCount += 1;
    if (match.usedDefaultCarrier) defaultCarrierCount += 1;
  }

  if (!dataRows.length) {
    throw new Error(
      "G마켓 발송관리 파일에 주문 데이터가 없습니다. 주문이 포함된 파일을 넣어주세요.",
    );
  }
  if (!matchedCount) {
    throw new Error(
      "1번 파일의 기타2와 2번 파일의 장바구니번호(결제번호)가 일치하는 주문을 찾지 못했습니다.",
    );
  }

  const output = await patchWorkbookCells(gmarketBuffer, cellUpdates);
  return {
    fileName: `G마켓_발송처리_${fileStamp()}.xlsx`,
    matchedCount,
    totalCount: dataRows.length,
    defaultCarrierCount,
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

async function convertWorkbook(file: File): Promise<ConversionResult> {
  const XLSXModule = await import("xlsx-js-style");
  const XLSX = XLSXModule.default ?? XLSXModule;
  const source = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: true,
    raw: true,
  });
  const firstSheetName = source.SheetNames[0];
  if (!firstSheetName) throw new Error("엑셀 안에 시트가 없습니다.");

  const sourceRows = XLSX.utils.sheet_to_json<unknown[]>(
    source.Sheets[firstSheetName],
    { header: 1, defval: "", raw: true },
  );
  const headerRow = (sourceRows[0] ?? []).map(textValue);
  const missing = REQUIRED_HEADERS.filter(
    (header) => !headerRow.includes(header),
  );
  if (missing.length) {
    throw new Error(`필수 열을 찾지 못했습니다: ${missing.join(", ")}`);
  }

  const indexOf = (header: string) => headerRow.indexOf(header);
  const rows: ConvertedRow[] = sourceRows
    .slice(1)
    .filter((row) => row.some((cell) => textValue(cell) !== ""))
    .map((row) => {
      return {
        orderNumber: outputValue(row[indexOf("주문번호")]),
        productCode: outputValue(row[indexOf("상품번호")]),
        productName: textValue(row[indexOf("상품명")]),
        recipientName: textValue(row[indexOf("수령인명")]),
        buyerName: textValue(row[indexOf("구매자명")]),
        quantity: Number(row[indexOf("수량")]) || 0,
        itemName: textValue(row[indexOf("옵션")]),
        recipientPhone: textValue(row[indexOf("수령인 휴대폰")]),
        postalCode: textValue(row[indexOf("우편번호")]),
        address: textValue(row[indexOf("주소")]),
        buyerPhone: textValue(row[indexOf("구매자 휴대폰")]),
        cartNumber: outputValue(row[indexOf("장바구니번호(결제번호)")]),
        boxType: "극소",
      };
    });

  if (!rows.length) throw new Error("변환할 주문 데이터가 없습니다.");

  const outputRows = rows.map((row) => [
    row.orderNumber,
    row.productCode,
    row.productName,
    row.recipientName,
    row.buyerName,
    row.quantity,
    row.itemName,
    row.recipientPhone,
    row.postalCode,
    row.address,
    row.buyerPhone,
    row.cartNumber,
    row.boxType,
  ]);

  const worksheet = XLSX.utils.aoa_to_sheet([OUTPUT_HEADERS, ...outputRows]);
  worksheet["!cols"] = [
    { wch: 15 },
    { wch: 14 },
    { wch: 42 },
    { wch: 13 },
    { wch: 19 },
    { wch: 10 },
    { wch: 38 },
    { wch: 17 },
    { wch: 16 },
    { wch: 58 },
    { wch: 17 },
    { wch: 15 },
    { wch: 10 },
  ];
  worksheet["!rows"] = [{ hpt: 22 }];
  worksheet["!autofilter"] = { ref: `A1:M${rows.length + 1}` };

  const headerStyle = {
    font: { name: "맑은 고딕", sz: 11 },
    fill: { patternType: "solid", fgColor: { rgb: "DDDDDD" } },
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } },
    },
    alignment: { horizontal: "center", vertical: "center" },
  };
  const bodyStyle = {
    font: { name: "맑은 고딕", sz: 11 },
    alignment: { vertical: "center" },
  };

  for (let column = 0; column < OUTPUT_HEADERS.length; column += 1) {
    const headerAddress = XLSX.utils.encode_cell({ r: 0, c: column });
    worksheet[headerAddress].s = headerStyle;
  }
  for (let row = 1; row <= rows.length; row += 1) {
    for (let column = 0; column < OUTPUT_HEADERS.length; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      if (worksheet[address]) worksheet[address].s = bodyStyle;
    }
  }
  for (const column of [1, 7, 8, 10]) {
    for (let row = 1; row <= rows.length; row += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      if (worksheet[address]) worksheet[address].z = "@";
    }
  }

  const output = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(output, worksheet, "Sheet1");
  const buffer = XLSX.write(output, {
    type: "array",
    bookType: "xlsx",
    cellStyles: true,
  });
  const fileName = `G마켓_송장입출력_${fileStamp()}.xlsx`;

  return {
    fileName,
    rowCount: rows.length,
    preview: rows.slice(0, 5),
    blob: new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  };
}

export function InvoiceConverter() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [selectedName, setSelectedName] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [invoiceResultFile, setInvoiceResultFile] = useState<File | null>(null);
  const [gmarketShippingFile, setGmarketShippingFile] = useState<File | null>(
    null,
  );
  const [isMergingShipping, setIsMergingShipping] = useState(false);
  const [shippingError, setShippingError] = useState("");
  const [shippingResult, setShippingResult] =
    useState<ShippingMergeResult | null>(null);
  const shippingRequestIdRef = useRef(0);

  const processFile = useCallback(async (file?: File) => {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) {
      setError("엑셀 파일(.xlsx 또는 .xls)만 업로드해 주세요.");
      return;
    }
    setSelectedName(file.name);
    setError("");
    setResult(null);
    setIsConverting(true);
    try {
      setResult(await convertWorkbook(file));
    } catch (caught) {
      setError(
        caught instanceof Error
          ? caught.message
          : "파일을 변환하는 중 문제가 생겼습니다.",
      );
    } finally {
      setIsConverting(false);
    }
  }, []);

  const onFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    void processFile(event.target.files?.[0]);
    event.target.value = "";
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    void processFile(event.dataTransfer.files?.[0]);
  };

  const download = () => {
    if (!result) return;
    const url = URL.createObjectURL(result.blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = result.fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const processShippingFiles = useCallback(
    async (nextInvoiceFile: File | null, nextGmarketFile: File | null) => {
      const requestId = shippingRequestIdRef.current + 1;
      shippingRequestIdRef.current = requestId;
      setIsMergingShipping(Boolean(nextInvoiceFile && nextGmarketFile));
      setShippingError("");
      setShippingResult(null);
      if (!nextInvoiceFile || !nextGmarketFile) return;

      try {
        const mergedResult = await fillShippingWorkbook(
          nextInvoiceFile,
          nextGmarketFile,
        );
        if (shippingRequestIdRef.current === requestId) {
          setShippingResult(mergedResult);
        }
      } catch (caught) {
        if (shippingRequestIdRef.current !== requestId) return;
        setShippingError(
          caught instanceof Error
            ? caught.message
            : "발송관리 파일을 처리하는 중 문제가 생겼습니다.",
        );
      } finally {
        if (shippingRequestIdRef.current === requestId) {
          setIsMergingShipping(false);
        }
      }
    },
    [],
  );

  const selectInvoiceResultFile = (file: File | null) => {
    setInvoiceResultFile(file);
    void processShippingFiles(file, gmarketShippingFile);
  };

  const selectGmarketShippingFile = (file: File | null) => {
    setGmarketShippingFile(file);
    void processShippingFiles(invoiceResultFile, file);
  };

  const resetShippingResult = () => {
    shippingRequestIdRef.current += 1;
    setIsMergingShipping(false);
    setShippingError("");
    setShippingResult(null);
  };

  const downloadShippingResult = () => {
    if (!shippingResult) return;
    const url = URL.createObjectURL(shippingResult.blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = shippingResult.fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  return (
    <main>
      <section className="hero">
        <h1>
          G마켓 송장 입,출력용
          <br />
          <span>파스텔 전용 처리기</span>
        </h1>
      </section>

      <section className="workspace" id="excel-tools" aria-label="엑셀 변환">
        <div
          className={`dropzone ${isDragging ? "is-dragging" : ""} ${
            result ? "has-file" : ""
          }`}
          onDragEnter={(event) => {
            event.preventDefault();
            setIsDragging(true);
          }}
          onDragOver={(event) => event.preventDefault()}
          onDragLeave={() => setIsDragging(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") {
              inputRef.current?.click();
            }
          }}
          role="button"
          tabIndex={0}
          aria-label="G마켓 주문 엑셀 업로드"
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={onFileChange}
            hidden
          />
          <div className="file-icon" aria-hidden="true">
            <span>X</span>
          </div>
          {isConverting ? (
            <>
              <h2>송장 양식으로 바꾸는 중...</h2>
              <div className="progress-track">
                <span />
              </div>
              <p>주문 데이터를 안전하게 정리하고 있어요.</p>
            </>
          ) : (
            <>
              <h2>{selectedName || "G마켓 파일 업로드"}</h2>
              <p>
                파일을 끌어다 놓거나{" "}
                <span className="text-link">내 컴퓨터에서 선택</span>
              </p>
              <span className="file-hint">XLSX · XLS</span>
            </>
          )}
        </div>

        {error && (
          <div className="alert error-alert" role="alert">
            <span>!</span>
            <div>
              <strong>파일을 확인해 주세요</strong>
              <p>{error}</p>
            </div>
          </div>
        )}

        {result && (
          <div className="result-card">
            <div className="result-heading">
              <div className="success-icon" aria-hidden="true">
                ✓
              </div>
              <div>
                <span className="result-kicker">변환 완료</span>
                <h2>{result.rowCount.toLocaleString("ko-KR")}건을 정리했어요</h2>
              </div>
              <button className="download-button" onClick={download}>
                결과 엑셀 내려받기
                <span aria-hidden="true">↓</span>
              </button>
            </div>

            <div className="preview-wrap">
              <div className="preview-label">
                <span>미리보기</span>
                <span>앞의 {result.preview.length}건</span>
              </div>
              <div className="table-scroll">
                <table>
                  <thead>
                    <tr>
                      {PREVIEW_COLUMNS.map((column) => (
                        <th key={column.key}>{column.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {result.preview.map((row, index) => (
                      <tr key={`${row.orderNumber}-${index}`}>
                        {PREVIEW_COLUMNS.map((column) => (
                          <td key={column.key}>{row[column.key]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </section>

      <section className="shipping-tool" aria-labelledby="shipping-tool-title">
        <div className="shipping-tool-heading">
          <span className="tool-number">02</span>
          <div>
            <h2 id="shipping-tool-title">택배사·송장번호 자동 입력</h2>
            <p>
              CJ 출력확정 파일 1과 G마켓 발송관리 파일 2를 넣으면 결제번호가
              같은 모든 주문에 자동으로 입력됩니다.
            </p>
          </div>
        </div>

        <div className="shipping-upload-grid">
          <label className={`mini-upload ${invoiceResultFile ? "is-ready" : ""}`}>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(event) => {
                const file = event.target.files?.[0] ?? null;
                resetShippingResult();
                selectInvoiceResultFile(file);
                event.target.value = "";
              }}
            />
            <span className="mini-upload-step">1</span>
            <strong>1. CJ 출력확정 파일</strong>
            <span className="mini-upload-name">
              {invoiceResultFile?.name || "운송장번호가 들어간 1번 엑셀"}
            </span>
            <span className="mini-upload-action">
              {invoiceResultFile ? "파일 다시 선택" : "파일 선택"}
            </span>
          </label>

          <span className="merge-arrow" aria-hidden="true">
            +
          </span>

          <label
            className={`mini-upload ${gmarketShippingFile ? "is-ready" : ""}`}
          >
            <input
              type="file"
              accept=".xlsx"
              onChange={(event) => {
                const file = event.target.files?.[0] ?? null;
                resetShippingResult();
                selectGmarketShippingFile(file);
                event.target.value = "";
              }}
            />
            <span className="mini-upload-step">2</span>
            <strong>2. G마켓 발송관리 파일</strong>
            <span className="mini-upload-name">
              {gmarketShippingFile?.name ||
                "택배사·송장번호를 채울 2번 XLSX 파일"}
            </span>
            <span className="mini-upload-action">
              {gmarketShippingFile ? "파일 다시 선택" : "파일 선택"}
            </span>
          </label>
        </div>

        {isMergingShipping && (
          <div className="shipping-status" role="status">
            <span className="status-spinner" />
            주문 정보를 대조하고 있어요...
          </div>
        )}

        {shippingError && (
          <div className="alert error-alert shipping-alert" role="alert">
            <span>!</span>
            <div>
              <strong>두 파일을 확인해 주세요</strong>
              <p>{shippingError}</p>
            </div>
          </div>
        )}

        {shippingResult && (
          <div className="shipping-success">
            <div>
              <span className="shipping-success-mark">✓</span>
              <div>
                <strong>
                  {shippingResult.totalCount.toLocaleString("ko-KR")}건 중{" "}
                  {shippingResult.matchedCount.toLocaleString("ko-KR")}건 입력
                  완료
                </strong>
                <p>
                  택배사와 송장번호가 입력된 G마켓 파일을 내려받으세요.
                  {shippingResult.defaultCarrierCount > 0 &&
                    ` 택배사 정보가 없던 ${shippingResult.defaultCarrierCount.toLocaleString(
                      "ko-KR",
                    )}건은 CJ택배로 입력했습니다.`}
                </p>
              </div>
            </div>
            <button onClick={downloadShippingResult}>
              완성 파일 내려받기
              <span aria-hidden="true">↓</span>
            </button>
          </div>
        )}
      </section>

      <p className="made-by">made by 영중팀장</p>
    </main>
  );
}
