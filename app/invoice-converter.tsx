"use client";

import { ChangeEvent, DragEvent, useCallback, useRef, useState } from "react";

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

async function convertWorkbook(file: File): Promise<ConversionResult> {
  const XLSX = await import("xlsx-js-style");
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
      const option = textValue(row[indexOf("옵션")]);
      return {
        orderNumber: outputValue(row[indexOf("주문번호")]),
        productCode: outputValue(row[indexOf("상품번호")]),
        productName: textValue(row[indexOf("상품명")]),
        recipientName: textValue(row[indexOf("수령인명")]),
        buyerName: textValue(row[indexOf("구매자명")]),
        quantity: Number(row[indexOf("수량")]) || 0,
        itemName: option && option !== "67" ? option : "76",
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

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#" aria-label="처음으로">
          <span className="brand-mark">G</span>
          <span>송장 딸깍</span>
        </a>
        <span className="privacy-pill">
          <span className="privacy-dot" />
          파일은 서버에 저장되지 않아요
        </span>
      </header>

      <section className="hero">
        <div className="eyebrow">
          <span>G마켓 주문서</span>
          <span className="eyebrow-arrow">→</span>
          <span>송장 출력 양식</span>
        </div>
        <h1>
          G마켓 송장 입,출력용
          <br />
          <span>딸깍</span>
        </h1>
        <p className="hero-copy">
          복잡한 열 정리는 이제 그만.
          <br className="mobile-break" /> 주문 엑셀을 올리면 송장용 파일이
          바로 완성됩니다.
        </p>
      </section>

      <section className="workspace" aria-label="엑셀 변환">
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
              <h2>
                {selectedName || "1-1 형태의 엑셀을 여기에 놓아주세요"}
              </h2>
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

      <section className="how-it-works" aria-label="이용 방법">
        <div className="section-heading">
          <span>이용 방법</span>
          <h2>세 단계면 끝나요</h2>
        </div>
        <ol>
          <li>
            <span className="step-number">01</span>
            <strong>주문서 올리기</strong>
            <p>G마켓에서 내려받은 원본 주문 엑셀을 선택해요.</p>
          </li>
          <li>
            <span className="step-number">02</span>
            <strong>자동 변환</strong>
            <p>필요한 13개 열만 정확한 송장 순서로 정리해요.</p>
          </li>
          <li>
            <span className="step-number">03</span>
            <strong>결과 받기</strong>
            <p>완성된 송장 출력용 엑셀을 바로 내려받아요.</p>
          </li>
        </ol>
      </section>

      <footer>
        <p>G마켓 송장 입,출력용 딸깍</p>
        <span>개인정보가 담긴 파일은 내 브라우저 안에서만 처리됩니다.</span>
      </footer>
    </main>
  );
}
