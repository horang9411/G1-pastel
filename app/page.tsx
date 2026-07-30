import type { Metadata } from "next";
import { InvoiceConverter } from "./invoice-converter";

export const metadata: Metadata = {
  title: "G마켓 송장 입,출력용 딸깍",
  description: "G마켓 주문 엑셀을 송장 출력 양식으로 한 번에 변환하세요.",
};

export default function Home() {
  return <InvoiceConverter />;
}
