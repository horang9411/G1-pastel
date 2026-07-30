import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

const title = "G마켓 송장 입,출력용 딸깍";
const description =
  "G마켓 주문 엑셀을 송장 출력 양식으로 한 번에 변환하세요.";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const protocol = requestHeaders.get("x-forwarded-proto") ?? "https";
  const host = requestHeaders.get("host") ?? "localhost:3000";
  const imageUrl = new URL("/og.png", `${protocol}://${host}`).toString();

  return {
    title,
    description,
    icons: {
      icon: "/favicon.svg",
      shortcut: "/favicon.svg",
    },
    openGraph: {
      title,
      description,
      images: [{ url: imageUrl, width: 1200, height: 630 }],
    },
    twitter: {
      card: "summary_large_image",
      title,
      description,
      images: [imageUrl],
    },
  };
}

export default function RootLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
