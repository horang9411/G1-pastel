import assert from "node:assert/strict";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the finished Gmarket invoice service", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<html lang="ko">/i);
  assert.match(html, /<title>G마켓 송장 입,출력용 딸깍<\/title>/i);
  assert.doesNotMatch(html, /G마켓 - CJ처리기/);
  assert.match(html, /Gmarket \/ Pastelcraft/);
  assert.match(html, /Invoice Machine/);
  assert.match(html, /G마켓 파일 업로드/);
  assert.match(html, /택배사·송장번호 자동 입력/);
  assert.match(html, /made by 영중팀장/);
  assert.match(html, /accept="\.xlsx,\.xls"/);
  assert.doesNotMatch(html, /codex-preview|Building your site|Starter Project/i);
});

test("renders accessible upload controls and privacy notice", async () => {
  const response = await render();
  const html = await response.text();

  assert.match(html, /aria-label="G마켓 주문 엑셀 업로드"/);
  assert.match(html, /role="button"/);
  assert.match(html, /tabindex="0"/);
  assert.doesNotMatch(html, /파일은 서버에 저장되지 않아요/);
  assert.match(html, /CJ 출력확정 파일/);
  assert.match(html, /G마켓 발송관리 파일/);
});
