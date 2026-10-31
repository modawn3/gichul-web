// 문항·해설 그림 자동 올리기: GitHub 저장소의 _upload/pack_NNN.bin(암호화 묶음)을 받아 풀어서
// 비공개 저장소 items에 올린다. 키가 맞지 않는 묶음은 풀리지 않으므로 아무나 불러도 다른 그림은 못 넣는다.
// 요청: POST {"sha": "<커밋 40자리>", "pack": 0}  (pack 없으면 묶음 개수만 알려 줌)
import { createClient } from "jsr:@supabase/supabase-js@2";

const REPO = "modawn3/gichul-web";
const json = (b: unknown, status = 200) =>
  new Response(JSON.stringify(b), { status, headers: { "Content-Type": "application/json" } });
const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, {
  auth: { persistSession: false, autoRefreshToken: false },
});

async function aesKey() {
  let k = Deno.env.get("IMG_PACK_KEY");
  if (!k) { const { data } = await sb.rpc("img_pack_key"); k = data; }
  if (!k) throw new Error("키가 없습니다");
  return crypto.subtle.importKey("raw", Uint8Array.from(atob(k), (c) => c.charCodeAt(0)), "AES-GCM", false, ["decrypt"]);
}
const raw = async (sha: string, f: string) => {
  const r = await fetch(`https://raw.githubusercontent.com/${REPO}/${sha}/_upload/${f}`);
  if (!r.ok) throw new Error(`${f}: ${r.status}`);
  return r;
};

Deno.serve(async (req) => {
  if (req.method !== "POST") return json({ error: "POST only" }, 405);
  let body: any;
  try { body = await req.json(); } catch { return json({ error: "잘못된 요청" }, 400); }
  const sha = String(body.sha ?? "");
  if (!/^[0-9a-f]{40}$/.test(sha)) return json({ error: "sha는 커밋 40자리" }, 400);
  try {
    const man = await (await raw(sha, "manifest.json")).json();
    if (body.pack === undefined) return json({ packs: man.packs.length });
    const i = Number(body.pack), p = man.packs[i];
    if (!p) return json({ error: "없는 묶음" }, 404);
    const { data: done } = await sb.from("img_ingest_log").select("n").eq("sha", sha).eq("pack", i).maybeSingle();
    if (done) return json({ pack: i, skipped: true, n: done.n });

    const buf = new Uint8Array(await (await raw(sha, p.file)).arrayBuffer());
    let plain: ArrayBuffer;
    try { plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: buf.slice(0, 12) }, await aesKey(), buf.slice(12)); }
    catch { return json({ error: "묶음을 풀 수 없습니다(키 불일치)" }, 403); }
    const text = await new Response(new Blob([plain]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
    const todo = Object.entries(JSON.parse(text) as Record<string, string>)
      .filter(([path]) => /^[qs]\/[\x21-\x7e]+\.png$/.test(path));
    const fail: string[] = [];
    const next = async () => {
      while (todo.length) {
        const [path, b64] = todo.shift()!;
        const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
        let error;
        for (let t = 0; t < 3; t++) {
          ({ error } = await sb.storage.from("items").upload(path, bytes, { contentType: "image/png", upsert: true, cacheControl: "604800" }));
          if (!error) break;
        }
        if (error) fail.push(`${path}: ${error.message}`);
      }
    };
    await Promise.all(Array.from({ length: 6 }, next));
    if (!fail.length) await sb.from("img_ingest_log").insert({ sha, pack: i, n: p.n });
    return json({ pack: i, n: p.n, fail });
  } catch (e) {
    return json({ error: String((e as Error).message ?? e) }, 500);
  }
});
