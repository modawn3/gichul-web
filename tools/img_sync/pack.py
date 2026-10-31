# 오프라인 프로그램의 img_*.js(문제 그림 img_q*, 해설 그림 img_s*)에서 그림을 꺼내,
# Supabase 저장소(items)에 아직 없거나 바뀐 것만 암호화 묶음으로 만든다.
#   python3 pack.py <저장소 폴더> <있는그림.json> <img_*.js ...>
# 키는 환경변수 IMG_PACK_KEY(base64 32바이트, Supabase Vault 'img_pack_key').
# 있는그림.json은 회차(그룹)별 요약 {"q/2024_03h": "<md5>", ...} — 아래 SQL 결과를 그대로 저장한다.
#   select json_object_agg(g, h) from (select regexp_replace(name, '_\d\d(_\d+)?\.png$', '') g,
#     md5(string_agg(name || ':' || trim(both '"' from metadata->>'eTag'), ',' order by name collate "C")) h
#     from storage.objects where bucket_id = 'items' group by 1) t
# 요약이 다른 회차(새 회차, 그림이 바뀐 회차)는 그 회차 그림을 전부 다시 올린다.
# 결과: <저장소>/_upload/manifest.json, pack_000.bin … 이걸 GitHub에 올린 뒤 img-ingest 함수를 부르면 저장소에 들어간다.
import base64, gzip, hashlib, json, os, re, sys
from pathlib import Path
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

PACK_MAX = 6_000_000  # 묶음 하나의 그림 크기 합(바이트). 함수 한 번에 하나씩 처리.

def img_key(i):  # web.js imgKey와 같아야 함
    i = i.replace('월학력평가', 'h').replace('월모의평가', 'm').replace('수능', 'csat')
    return ''.join(c if ord(c) < 128 else 'u%x' % ord(c) for c in i)

def extract(text):
    for m in re.finditer(r'"((?:[a-z]+_)?\d{4}_[^"]+?_\d\d)":("data:image/png;base64,[^"]+"|\[[^\]]*\])', text):
        if m[2][0] == '"':
            yield f'q/{img_key(m[1])}.png', m[2][23:-1]
        else:
            for k, x in enumerate(re.findall(r'"data:image/png;base64,([^"]+)"', m[2])):
                yield f's/{img_key(m[1])}_{k + 1}.png', x

def main():
    repo, have_f, srcs = Path(sys.argv[1]), sys.argv[2], sys.argv[3:]
    key = base64.b64decode(os.environ['IMG_PACK_KEY'])
    have = json.load(open(have_f)) or {}
    groups = {}
    for f in srcs:
        for path, b64 in extract(Path(f).read_text()):
            groups.setdefault(re.sub(r'_\d\d(_\d+)?\.png$', '', path), {})[path] = b64
    todo = {}
    for g, imgs in groups.items():
        h = hashlib.md5(','.join(f'{p}:{hashlib.md5(base64.b64decode(b)).hexdigest()}' for p, b in sorted(imgs.items())).encode()).hexdigest()
        if have.get(g) != h:
            todo.update(imgs)
    out = repo / '_upload'
    if out.exists():
        for p in out.iterdir(): p.unlink()
    if not todo:
        print('올릴 그림 없음'); return
    out.mkdir(exist_ok=True)
    packs, cur, size = [], {}, 0
    for path, b64 in sorted(todo.items()):
        if cur and size + len(b64) * 3 // 4 > PACK_MAX:
            packs.append(cur); cur, size = {}, 0
        cur[path] = b64; size += len(b64) * 3 // 4
    packs.append(cur)
    aes = AESGCM(key)
    for i, p in enumerate(packs):
        iv = os.urandom(12)
        (out / f'pack_{i:03d}.bin').write_bytes(iv + aes.encrypt(iv, gzip.compress(json.dumps(p).encode()), None))
    (out / 'manifest.json').write_text(json.dumps({'packs': [{'file': f'pack_{i:03d}.bin', 'n': len(p)} for i, p in enumerate(packs)]}))
    print(f'그림 {len(todo)}장 → 묶음 {len(packs)}개 ({out})')

main()
