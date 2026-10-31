# 문항·해설 그림 자동 올리기

선생님 화면의 '문항 이미지 올리기'(img_q*.js, img_s*.js 직접 고르기)를 대신한다. 과목 상관없이 img_*.js면 된다.

1. 저장소에 이미 있는 그림의 회차별 요약을 SQL로 받아 `있는그림.json`으로 저장 (pack.py 머리말의 SQL).
2. 키 받기: `select public.img_pack_key()` (service_role 전용이라 Supabase MCP execute_sql로) → 환경변수 `IMG_PACK_KEY`.
3. `python3 pack.py <저장소 폴더> 있는그림.json <img_*.js ...>` → 새 회차·바뀐 회차 그림만 `_upload/`에 암호화 묶음으로.
4. `_upload/`를 커밋해 GitHub에 올리고, 그 커밋 sha로 함수 호출 (SQL, 묶음 수만큼):
   `select net.http_post('https://wlshvleklbctladudjkd.supabase.co/functions/v1/img-ingest', jsonb_build_object('sha','<sha>','pack',i), timeout_milliseconds => 150000) from generate_series(0, <묶음수-1>) i;`
5. 결과 확인: `select status_code, content from net._http_response order by id desc limit <묶음수>;` 와 `public.img_ingest_log`.
   다시 1번 SQL을 돌려 pack.py가 '올릴 그림 없음'이면 끝. 그 뒤 `_upload/`는 지워서 커밋한다.

묶음은 AES-GCM으로 암호화돼 공개 저장소에 있어도 그림을 볼 수 없다. 키는 Supabase Vault('img_pack_key')에만 있다.
함수 원본: supabase/functions/img-ingest/index.ts (verify_jwt 끔, 키가 맞는 묶음만 풀림).
