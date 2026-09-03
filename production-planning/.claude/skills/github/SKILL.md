---
name: github
description: 변경분을 스테이징하고 한국어 커밋 메시지로 origin main에 커밋·푸시한다. 커밋 전 익명화 검사와 비밀값(토큰) 유출 검사를 통과해야만 커밋한다. "커밋해줘", "깃허브 올려줘", #N 단계 완료·E2E 성공 시점에 사용.
---

# GitHub 커밋·푸시

한 단계(#N)가 끝나거나 E2E가 성공한 시점에 커밋한다. 새벽 작업의 롤백 안전망이며, 저장소가 나중에 public으로 바뀌므로 **비밀값이 한 번도 커밋되지 않아야** 한다.

## 절차
1. `git status --short` + `git diff --stat` 로 변경 확인. 변경이 없으면 커밋하지 않고 보고.
2. **익명화·비밀값 검사**: `python .claude/skills/anon-check/anon_check.py` — 종료 코드 2면 **커밋 중단**, 위반을 먼저 고친다.
3. `.gitignore` 확인: `.clasp.json`, `*.secret*`, `node_modules/` 가 추적되지 않는지. 실수로 추적되면 `git rm --cached` 후 알린다.
4. `git add -A`
5. 한국어 커밋 메시지 — 첫 줄은 "#N 단계명: 한 줄 요약" 또는 "#N 단계명 E2E 성공: …" 형식. 끝에 다음 줄:
   `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`
6. `git commit` → `git push origin main`
7. 커밋 해시·메시지·푸시 결과 보고. 푸시 인증 실패 시 `gh auth status` 확인 후 사용자에게 보고(토큰 임의 생성·우회 금지).

## 주의
- 미완료 중간 상태에서는 커밋하지 않는다.
- `--no-verify`, `--force`, `reset --hard` 는 사용자가 명시 요청할 때만.
- 원격이 아직 없으면(`git remote -v` 비어 있음) 커밋만 하고 "원격 미설정"을 보고한다.
