# RPB 분석표 자동 업데이트 봇

RAW DATA xlsx를 업로드하면 RPB 분석표를 자동으로 채워주는 웹앱입니다.

---

## 배포 순서

### 1단계 — Supabase 설정

1. [supabase.com](https://supabase.com) 에서 새 프로젝트 생성
2. **Storage** → 버킷 2개 생성:
   - `templates` — Public 체크 ✓
   - `results`   — Public 체크 ✓
3. `templates` 버킷에 **RPB_template.xlsx** 파일 업로드
   - 파일명은 반드시 `RPB_template.xlsx` 로 업로드
4. **Table Editor** → 새 테이블 생성:

```sql
create table rpb_history (
  id          bigserial primary key,
  filename    text,
  label       text,
  url         text,
  total_teu   int,
  created_at  timestamptz default now()
);
```

5. **Project Settings → API** 에서 복사:
   - `Project URL` → `SUPABASE_URL`
   - `service_role` 키 → `SUPABASE_SERVICE_KEY`

---

### 2단계 — GitHub 업로드

```bash
cd rpb-bot
git init
git add .
git commit -m "init"
git remote add origin https://github.com/본인계정/rpb-bot.git
git push -u origin main
```

---

### 3단계 — Vercel 배포

1. [vercel.com](https://vercel.com) → **Add New Project** → GitHub 저장소 선택
2. **Environment Variables** 탭에서 추가:

| Key | Value |
|-----|-------|
| `SUPABASE_URL` | `https://xxxx.supabase.co` |
| `SUPABASE_SERVICE_KEY` | `eyJ...` (service_role 키) |

3. **Deploy** 클릭 → 완료

---

## 사용법

1. 웹앱 URL 접속 (팀원 전체 공유 — 로그인 불필요)
2. RAW DATA xlsx 파일 업로드
3. 기간 레이블 입력 (예: `2026-06`)
4. **분석표 생성 →** 클릭
5. 자동으로 xlsx 다운로드 + Supabase에 히스토리 저장

---

## 양식(템플릿) 변경 방법

Supabase Storage → `templates` 버킷 → `RPB_template.xlsx` 파일 교체  
(코드 수정 없이 템플릿만 바꾸면 바로 반영됩니다)

---

## 집계 로직 요약

| POL | POD 필터 |
|-----|---------|
| VNHPH / CNSHK / CNXMN / VNSGN / TH | 전체 KR |
| JP (JPYOK/JPTYO/JPNGO/JPSMZ/JPOSA/JPUKB/JPHKT/JPMOJ) | 전체 KR |
| CNSHA / CNNGB | KRINC, KRKUV 별도 / KRPUS, KRKAN, KRUSN 별도 |
| CNTXG | KRINC/KRPUS/KRKAN/KRUSN 통합 / KRPTK 별도 |
| CNZJG | KRPUS, KRKAN, KRUSN 만 |
| CNNTG / CNTAG / CNZJG | KRINC, KRPTK 만 |
| RU (RUVVO/RUVFP) | 전체 KR |

- TEU: 각 그룹 합산 (반올림)
- RPB: S.TTL1 합계 ÷ TEU
- PP/CLT 구분: P/C 컬럼 값 `P` → PP, 나머지 → CLT
- ALL 행: 템플릿 기존 함수(`=AK5+AK6`) 유지
