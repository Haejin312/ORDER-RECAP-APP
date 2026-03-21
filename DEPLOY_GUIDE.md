# Order Recap Maker — 배포 가이드

## 폴더 구조
```
order-recap-app/
├── backend/
│   ├── app.py
│   ├── requirements.txt
│   ├── render.yaml
│   └── template.xlsx        ← 반드시 추가 필요!
└── frontend/
    └── index.html
```

---

## Step 1. GitHub 저장소 만들기

1. github.com 접속 → **New repository**
2. Repository name: `order-recap-app`
3. **Private** 선택 (API 키 보안)
4. **Create repository**

---

## Step 2. 파일 업로드

### 2-1. backend 폴더
- `app.py`, `requirements.txt`, `render.yaml` 업로드
- **중요**: `template.xlsx` (현재 사용 중인 Order Recap 양식 파일) 도 함께 업로드

### 2-2. frontend 폴더
- `index.html` 업로드

---

## Step 3. Render.com 백엔드 배포

1. **render.com** 가입 (GitHub 계정으로 로그인)
2. **New → Web Service**
3. GitHub 저장소 `order-recap-app` 연결
4. 설정:
   - **Root Directory**: `backend`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. **Deploy** 클릭
6. 배포 완료 후 URL 복사 (예: `https://order-recap-api.onrender.com`)

---

## Step 4. GitHub Pages 프론트엔드 배포

1. 저장소 → **Settings → Pages**
2. Source: `Deploy from a branch`
3. Branch: `main` / Folder: `/frontend`
4. Save → URL 확인 (예: `haejin312.github.io/order-recap-app`)

---

## Step 5. 앱 설정

1. 브라우저에서 앱 열기
2. **Settings** 탭:
   - API Key 입력 (`sk-ant-...`)
   - Backend URL 입력 (Render URL)
   - 연결 테스트 클릭
3. 설정 저장

---

## Step 6. 바탕화면 바로가기 만들기 (Windows)

1. 바탕화면 우클릭 → **새로 만들기 → 바로 가기**
2. 위치: `https://haejin312.github.io/order-recap-app`
3. 이름: `Order Recap Maker`
4. 마우스 우클릭 → 속성 → 아이콘 변경 (원하는 아이콘 선택)

---

## 주의사항

- Render.com 무료 플랜은 15분 비활동 시 슬립 상태 진입 (첫 요청 시 30초 대기)
- 유료 플랜($7/월) 사용 시 항상 활성 상태 유지
- API 키는 앱에만 입력, 절대 GitHub에 업로드 금지
