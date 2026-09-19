# broadcast-ppt

현장 예배용 PowerPoint 원본을 방송용(크로마키) PPT로 바꿔 주는 프로그램입니다. 방송 템플릿의 슬라이드를 복제해서 글자만 바꿔 넣습니다.

## 어떻게 동작하나

- 파일 작업은 전부 파이썬이 합니다. AI는 `GEMINI_API_KEY` 환경변수가 있을 때만, 규칙으로 종류를 못 정한 슬라이드에 한해 Gemini에게 종류만 묻습니다. 키가 없으면 AI에 아무것도 보내지 않습니다.
- 원본 슬라이드마다 규칙으로 종류를 판정합니다. 예배로의 부름, 기도, 성경봉독, 설교, 묵상, 교리문답(제N문), 교회소식, 성경 본문, 핵심 문장 같은 종류입니다.
- 템플릿 슬라이드를 복제하고 글자 상자 안의 글자만 바꿉니다. 글꼴과 색은 템플릿 그대로입니다.
- 긴 글은 줄 수에 맞는 높이의 템플릿으로 나눠 여러 장에 담습니다.
- 성경 본문과 설교 인용은 두 번 만듭니다. 원본 슬라이드를 화면 전체용으로 한 장 넣고, 그 뒤에 녹색 하단 자막을 넣습니다.
- 교리문답(제N문) 슬라이드도 원본 한 장을 넣고, 문답을 하단 자막 한 장에, 함께 적힌 말씀을 참조 자막 한 장에 따로 넣습니다.
- 사진, 영상, 찬송가 악보 슬라이드는 그대로 복사하고 검수 표시를 붙입니다.

## 처음 설치

```bash
uv sync
uv run pytest
```

템플릿 원본은 `local/templates/24.상반기.pptx`에 둡니다. 위치는 `config/scope.yaml`에 적혀 있습니다. 명령은 어느 폴더에서 실행해도 저장소 기준 경로를 씁니다.

Windows PC에서는 아래 [Windows에서 실행](#windows에서-실행)을 따르세요.

## Windows에서 실행

Windows PC에서도 Mac과 같은 명령으로 변환할 수 있습니다. 명령은 모두 **PowerShell** 창에 입력합니다(시작 메뉴에서 "PowerShell" 검색). 64비트(x64) PC에서만 됩니다.

### 처음 한 번

1. **폴더 위치**: `C:\broadcast-ppt`처럼 짧은 경로에 둡니다. 바탕 화면이나 문서 폴더처럼 OneDrive가 동기화하는 곳은 피하세요.
2. **코드 받기**: 비공개 저장소라서 먼저 브라우저에서 GitHub에 로그인해야 합니다. 저장소 페이지의 **Code → Download ZIP**으로 받아 압축을 풀고 폴더 이름을 `C:\broadcast-ppt`로 맞춥니다. git이 있으면 이렇게 받아도 됩니다.
   ```powershell
   git clone https://github.com/diskeyd/broadcast-ppt.git C:\broadcast-ppt
   ```
3. **uv 설치**: `uv --version`이 안 되면 아래 한 줄로 설치합니다. 설치가 끝나면 **PowerShell 창을 닫고 다시 엽니다.** 새 창에서만 `uv`를 찾습니다.
   ```powershell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```
4. **설치**: 처음에는 파이썬까지 내려받아서 몇 분 걸립니다. 마지막 줄은 템플릿과 원본을 둘 폴더를 만듭니다.
   ```powershell
   cd C:\broadcast-ppt
   uv sync
   mkdir local\templates, local\input
   ```
5. **템플릿 복사**: Mac의 `local/templates/24.상반기.pptx`(또는 Drive에 있는 같은 원본)를 **그대로** `C:\broadcast-ppt\local\templates\24.상반기.pptx`에 복사합니다. PowerPoint로 열어서 다시 저장하면 안 됩니다. 복사한 뒤 원본과 같은지 확인하세요.
   ```powershell
   Get-FileHash local\templates\24.상반기.pptx -Algorithm SHA256
   (Get-Item local\templates\24.상반기.pptx).Length
   ```
   첫 명령의 Hash 값이 `config/templates.yaml`의 `sha256` 값과 같아야 합니다(대소문자는 달라도 됩니다). 두 번째 명령은 `28279381`이 나와야 합니다. 다르면 다시 복사하세요. 변환할 때 템플릿이 없다고 나오면 `--template <경로>`로 템플릿 파일을 직접 지정합니다.

### 매주 실행

1. 원본 PPT를 `C:\broadcast-ppt\local\input\` 아래(예: `local\input\미편집\`)에 넣습니다. Mac의 `local\work` 폴더는 복사하지 마세요. PC에서 새로 만듭니다.
2. PowerShell에서 실행합니다. 한글 파일 이름은 직접 치지 말고 `local\input\`까지 친 다음 **Tab** 키로 채웁니다.
   ```powershell
   cd C:\broadcast-ppt
   uv run broadcast-ppt run "local\input\미편집\주일오전예배ppt(YYMMDD).pptx"
   ```
3. 결과는 Mac과 같이 `local\work\<원본 파일 이름>\broadcast.pptx`에 생깁니다. 탐색기에서 열면 됩니다.

주의할 점:

- 수정 파일을 고친 뒤 다시 실행할 때는 **PowerPoint에서 `broadcast.pptx`를 먼저 닫습니다.** 열려 있으면 덮어쓸 수 없다는 안내가 나옵니다.
- `overrides.yaml`은 메모장에서 **UTF-8**로 저장합니다(다른 이름으로 저장 → 인코딩 UTF-8).
- 미리 열어 둔 PowerShell 창에서 실행하세요. 오류가 나도 글이 창에 남습니다. `> log.txt`처럼 출력을 파일로 보내거나 Git Bash에서 실행하지 마세요.
- Gemini 키(`GEMINI_API_KEY`)가 없으면 규칙만으로 변환합니다. Mac과 PC의 결과가 같으려면 키 설정을 두 곳에서 똑같이 둡니다(둘 다 있거나 둘 다 없거나).

## 매주 사용법

### 1. 변환

```bash
uv run broadcast-ppt run "local/input/미편집/주일오전예배ppt(YYMMDD).pptx"
```

한 번에 계획, 생성, 검증까지 합니다. 결과는 `local/work/<원본 파일 이름>/` 폴더에 생깁니다.

| 파일 | 내용 |
|---|---|
| `broadcast.pptx` | 방송용 PPT |
| `plan.json` | 원본 슬라이드마다 어떤 템플릿으로 바꿨는지 |
| `validation.json` | 검증 결과 |

### 2. 검증 결과 보기

실행이 끝나면 `critical`, `warning`, `info` 개수가 나옵니다.

- **critical**이 1개라도 있으면 방송에 쓰면 안 됩니다. 먼저 고쳐야 합니다.
- **warning**은 자막 줄 수가 빠듯한 장입니다. 해당 슬라이드를 눈으로 확인하세요.
- **info**는 검수 표시입니다. 아래 표를 보고 해당 장을 확인하세요.

### 3. 눈으로 확인

`broadcast.pptx`를 Mac에서는 Keynote로, 송출 PC에서는 PowerPoint로 엽니다. 최소한 아래 장은 확인하세요.

- 카드: 예배로의 부름, 기도, 성경봉독, 설교
- 성경 자막 한 장
- 묵상
- 교회소식

Mac에서는 명령으로 Keynote 그림을 뽑아 볼 수도 있습니다.

```bash
uv run python scripts/keynote_check.py render "local/work/<원본 파일 이름>/broadcast.pptx" --slides 2,4,12
```

그림은 `local/work/keynote/check/png/`에 저장됩니다. Keynote는 보안 제한 때문에 자기 iCloud 문서 폴더의 파일만 열 수 있습니다. 그래서 잠깐 그 폴더에 복사했다가 끝나면 지웁니다. 이름이 들어간 슬라이드는 가능하면 넣지 마세요.

#### 송출 PC에서 처음 열 때

방송본을 PowerPoint로 여는 것은 처음입니다. 첫 주에는 방송 전에 아래를 확인하세요.

- 파일을 열 때 "복구" 창이 뜨는지. 뜨면 그 문구를 적어 둡니다.
- **홈 → 바꾸기 → 글꼴 바꾸기**에서 쓰인 글꼴 목록을 보고 PC에 없는 글꼴이 있는지(나눔 글꼴, Script MT Bold, 맑은 고딕).
- 카드 4종(예배로의 부름, 기도, 성경봉독, 설교)에서 영문 줄과 한글 줄이 겹치거나 너무 붙지 않았는지.
- 문제가 있으면 그 주는 원본 PPT와 평소 하던 수작업 방식으로 방송합니다.

### 4. 단계별로 나눠서 실행

중간에 계획표를 보고 싶을 때 씁니다.

```bash
uv run broadcast-ppt plan  "<원본.pptx>"                                  # 원본 슬라이드 → 종류 → 템플릿 표 출력
uv run broadcast-ppt build "local/work/<원본 파일 이름>/plan.json"           # 생성 + 검증
uv run broadcast-ppt validate "local/work/<원본 파일 이름>/broadcast.pptx" --plan "local/work/<원본 파일 이름>/plan.json"
```

### 5. 테스트 피드백 남기기

주일에 써 본 결과는 `local/work/<원본 파일 이름>/feedback.md`에 적습니다. `local/`은 GitHub에 올라가지 않습니다. 이름, 직분, 교인 정보는 적지 말고 장 번호와 증상만 적습니다. 스크린샷은 이름이 없는 장만 남깁니다.

```markdown
# 2026-09-20 주일 테스트
- 확인한 곳: 송출 PC PowerPoint / Mac Keynote
- 검증: critical 0 · warning 2 · info 9
- 수정 파일로 해결: 14번 quote, 13번 ref (없으면 "없음")

| 방송본 장 | 원본 장 | 종류 | 증상 | 수정 파일로 해결? |
|---|---|---|---|---|
| 30 | 15 | 교리문답 자막 | 답 글자가 띠 밖으로 넘침 | 아니요 |
| 41 | 22 | 새 양식 | 원본 그대로 복사됨, 자막이 필요함 | 예 (caption) |
```

- **원본 장**은 `plan` 표의 `src` 열 번호입니다.
- 증상은 이런 식으로 적습니다: 글자 넘침, 띠 아래 여백, 색·크기, 영문·한글 줄 겹침, 종류가 틀림, 참조가 틀림·빠짐, 장이 빠짐·중복.
- 수정 파일(`overrides.yaml`)로 해결되는 것은 그 주에 바로 고치고 표에 "예"로 남깁니다. **수정 파일로 안 되는 것만** 모아서 나중에 한 번에 코드 수정을 요청합니다.
- 처음 몇 주는 특히 송출 PC에서 카드의 영문·한글 줄 간격, 교리문답 자막(주황 질문·28pt 답), 참조 박스를 숨긴 성경 자막, 줄여진 3줄 자막 띠를 봐 주세요.

## 고치고 싶을 때

`local/work/<원본 파일 이름>/overrides.yaml` 파일을 만들고 `plan --overrides` 다음 `build`를 다시 실행합니다.

```yaml
slides:
  17: {skip: true}      # 원본 17번 슬라이드는 방송에 안 씀
  22: {op: copy}        # 원본 22번은 바꾸지 말고 그대로 복사
  13: {ref: "요일 1:1-4"}  # 원본에 참조가 없는 성경 본문: 자막마다 이 참조를 넣음
  14: {op: quote}       # 노란 인용 카드로 만들되 '묵 상 Meditate' 글자는 비움
  30: {op: caption}     # 슬라이드 글자를 하단 자막으로 (원본 한 장 + 자막)
  31: {op: scripture, ref: "시 23:1-6"}  # 슬라이드 글자를 성경 자막으로. ref는 없어도 됨(참조 칸 비움)
```

새 양식이 나오면 코드를 고치지 않습니다. 처음에는 원본 그대로 복사되고 `review:unclassified`가 붙으니, 방송본에서 바꾸고 싶은 슬라이드만 위 파일에 한 줄 적습니다. `op`에는 `copy`, `quote`, `caption`, `scripture`만 쓸 수 있고, 다른 값을 쓰면 `plan`이 멈추고 알려 줍니다.

`ref`는 참조 칸이 있는 슬라이드(성경 본문·성경봉독·설교)나 `op: scripture`와 함께 쓸 수 있고, `copy`·`quote`·`caption`과는 같이 쓸 수 없습니다. 참조 형식이 틀리거나, 글자가 없는 슬라이드에 `quote`·`caption`·`scripture`를 쓰면 `plan`이 멈추고 알려 줍니다. `op: quote`는 슬라이드의 글자를 모두 한 줄 인용문으로 넣습니다.

## 검수 표시

| 표시 | 뜻 |
|---|---|
| `review:inherited_reference` | 절 번호로 시작하는 본문에 참조가 없어서, 앞 참조나 성경봉독 참조를 이어받음 |
| `review:reference_unknown` | 절 번호가 알고 있는 어느 참조에도 안 맞음. 자막의 참조 박스를 숨김 |
| `review:no_reference_tag` | 자막 한 장에 넣을 참조가 없어서 참조 박스를 숨김. 틀린 참조는 넣지 않음 |
| `review:no_reference` | 성경봉독이나 설교 카드에 참조가 없음 |
| `review:no_page_number` | 성경봉독 카드에 쪽수가 없음. 쪽수 줄을 비워 둠 |
| `review:no_sermon_title` | 설교 카드에 따옴표 제목이 없음 |
| `review:no_quote` | 인용문 없는 묵상 슬라이드라서 그대로 복사함 |
| `review:quote_too_long` | 묵상 인용이 너무 길어서 최소 글자 크기로 줄임 |
| `review:sentence_over_picture` | 사진 위의 짧은 문장을 자막으로 만듦 |
| `review:empty_item_body` | 교회소식 항목에 제목만 있고 내용이 없음 |
| `review:empty_slide` | 글자도 사진도 없는 슬라이드 |
| `review:graphic` | 도표나 질문지 슬라이드라서 그대로 복사함 |
| `review:unclassified` | 어느 규칙에도 안 맞아서 그대로 복사함 |
| `review:ai_classified` | 규칙으로 못 정한 슬라이드를 Gemini가 분류함(확신도 0.8 이상). 종류가 맞는지 확인. 글자는 원본 그대로 |
| `deferred:hymn` | 찬송가 악보·가사 그림을 그대로 복사함. 크로마 처리는 나중에 |

## 명령 옵션

| 명령 | 옵션 | 효과 |
|---|---|---|
| `plan`, `run` | `--overrides 파일` | 위의 수정 파일 적용 |
| `plan`, `run` | `--slide-size template` 또는 `source` | 출력 크기. 기본은 템플릿 크기 13.33인치이고 원본 슬라이드는 4/3배로 키움 |
| `plan`, `run` | `--no-scripture-both` | 성경 본문을 화면 전체용 없이 자막만 만듦 |
| `plan`, `run` | `--sermon-material lower-third` 또는 `passthrough` | 설교 문장을 자막으로 만들지(기본), 그대로 복사할지 |
| `build`, `run` | `--template 파일` | 다른 템플릿으로 생성 |
| `build`, `run` | `--force` | 계획 뒤에 템플릿이나 원본이 바뀌어도 생성 |
| `build` | `--no-validate` | 생성 후 검증을 건너뜀 |
| `analyze` | `--workers N`, `--force`, `--scope 파일` | 병렬 작업 수, 이미 분석한 파일도 다시 분석, 요약 범위 |
| `plan`, `run` | 환경변수 `GEMINI_API_KEY` (선택 `GEMINI_MODEL`) | 키가 있으면 분류 못 한 슬라이드를 Gemini에 물어봄. 기본 모델 `gemini-3.5-flash`. 확신도 0.8 미만이면 그대로 복사하고 `plan.json`의 `ai`에 제안만 남김 |

## 설정 파일

- `config/scope.yaml`: 변환 대상 예배와 템플릿 위치.
- `config/roles.yaml`: 템플릿 슬라이드마다 어떤 역할이고 어느 글자 칸에 무엇을 넣을지 손으로 적은 표. 배치 규칙도 여기 있습니다.
  - `slots`, `mode`, `single_segment`: 어느 서식 구간에 어떤 값을 넣을지. `mode`는 `fill`(반드시 채움), `optional`(값이 있을 때만), `fixed`(안 건드림).
  - `clamp_to`: 자막 본문 상자를 막대 안으로 좁힘.
  - `shrink`: 줄이 남으면 막대 위쪽을 줄임.
  - `nudge_y_pt`: 상자를 아래로 내림. 자막은 3pt.
  - `fit_one_line`: 묵상 인용이 한 줄에 들어가게 글자 크기를 줄임.
  - `drop_empty_paragraphs`: 빈 줄을 지워서 가운데 정렬이 맞게 함.
  - `hide_when_empty`: 넣을 값이 없으면 도형을 숨김(도형은 남기고 채우기·선·그림자만 없앰). 참조 박스에 씀.
  - `stack`, `card_gap_pt`, `card_max_bottom_pt`: 카드의 노란 줄과 흰 줄을 아래에서부터 쌓음. 줄 사이 간격과 맨 아래 글자선을 정합니다. 줄별 `glyph` 값은 Keynote에서 잰 글자 위치입니다. 템플릿을 고치면 `scripts/keynote_check.py calibrate-cards`로 다시 재야 합니다.
- `config/templates.yaml`: 템플릿과 `roles.yaml`로 자동 생성하는 파일입니다. 직접 고치지 마세요.
- `config/widths.yaml`: 글자 폭 보정값. PowerPoint가 저장한 자동 맞춤 상자 크기로 계산했습니다.

## 관리용 명령

```bash
uv run broadcast-ppt catalog                                          # roles.yaml이나 템플릿을 고친 뒤 templates.yaml 다시 만들기
uv run broadcast-ppt calibrate <방송본 PPT 또는 폴더>                    # 글자 폭 보정값 다시 계산
uv run broadcast-ppt analyze local/input --out local/analysis         # 샘플 PPT 구조 분석
uv run python scripts/pull_drive.py <Drive 폴더 ID>                     # Google Drive 폴더를 local/input으로 내려받기
uv run python scripts/keynote_check.py calibrate-cards                # 카드 줄 위치를 Keynote로 다시 재기(가짜 이름 사용)
uv run python scripts/keynote_check.py bands <그림.png> ...             # 그림에서 글자가 그려진 세로 위치 재기
```

## 아직 안 되는 것

- 찬송가 악보·가사 그림은 아직 크로마 처리를 안 합니다.
- 담임목사 위임식 같은 특별 예배는 주일 규칙으로 바뀌어서 따로 규칙이 필요합니다.
- Mac에서만 확인했습니다. 이 Mac에는 Script MT Bold와 맑은 고딕이 없어서 Keynote가 다른 글꼴로 바꿔 보여 줍니다. 카드 간격은 Keynote 기준으로 맞췄습니다. 송출 PC에 진짜 Script MT Bold가 있으면 영문 부제가 한글 줄과 가까워지거나 겹칠 수 있으니 카드 한 장을 확인하세요.
- 템플릿 성경 본문의 강조 색은 재현하지 않습니다. 자막 본문은 한 가지 색입니다.

## 데이터 보관 원칙

PPT에는 이름, 직분, 교회 정보가 들어 있습니다. 원본, 템플릿, 분석 결과, 계획, 수정 파일, 검증 결과, 확인용 그림, 방송본은 모두 `local/`에 두고 GitHub에 올리지 않습니다. 저장소에는 코드, 일반 설정, 문서만 올립니다. `config/templates.yaml`에는 위치, 서식, 글자 수, 템플릿 샘플 글자의 SHA-1 요약값만 저장하고 글자 자체는 저장하지 않습니다. 명령 출력과 보고서에는 글자 수 같은 통계만 나오고 슬라이드 글자는 나오지 않습니다.

`GEMINI_API_KEY`가 있으면 규칙으로 분류 못 한 슬라이드의 글자 줄과 도형 개수만 Google Gemini로 나갑니다. 직분 앞 이름(`홍길동 집사`, `홍길동 장로님` → `[이름] 집사`, `[이름] 장로님`)은 가리고 보내지만, 직분 없이 쓴 이름이나 명단은 가려지지 않습니다. 슬라이드 노트, 표, 파일 자체와 수정 파일에 적은 슬라이드는 보내지 않습니다. 보내고 싶지 않은 주에는 키를 빼고 실행하세요.

## 폴더를 지우기 전에 (보관과 복구)

GitHub에는 코드, 설정, 문서만 있습니다. 아래 파일은 git에 없어서 프로젝트 폴더를 지우면 같이 사라집니다. 다시 쓸 파일은 교회 자료를 두는 비공개 위치(예: Google Drive)에 따로 보관하세요. 지우기 전에 커밋이 GitHub에 push됐는지(`git status`에 `ahead`가 없는지)도 확인하세요.

| 경로 | 내용 | 다시 쓸 때 |
|---|---|---|
| `local/templates/24.상반기.pptx` | 방송 템플릿 정본 | 변환에 필수 |
| `local/work/<원본 이름>/overrides.yaml` | 주마다 적은 수정 파일 | 그 주를 다시 만들 때만 |
| `local/pii-patterns.txt`, `.git/hooks/pre-commit` | 커밋 전에 이름·개인정보를 막는 검사 | 코드를 고칠 때 필수 |
| `local/input/`, `local/analysis/`, `local/work/`, `local/requirements/` | 원본 PPT(Drive에 원본 있음), 분석, 결과물, 요구사항 원문(볼트에 사본) | 보통 필요 없음 |

복구 순서: `git clone` → `uv sync` → 템플릿을 `local/templates/`에 복사 → `uv run pytest` → 코드를 고칠 계획이면 표의 검사 파일도 제자리에 복사.
