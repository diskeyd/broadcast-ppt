# broadcast-ppt

현장 예배용 PowerPoint 원본을 방송용(크로마키) PPT로 바꾸는 명령줄 프로그램입니다. 방송 템플릿의 슬라이드를 복제한 뒤 글자만 원본 내용으로 바꿔 넣습니다. 글꼴과 색은 템플릿을 따르고, 글자 길이에 맞춰 상자 폭·위치와 자막 띠 높이만 조금 조정합니다.

- 실행 환경: Mac, Windows(x64). 파이썬 3.12를 쓰지만 따로 설치하지 않습니다. `uv`가 알아서 내려받습니다.
- 입력: 원본 PowerPoint 파일(**.pptx**) 한 개. 옛 형식 `.ppt`는 PowerPoint에서 `.pptx`로 다시 저장해서 넣습니다.
- 출력: 방송용 PPT, 변환 계획(`plan.json`), 검증 결과(`validation.json`).
- 슬라이드 종류는 규칙으로 판정합니다. AI(Gemini)는 선택 사항이고 기본은 꺼져 있습니다.

용어: **템플릿**은 방송용 양식 슬라이드가 모두 들어 있는 `24.상반기.pptx`입니다. **크로마키**는 녹색 부분을 방송 장비가 투명하게 합성하는 방식입니다. **카드**는 예배로의 부름·기도·성경봉독·설교처럼 제목을 크게 보여 주는 장, **하단 자막(띠)**은 화면 아래 녹색 띠에 글자를 넣은 장입니다. **원본 장 번호**와 **방송본 장 번호**는 서로 다릅니다(원본 한 장이 방송본 여러 장이 되기도 함).

> **처음 쓰는 Windows PC라면** [Windows 설치](#windows-설치-처음-한-번-약-10분) 1~5단계 → [매주 사용법](#매주-사용법) 1~3단계만 하면 됩니다. 막히면 [문제 해결](#문제-해결)을 봅니다.

**목차**: [한눈에 보기](#한눈에-보기) · [설치](#설치) · [매주 사용법](#매주-사용법) · [고치고 싶을 때](#고치고-싶을-때-수정-파일) · [문제 해결](#문제-해결) · [명령과 옵션](#명령과-옵션) · [동작 원리](#동작-원리) · [설정 파일](#설정-파일) · [아직 안 되는 것](#아직-안-되는-것) · [데이터 보관](#데이터-보관-원칙) · [백업과 복구](#백업과-복구)

## 한눈에 보기

```
원본.pptx ─▶ ① 분석 ─▶ ② 분류 ─▶ ③ 계획 ─▶ ④ 생성 ─▶ ⑤ 검증
             슬라이드   종류 판정   plan.json   broadcast.pptx   validation.json
             구조 읽기  (규칙 먼저)  어떤 템플릿에  템플릿 복제 +    빈 칸·남은 견본
                                    무엇을 넣을지  글자 바꾸기      글자 등 점검
```

`run` 명령 하나가 ①~⑤를 모두 합니다. 폴더는 이렇게 씁니다. `local/`은 교회 자료라서 git에 올라가지 않습니다.

```
broadcast-ppt/
├── src/, tests/, config/         코드, 테스트, 설정 (git에 있음)
└── local/                        git에 없음. 직접 만들고 채움
    ├── templates/24.상반기.pptx   방송 템플릿 원본 (필수, 직접 복사)
    ├── input/                    원본 PPT를 넣는 곳
    └── work/<원본 파일 이름>/      결과가 생기는 곳 (.pptx를 뺀 이름)
        ├── broadcast.pptx        방송용 PPT
        ├── plan.json             원본 장마다 어떤 템플릿으로 바꿨는지 (슬라이드 글자 포함)
        ├── validation.json       검증 결과
        └── overrides.yaml        (선택) 그 주의 수정 파일
```

## 설치

### 준비물

| 항목 | 용도 | Windows | Mac |
|---|---|---|---|
| **uv** | 파이썬과 라이브러리를 설치하고 실행하는 도구 | 필수 (아래 2단계) | 필수 (`brew install uv`) |
| 파이썬 3.12 | 프로그램 실행 | **따로 설치 안 함.** `uv sync`가 내려받음 | 같음 |
| **Git** | 코드 받기·업데이트 | 권장 (없으면 ZIP으로 받음) | 보통 이미 있음 |
| GitHub 권한 | 비공개 저장소라서 로그인 필요 | 필수 | 필수 |
| 템플릿 원본 | `24.상반기.pptx`. git에 없음 | 필수 (USB·Mac·Drive에서 복사) | 필수 |
| 인터넷 | 처음 설치 때 파이썬·라이브러리 내려받기(수백 MB) | 처음 한 번 | 처음 한 번 |
| PowerPoint와 글꼴 | 결과를 열어 볼 때만 필요. 변환에는 필요 없음 | 아래 [글꼴](#글꼴) 참고 | Keynote로 확인 |

### Windows 설치 (처음 한 번, 약 10분)

모든 명령은 **PowerShell**에 입력합니다. 시작 메뉴에서 "PowerShell"을 검색해 엽니다(관리자 권한으로 열 필요는 없음). 64비트(x64) PC만 됩니다.

- `winget`을 처음 쓰면 약관 동의를 묻습니다. `Y`를 누릅니다.
- Git 설치 중 "이 앱이 디바이스를 변경하도록 허용하시겠어요?" 창이 뜨면 **예**를 누릅니다. 관리자 암호를 모르면 Git은 건너뛰고 3단계에서 ZIP으로 받습니다.

**1. Git 설치** (ZIP으로 받을 거면 건너뜀)

```powershell
winget install --id Git.Git -e
```

**2. uv 설치**

```powershell
winget install --id astral-sh.uv -e
```

`winget`이 없는 PC라면 이 명령을 대신 씁니다.

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

설치가 끝나면 **PowerShell 창을 완전히 닫고**(Windows 터미널이면 탭이 아니라 창 전체) 새로 엽니다. 새 창에서 `uv --version`이 버전을 출력하면 됩니다. Git을 설치했다면 `git --version`도 확인합니다. 그래도 찾을 수 없다고 나오면 Windows에서 로그아웃했다가 다시 로그인합니다.

**3. 코드 받기.** 바탕 화면에 둡니다. 탐색기에서 바로 보여 파일 넣기가 편합니다. 단 바탕 화면이 OneDrive로 동기화되는 PC(탐색기 주소에 `OneDrive`가 들어감)라면 `C:\broadcast-ppt`처럼 동기화 밖 짧은 경로에 두고, 아래 명령의 경로도 그에 맞춰 바꿉니다. 교회 PPT가 들어가는 `local` 폴더가 클라우드로 올라가면 안 됩니다.

```powershell
git clone https://github.com/diskeyd/broadcast-ppt.git $HOME\Desktop\broadcast-ppt
```

처음이면 GitHub 로그인 창이 뜹니다. **Sign in with your browser**를 눌러 저장소 권한이 있는 계정으로 로그인합니다(2단계 인증을 쓰면 휴대폰이 필요합니다). 로그인 정보는 이 PC에 저장되어 다음 `git pull`부터는 묻지 않습니다.

Git이 없으면 브라우저에서 GitHub에 로그인한 뒤 저장소 페이지의 **Code → Download ZIP**으로 받습니다. 압축을 풀면 `broadcast-ppt-main` 폴더 안에 같은 이름의 폴더가 한 겹 더 있을 수 있습니다. `pyproject.toml`이 바로 보이는 폴더를 바탕 화면으로 옮기고 이름을 `broadcast-ppt`로 바꿉니다. `바탕 화면\broadcast-ppt\pyproject.toml`이 있으면 맞습니다. 여러 사람이 쓰는 PC라면 받은 뒤 브라우저에서 GitHub 로그아웃합니다.

**4. 파이썬과 라이브러리 설치.** 처음에는 파이썬 3.12까지 내려받아서 몇 분 걸립니다. 멈춘 것처럼 보여도 기다립니다.

```powershell
cd $HOME\Desktop\broadcast-ppt
uv sync
uv run broadcast-ppt --help
```

마지막 명령이 `usage: broadcast-ppt`로 시작하는 영어 안내를 출력하면 설치가 된 것입니다.

**5. 템플릿 넣기**

```powershell
New-Item -ItemType Directory -Force -Path local\templates, local\input
ii local\templates
```

두 번째 명령이 탐색기로 `local\templates` 폴더를 엽니다. USB(또는 Mac·Drive)에 있는 템플릿 원본을 이 폴더에 **그대로** 복사합니다.

- 파일 이름은 정확히 `24.상반기.pptx`여야 합니다. `24.상반기 (1).pptx`처럼 바뀌면 "템플릿이 없습니다"가 나옵니다.
- PowerPoint로 열어서 다시 저장하면 안 됩니다. 한 바이트라도 다르면 변환할 때 "템플릿이 원본과 다릅니다"가 나오고 멈춥니다.
- 미리 확인하려면 아래 명령의 `Hash`가 `EBC589F5`로 시작하는지 봅니다(전체 값은 `config\templates.yaml`의 `sha256`, 크기는 28,279,381바이트).
- 기준이 되는 판은 **Drive에 올려 둔 사본**입니다. 송출 PC에서 계속 쓰는 파일은 PowerPoint로 다시 저장되면서 값이 달라져 있을 수 있으니, 그 파일 말고 Drive 사본을 받아 넣습니다. 템플릿 자체를 고쳤다면 `catalog`로 `config\templates.yaml`을 다시 만들고 그 파일을 Drive에 올려 기준을 바꿉니다.

```powershell
Get-FileHash local\templates\*.pptx -Algorithm SHA256 | Format-List
```

**6. (선택) 동작 확인**

```powershell
uv run pytest
```

전부 `passed`면 이 PC에서 문제없이 돌아갑니다.

### Mac 설치

```bash
brew install uv                       # 또는 curl -LsSf https://astral.sh/uv/install.sh | sh
git clone https://github.com/diskeyd/broadcast-ppt.git
cd broadcast-ppt
uv sync
mkdir -p local/templates local/input  # 그다음 템플릿을 local/templates/에 복사
uv run pytest
```

### 업데이트

```powershell
cd $HOME\Desktop\broadcast-ppt
git pull
uv sync
```

ZIP으로 받았다면 새 ZIP을 받아 코드 폴더만 바꾸고 `local` 폴더는 그대로 옮깁니다. `.venv` 폴더는 복사하지 않습니다. `uv sync`가 다시 만듭니다.

### 글꼴

변환 자체에는 글꼴이 필요 없습니다. 결과를 PowerPoint로 열 PC(송출 PC)에는 템플릿이 쓰는 글꼴이 있어야 모양이 맞습니다.

- **나눔스퀘어**(OTF 포함), **나눔고딕**, **나눔스퀘어라운드**, **나눔손글씨 붓**(묵상 인용): Windows 기본 글꼴이 아닙니다. 네이버 나눔 글꼴(무료)에서 받아 설치합니다. 템플릿을 만든 송출 PC에는 보통 이미 있습니다.
- **Script MT Bold**: Microsoft Office와 함께 설치됩니다.

PowerPoint의 **홈 → 바꾸기 → 글꼴 바꾸기** 목록에서 PC에 없는 글꼴을 확인할 수 있습니다.

## 매주 사용법

모든 명령은 저장소 폴더(`$HOME\Desktop\broadcast-ppt`)에서 실행합니다. 새 PowerShell 창을 열었으면 먼저 `cd $HOME\Desktop\broadcast-ppt`를 합니다.

### 1. 원본 넣고 변환

USB의 원본 PPT를 `바탕 화면\broadcast-ppt\local\input\`에 복사합니다(`ii local\input`으로 폴더 열기). 하위 폴더 없이 바로 넣으면 입력이 쉽습니다.

```powershell
cd $HOME\Desktop\broadcast-ppt
uv run broadcast-ppt run "local\input\주일오전예배ppt(YYMMDD).pptx"
```

- 한글 파일 이름은 직접 치지 말고 `"local\input\`까지 친 다음 **Tab** 키로 채웁니다. 파일이 여러 개면 Tab을 여러 번 눌러 고릅니다.
- `YYMMDD`는 예배 날짜(예: `260920`)입니다. 파일 이름을 `0920.pptx`처럼 영문·숫자로 바꿔 넣어도 됩니다. 이름은 결과 폴더 이름에만 쓰입니다.
- Mac이 USB에 만든 `._`로 시작하는 파일은 고르지 않습니다(진짜 파일이 아님).

결과는 `local\work\<원본 파일 이름(.pptx 뺀 것)>\broadcast.pptx`에 생깁니다. 실행 끝의 `built … slides -> <경로>` 줄에 전체 경로가 나옵니다.

### 2. 검증 결과 보기

화면에는 이 순서로 나옵니다.

1. **plan 표**: 원본 장마다 한 줄. `src`=원본 장 번호, `kind`=판정한 종류, `rule`=쓴 규칙, `ops`=만든 장(`T23`=템플릿 23번 슬라이드, `copy`=원본 그대로), `[ ]`=검수 표시.
2. `built … slides -> <broadcast.pptx 경로>`
3. `validation: critical … · warning … · info …`와, critical·warning이 있으면 `critical  slide 12: …`처럼 한 줄씩.

| 등급 | 뜻 | 할 일 |
|---|---|---|
| **critical** | 방송에 쓰면 안 되는 문제(빈 칸, 남은 견본 글자, 끊긴 연결 등) | 먼저 고침. 아래 참고 |
| **warning** | 자막 줄 수·폭이 빠듯함 등 | 해당 장을 눈으로 확인 |
| **info** | 검수 표시 | 화면에는 종류별 개수만 나옴. 어느 장인지는 plan 표에서 `[review:…]`가 붙은 줄로 확인. 뜻은 [검수 표시](#검수-표시) 표 |

`slide 12`의 12는 **방송본** 장 번호입니다. [수정 파일](#고치고-싶을-때-수정-파일)에는 **원본** 장 번호를 써야 하므로 아래 명령으로 찾습니다(12 자리에 번호).

```powershell
(Get-Content "local\work\<원본 파일 이름>\plan.json" -Raw -Encoding UTF8 | ConvertFrom-Json).slides | Where-Object { $_.out -contains 12 } | Select-Object src, kind
```

critical이 있어도 `broadcast.pptx`는 만들어지지만 그대로 방송에 쓰지 않습니다. 빈 칸이나 견본 글자가 남은 장은 그 원본 장을 수정 파일에 `{op: copy}`로 적고 다시 실행하면 원본 그대로 들어갑니다. 시간이 없으면 그 주는 원본 PPT로 방송합니다.

### 3. 눈으로 확인

`ii local\work`로 탐색기를 열고 `<원본 파일 이름>` 폴더의 `broadcast.pptx`를 더블클릭해 PowerPoint로 엽니다. 최소한 아래 장은 봅니다.

- 카드: 예배로의 부름, 기도, 성경봉독, 설교
- 성경 자막 한 장, 묵상, 교회소식
- 교리문답 자막(주황 질문·28pt 답), 참조 박스를 숨긴 성경 자막, 줄여진 3줄 자막 띠

**송출 PC에서 처음 열 때** (첫 몇 주):

- 파일을 열 때 "복구" 창이 뜨는지. 뜨면 그 문구를 적어 둡니다.
- 없는 글꼴이 있는지([글꼴](#글꼴)).
- 카드 4종에서 영문 줄과 한글 줄이 겹치거나 너무 붙지 않았는지. 카드 간격은 Mac Keynote 기준으로 맞춰서 PowerPoint에서는 다를 수 있습니다.
- 문제가 있으면 그 주는 원본 PPT와 평소 하던 수작업으로 방송합니다.

### 4. 필요하면 고치고 다시 실행

[수정 파일](#고치고-싶을-때-수정-파일)을 적고 `--overrides`를 붙여 다시 실행합니다. **PowerPoint에서 `broadcast.pptx`를 먼저 닫습니다.** 열려 있으면 "파일을 열 수 없습니다 … PowerPoint에서 broadcast.pptx를 닫고 다시 실행하세요"가 나옵니다.

### 5. 테스트 결과 남기기

써 본 결과는 `local\work\<원본 파일 이름>\feedback.md`에 적습니다. `local`은 GitHub에 올라가지 않습니다. 이름, 직분, 교인 정보는 적지 말고 장 번호와 증상만 적습니다. 스크린샷은 이름이 없는 장만 남깁니다.

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

- 증상 예: 글자 넘침, 띠 아래 여백, 색·크기, 영문·한글 줄 겹침, 종류가 틀림, 참조가 틀림·빠짐, 장이 빠짐·중복.
- 수정 파일로 해결되는 것은 그 주에 바로 고치고 표에 "예"로 남깁니다. **수정 파일로 안 되는 것만** 모아서 나중에 한 번에 코드 수정을 요청합니다.

### Mac에서

같은 명령을 경로만 `/`로 써서 실행합니다(`uv run broadcast-ppt run "local/input/….pptx"`). 결과는 Keynote로 엽니다. Keynote 그림을 명령으로 뽑아 볼 수도 있습니다. 그림은 `local/work/keynote/check/png/`에 저장됩니다.

```bash
uv run python scripts/keynote_check.py render "local/work/<원본 파일 이름>/broadcast.pptx" --slides 2,4,12
```

## 고치고 싶을 때 (수정 파일)

새 양식이 나와도 코드를 고치지 않습니다. 결과 폴더에 `overrides.yaml`을 만들고 **원본** 장 번호(plan 표의 `src`, 방송본 번호가 아님)별로 한 줄씩 적은 뒤 `--overrides`를 붙여 다시 실행합니다. 결과 폴더는 첫 실행 뒤에 생깁니다.

파일은 메모장으로 만듭니다. 없는 파일이면 새로 만들지 묻는데 **예**를 누릅니다.

```powershell
notepad "local\work\주일오전예배ppt(YYMMDD)\overrides.yaml"
```

```yaml
slides:
  17: {skip: true}      # 원본 17번은 방송본에서 뺌
  22: {op: copy}        # 원본 22번은 바꾸지 말고 그대로 복사
  13: {ref: "요일 1:1-4"}  # 참조가 없는 성경 본문·성경봉독·설교: 이 참조를 넣음
  14: {op: quote}       # 노란 인용 카드로 만들되 '묵 상 Meditate' 글자는 비움
  30: {op: caption}     # 슬라이드 글자를 하단 자막으로 (원본 한 장 + 자막)
  31: {op: scripture, ref: "시 23:1-6"}  # 슬라이드 글자를 성경 자막으로. ref가 없으면 참조 박스를 숨기고 검수 표시
```

```powershell
uv run broadcast-ppt run "local\input\주일오전예배ppt(YYMMDD).pptx" --overrides "local\work\주일오전예배ppt(YYMMDD)\overrides.yaml"
```

- 다시 실행할 때마다 `--overrides`를 붙입니다. 빼면 수정 없이 다시 만들어집니다. **↑** 키로 앞 명령을 불러와 끝에 붙이면 한글 경로를 다시 칠 필요가 없습니다.
- 들여쓰기는 **공백 2칸**만 씁니다(Tab 키 금지). `17:`처럼 콜론 뒤에 공백이 한 칸 있어야 합니다.
- 메모장 '다른 이름으로 저장'을 쓸 때는 파일 형식을 **모든 파일(\*.\*)**로 바꿔야 `.txt`가 붙지 않습니다. 인코딩은 **UTF-8**(Windows 10 1903 이후 기본).
- `op`에는 `copy`, `quote`, `caption`, `scripture`만 씁니다. `skip: true`는 그 장을 뺍니다.
- `ref`는 참조 칸이 있는 슬라이드(성경 본문·성경봉독·설교)나 `op: scripture`와 함께 씁니다. `copy`·`quote`·`caption`과는 같이 못 씁니다.
- 규칙에 안 맞는 줄(모르는 `op`, 틀린 참조 형식, 글자 없는 장에 `quote`·`caption`·`scripture`)이 있으면 `plan refused: override slide N: …`가 나오고 멈춥니다.
- 처음 보는 양식은 원본 그대로 복사되고 `review:unclassified`가 붙습니다. 바꾸고 싶은 장만 적으면 됩니다.

## 문제 해결

오류가 나면 PowerShell 창의 마지막 몇 줄(특히 `오류:`로 시작하는 줄)을 그대로 복사해 전달합니다. 창 내용에는 슬라이드 글자가 나오지 않습니다.

| 증상 | 원인 | 해결 |
|---|---|---|
| `uv`(또는 `git`)를 찾을 수 없음 | 설치 직후 창에 경로가 반영 안 됨 | PowerShell 창 전체를 닫고 새로 열기. 그래도 안 되면 로그아웃 후 다시 로그인 |
| `winget`을 찾을 수 없음 | winget이 없는 Windows 10 | uv는 2단계의 대체 명령으로 설치, 코드는 ZIP으로 받기 |
| `git clone`이 `Repository not found` | 저장소 권한이 없는 계정으로 로그인됨 | 제어판 → 자격 증명 관리자 → Windows 자격 증명에서 `git:https://github.com`을 지우고 다시 시도, 또는 ZIP |
| `Failed to spawn: broadcast-ppt` 또는 `No pyproject.toml found` | 저장소 폴더 밖에서 실행했거나 ZIP이 한 겹 더 풀림 | `cd $HOME\Desktop\broadcast-ppt` 후 다시 실행. `바탕 화면\broadcast-ppt\pyproject.toml`이 있는지 확인 |
| `uv sync`가 `error sending request`·시간 초과로 실패 | 인터넷 차단·방화벽·프록시 | 다시 실행(받은 것은 캐시에 남음). 교회 망이 막혀 있으면 휴대폰 핫스팟으로 설치. ARM PC는 지원 안 함 |
| 처음 `uv sync`가 몇 분간 멈춘 듯함, 또는 `액세스가 거부되었습니다` | 파이썬 내려받는 중이거나 백신이 새 파일을 검사 중 | 5~10분 기다림. 오류면 다시 실행. 계속되면 백신 검사 예외에 `C:\Users\<사용자 이름>\Desktop\broadcast-ppt`와 `%LOCALAPPDATA%\uv` 추가 |
| "템플릿이 없습니다" | `local\templates\24.상반기.pptx`가 없거나 이름이 다름 | 템플릿을 그 이름으로 복사, 또는 `--template <경로>`로 지정 |
| "템플릿이 원본과 다릅니다" | PowerPoint로 다시 저장했거나 다른 판 | 원본 파일을 그대로 다시 복사. 안내에 나온 SHA256과 비교 |
| `plan refused: … .pptx로 읽을 수 없습니다` | 원본이 `.ppt`, Mac이 만든 `._` 파일, 또는 손상된 파일 | `.ppt`는 PowerPoint에서 `.pptx`로 다시 저장. `._` 파일이면 진짜 파일을 고름 |
| `오류: [Errno 2] …` / `지정된 파일을 찾을 수 없습니다` | 원본 경로·이름 오타 | Tab으로 이름 채우기. `$HOME\Desktop\broadcast-ppt`에서 실행했는지 확인 |
| "파일을 열 수 없습니다 … PowerPoint에서 닫고" | 결과 파일이 PowerPoint에 열려 있음 | PowerPoint에서 닫고 다시 실행 |
| `overrides file not found: …` | 수정 파일 경로가 틀림(메모장이 `.txt`를 붙였거나 폴더 이름이 다름) | 탐색기 **보기 → 파일 확장명** 켜고 이름 확인, 경로는 Tab으로 채우기 |
| "UTF-8로 저장해 주세요" | 수정 파일이 ANSI(cp949)로 저장됨 | 메모장 → 다른 이름으로 저장 → 인코딩 UTF-8 |
| `plan refused: override slide N: …` | 수정 파일 N번 줄 값이 틀림 | 메시지대로 고치고 다시 실행 |
| `오류:` 뒤에 `while scanning`·`found character '\t'` | 수정 파일 문법(Tab 들여쓰기, 콜론 뒤 공백 없음) | 공백 2칸 들여쓰기, `  17: {skip: true}` 형식 그대로 |
| critical이 1개 이상 | 방송에 쓰면 안 되는 장이 있음 | [2. 검증 결과 보기](#2-검증-결과-보기)대로 원본 장 번호를 찾아 수정 파일에 적기. 시간이 없으면 그 주는 원본 PPT |
| Mac과 Windows 결과가 다름 | 한쪽에만 `GEMINI_API_KEY`가 있음(실행할 때 `ai fallback: …` 줄이 나오면 켜진 것) | 두 곳을 똑같이. PowerShell 확인 `$env:GEMINI_API_KEY`, 끄기 `Remove-Item Env:GEMINI_API_KEY`. Mac은 `echo $GEMINI_API_KEY` / `unset GEMINI_API_KEY` |

## 명령과 옵션

모든 명령은 저장소 폴더에서 `uv run broadcast-ppt <명령>`으로 실행합니다. 설정·템플릿·`local/work` 기본 위치는 저장소 기준으로 정해지고, 명령에 직접 적는 경로(원본 PPT, `--overrides` 등)는 지금 있는 폴더 기준입니다.

| 명령 | 하는 일 |
|---|---|
| `run <원본.pptx>` | 계획 + 생성 + 검증을 한 번에. 매주 쓰는 명령 |
| `plan <원본.pptx>` | plan 표를 출력하고 `plan.json` 저장 |
| `build <plan.json>` | `plan.json`으로 방송본 생성 + 검증 |
| `validate <방송본.pptx> --plan <plan.json>` | 이미 만든 방송본만 다시 검증(`--report 파일`로 결과 위치 지정) |
| `analyze <파일 또는 폴더>` | 샘플 PPT 구조 분석(관리용). 실패한 파일이 있으면 종료 코드 1 |
| `summarize`, `catalog`, `calibrate` | 관리용. 아래 [관리용 명령](#관리용-명령) |

| 명령 | 옵션 | 효과 |
|---|---|---|
| `plan`, `run` | `--overrides 파일` | [수정 파일](#고치고-싶을-때-수정-파일) 적용 |
| `plan`, `run` | `--slide-size template` 또는 `source` | 그대로 복사하는 원본 장의 크기 맞춤. `template`(기본)은 방송본 폭(13.33인치)에 맞게 키움(4:3 원본이면 4/3배), `source`는 키우지 않음. 방송본 크기는 어느 쪽이든 템플릿 크기 |
| `plan`, `run` | `--no-scripture-both` | 성경 본문·교리문답·짧은 문장 장에서 원본(화면 전체용) 장을 빼고 자막만 만듦 |
| `plan`, `run` | `--sermon-material lower-third` 또는 `passthrough` | 짧은 문장 장(설교 문장, 수정 파일의 `op: caption` 포함)을 자막으로 만들지(기본), 원본 그대로 복사할지 |
| `plan`, `run` | `--work 폴더` | 결과 폴더의 상위 위치(기본 `local/work`) |
| `plan` / `run`, `build` | `--out 파일` | plan: `plan.json` 저장 위치 / run·build: 방송본 저장 위치(기본은 결과 폴더의 `broadcast.pptx`) |
| `build`, `run` | `--template 파일` | 템플릿 파일을 직접 지정 |
| `build`, `run` | `--force` | 계획 뒤에 템플릿이나 원본이 바뀌어도 생성 |
| `build` | `--no-validate` | 생성 후 검증을 건너뜀 |
| `analyze` | `--workers N`, `--force`, `--scope 파일` | 병렬 작업 수, 이미 분석한 파일도 다시 분석, 요약 범위 |
| `plan`, `run` | 환경변수 `GEMINI_API_KEY` (선택 `GEMINI_MODEL`) | 키가 있으면 규칙으로 못 정한 장을 Gemini에 물어봄. 기본 모델 `gemini-3.5-flash`. 확신도 0.8 미만이면 그대로 복사하고 `plan.json`의 `ai`에 제안만 남김 |

### 검수 표시

실행 결과의 info 항목입니다. 어느 장인지는 plan 표의 `[ ]` 안에 나옵니다.

| 표시 | 뜻 |
|---|---|
| `review:inherited_reference` | 절 번호로 시작하는 본문에 참조가 없어서, 앞 참조나 성경봉독 참조를 이어받음 |
| `review:reference_unknown` | 절 번호가 알고 있는 어느 참조에도 안 맞음. 자막의 참조 박스를 숨김 |
| `review:no_reference_tag` | 자막 한 장에 넣을 참조가 없어서 참조 박스를 숨김. 틀린 참조는 넣지 않음 |
| `review:no_reference` | 성경봉독 카드에 참조가 없음. 견본 참조가 남아 critical도 함께 나옴 → 수정 파일에 `N: {ref: "책 장:절-절"}` |
| `review:no_page_number` | 성경봉독 카드에 쪽수가 없음. 쪽수 줄을 비워 둠 |
| `review:no_sermon_title` | 설교 카드에 따옴표 제목이 없음. 견본 제목이 남아 critical도 함께 나옴 → 원본 PPT에서 제목을 “ ” 따옴표로 감싸고 다시 실행(수정 파일로는 제목을 넣을 수 없음. 급하면 `N: {op: copy}`). 설교 카드에 참조가 없어도 critical이 나오며 `N: {ref: "…"}`로 고침 |
| `review:no_quote` | 인용문 없는 묵상 장이라서 그대로 복사함 |
| `review:quote_too_long` | 묵상 인용이 너무 길어서 최소 글자 크기로 줄임 |
| `review:sentence_over_picture` | 사진 위의 짧은 문장을 자막으로 만듦 |
| `review:empty_item_body` | 교회소식 항목에 제목만 있고 내용이 없음 |
| `review:empty_slide` | 글자도 사진도 없는 장 |
| `review:graphic` | 도표나 질문지 장이라서 그대로 복사함 |
| `review:unclassified` | 어느 규칙에도 안 맞아서 그대로 복사함 |
| `review:text_dropped` | 성경 장에서 자막에도 참조 상자에도 넣지 못한 글상자가 있음(제목 줄, 쪽수 줄 등). 빠진 글자가 없는지 원본과 대조 |
| `review:shapes_layout` | 글자가 든 도형 2~4개짜리 장이라서 도형을 하단 띠 자리로 옮기고 배경을 크로마로 바꿈. 위치와 크기를 눈으로 확인하고, 원본 그대로가 나으면 수정 파일에 `N: {op: copy}` |
| `review:ai_classified` | 규칙으로 못 정한 장을 Gemini가 분류함(확신도 0.8 이상). 종류가 맞는지 확인. 글자는 원본 그대로 |
| `deferred:hymn` | 찬송가 악보·가사 그림을 그대로 복사함. 크로마 처리는 나중에 |

## 동작 원리

- **분류**: 원본 장마다 규칙(`rules/classify.py`, R010~R900)으로 종류를 정합니다. 예배로의 부름, 기도, 성경봉독, 설교, 묵상, 교리문답(제N문), 교회소식, 성경 본문, 핵심 문장 같은 종류입니다. 어느 규칙에도 안 맞으면 R900(미분류)이 되고 원본 그대로 복사합니다.
- **AI(선택)**: `GEMINI_API_KEY` 환경변수가 있을 때만, R900 장에 한해 Gemini에게 종류와 자막에 쓸 줄 번호만 묻습니다. 글자는 원본에서 다시 가져오고 좌표는 건드리지 않습니다. 호출이 실패하면 한 줄 안내만 출력하고 그 장은 그대로 복사합니다. 키가 없으면 아무것도 보내지 않습니다.
- **생성**: 템플릿 슬라이드를 복제하고 글자 상자 안의 글자만 바꿉니다. 긴 글은 줄 수에 맞는 높이의 템플릿(1·2·3줄)으로 나눠 여러 장에 담습니다.
- **두 번 만드는 장**: 성경 본문, 설교 인용, 짧은 문장(핵심 문장)은 원본(화면 전체용) 한 장 뒤에 녹색 하단 자막을 넣습니다. 교리문답은 원본, 문답 자막, 함께 적힌 말씀의 참조 자막을 따로 넣습니다.
- **그대로 복사하는 장**: 사진·영상·명단 장은 원본을 그대로 복사합니다(검수 표시 없음). 찬송가 악보, 도표·질문지, 규칙에 안 맞는 장도 그대로 복사하되 검수 표시(`deferred:hymn`, `review:graphic`, `review:unclassified`)를 붙입니다.
- **검증**: 만든 방송본을 다시 열어 점검합니다. critical은 장 수가 계획과 다름, 그림·영상 연결이 끊김, 채워야 할 칸이 비었거나 템플릿 견본 글자가 남음입니다. warning은 자막 줄 수·폭이 빠듯함, 녹색 배경이 아님, 도형 번호 중복입니다. info는 검수 표시입니다.

## 설정 파일

- `config/scope.yaml`: 분석 요약(`analyze`·`summarize`)에서 집계할 예배 종류와 기본 템플릿 위치(`templates:`). `run`은 예배 종류를 검사하지 않고 넣은 파일을 그대로 변환합니다.
- `config/roles.yaml`: 템플릿 슬라이드마다 역할, 글자 칸에 넣을 값, 배치 규칙(`clamp_to`, `shrink`, `stack`, `hide_when_empty` 등)을 손으로 적은 표. 각 키의 뜻은 파일 맨 위 주석에 있습니다. 카드 간격 값(`card_gap_pt`, `glyph`)은 Keynote에서 잰 값이라 템플릿을 고치면 `scripts/keynote_check.py calibrate-cards`로 다시 잽니다.
- `config/templates.yaml`: 템플릿과 `roles.yaml`로 자동 생성하는 파일입니다. 직접 고치지 말고 `catalog` 명령으로 다시 만듭니다.
- `config/widths.yaml`: 글자 폭 보정값. PowerPoint가 저장한 자동 맞춤 상자 크기로 계산했습니다.

### 관리용 명령

```bash
uv run broadcast-ppt catalog                                          # roles.yaml이나 템플릿을 고친 뒤 templates.yaml 다시 만들기
uv run broadcast-ppt calibrate <방송본 PPT 또는 폴더>                    # 글자 폭 보정값 다시 계산
uv run broadcast-ppt analyze local/input --out local/analysis         # 샘플 PPT 구조 분석
uv run broadcast-ppt summarize                                        # 이미 만든 분석 JSON으로 summary.csv, clusters.md 다시 만들기
uv run python scripts/pull_drive.py <Drive 폴더 ID>                     # Google Drive 폴더를 local/input으로 내려받기(gws 필요)
uv run python scripts/keynote_check.py calibrate-cards                # 카드 줄 위치를 Keynote로 다시 재기(Mac 전용, 가짜 이름 사용, local/input/미편집/에 주일오전예배ppt(…) 원본 하나 필요)
uv run python scripts/keynote_check.py bands <그림.png> ...             # 그림에서 글자가 그려진 세로 위치 재기
```

테스트는 `uv run pytest`입니다. `main`에 push하거나 PR을 열면 Windows에서도 같은 테스트가 자동으로 돕니다(GitHub Actions, `.github/workflows/tests.yml`).

## 아직 안 되는 것

- 찬송가 악보·가사 그림은 아직 크로마 처리를 안 합니다.
- 담임목사 위임식 같은 특별 예배는 주일 규칙으로 바뀌어서 따로 규칙이 필요합니다.
- 카드 간격은 Mac Keynote 기준으로 맞췄습니다. 개발용 Mac에는 Script MT Bold와 맑은 고딕이 없어서 Keynote가 다른 글꼴로 보여 줍니다. 송출 PC PowerPoint에서는 영문 부제가 한글 줄과 가까워지거나 겹칠 수 있습니다.
- 템플릿 성경 본문의 강조 색은 재현하지 않습니다. 자막 본문은 한 가지 색입니다.

## 데이터 보관 원칙

PPT에는 이름, 직분, 교회 정보가 들어 있습니다. 원본, 템플릿, 분석 결과, 계획, 수정 파일, 검증 결과, 확인용 그림, 방송본은 모두 `local/`에 두고 GitHub에 올리지 않습니다. 저장소에는 코드, 일반 설정, 문서만 올립니다. `config/templates.yaml`에는 위치, 서식, 글자 수, 템플릿 샘플 글자의 SHA-1 요약값만 저장하고 글자 자체는 저장하지 않습니다. 명령 출력과 검증 결과(`validation.json`)에는 글자 수 같은 통계만 나오고 슬라이드 글자는 나오지 않습니다. `plan.json`에는 슬라이드 글자(이름 포함)가 그대로 들어 있으니 전달하지 마세요.

`GEMINI_API_KEY`가 있으면 규칙으로 분류 못 한 장의 글자 줄, 도형 종류별 개수, 글상자 수, 전체 화면 사진 여부(사진 자체는 아님)만 Google Gemini로 나갑니다. 직분 앞 이름(`홍길동 집사`, `홍길동 장로님` → `[이름] 집사`, `[이름] 장로님`)은 가리고 보내지만, 직분 없이 쓴 이름이나 명단은 가려지지 않습니다. 슬라이드 노트, 표, 파일 자체와 수정 파일에 적은 장은 보내지 않습니다. 보내고 싶지 않은 주에는 키를 빼고 실행하세요.

## 백업과 복구

GitHub에는 코드, 설정, 문서만 있습니다. 아래 파일은 git에 없어서 프로젝트 폴더를 지우면 같이 사라집니다. 다시 쓸 파일은 교회 자료를 두는 비공개 위치(예: Google Drive)에 따로 보관하세요. 지우기 전에 커밋이 GitHub에 push됐는지(`git status`에 `ahead`가 없는지)도 확인하세요.

| 경로 | 내용 | 다시 쓸 때 |
|---|---|---|
| `local/templates/24.상반기.pptx` | 방송 템플릿 원본 | 변환에 필수 |
| `local/work/<원본 이름>/overrides.yaml` | 주마다 적은 수정 파일 | 그 주를 다시 만들 때만 |
| `local/pii-patterns.txt` | 커밋 전에 이름·개인정보를 막는 검사의 패턴 목록(검사 스크립트는 저장소에 없음) | 코드를 고칠 때 필수 |
| `local/input/`, `local/analysis/`, `local/work/`, `local/requirements/` | 원본 PPT(Drive에 원본 있음), 분석, 결과물, 요구사항 원문(따로 보관) | 보통 필요 없음 |

복구 순서: [설치](#설치) 순서대로 다시 설치 → 템플릿을 `local/templates/`에 복사 → `uv run pytest` → 코드를 고칠 계획이면 검사 패턴 파일도 제자리에 복사.
