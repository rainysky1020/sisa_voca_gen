# 단어장 삽화 대량 생성기

Excel의 `단어`, `뜻`을 읽어 캐릭터 시트 화풍의 투명 PNG 단어카드 삽화를 생성하고, 기술 검사와 비전 검수를 통과한 결과만 승인 폴더에 저장하는 Python 3.12+ 프로그램입니다.

## 가장 쉬운 실행 방법

Windows 탐색기에서 `단어장_삽화_실행.bat`를 더블클릭하고 화면의 번호를 선택하세요. 첫 실행에는 전용 가상환경과 필요한 패키지를 자동으로 설치합니다.

- `1`: 내 Excel을 비용 없이 검사
- `2`: 실제 이미지 1개 시험 생성
- `3`: 전체 이미지 생성
- `4`: 제공된 예제로 비용 없이 검사

실제 생성 메뉴는 비용 발생 경고 뒤에 `YES`를 입력해야 시작됩니다.

## 처리 규칙

- 장면 설계와 연상율 판정: 정확히 `gpt-5.6`
- 이미지 생성: 정확히 `gpt-image-2`, `low`, `1024x1024`, PNG, 투명 배경
- 결과 검수: 정확히 `gpt-5.4-mini`
- 연상율 40 미만은 이미지 API를 호출하지 않습니다.
- 최초 생성과 최대 2회 재생성을 합쳐 단어당 이미지 호출은 최대 3회입니다.
- `CONCRETE_OBJECT`는 사물만 허용합니다. 손가락·발가락 분리 표현은 무조건 실패합니다.
- 승인 파일명은 입력 단어와 정확히 같은 `{단어}.png`입니다. 금지문자나 충돌은 이름을 바꾸지 않고 보고합니다.

OpenAI API로 전송되는 데이터는 단어, 뜻, 캐릭터 시트, 생성 이미지입니다. 입력에 개인정보가 없는지 실행 전에 확인하세요.

## 설치

Python 3.12 이상에서 가상 환경을 만들고 설치합니다.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
Copy-Item .env.example .env
```

`.env`에 키를 저장합니다. 키는 코드·감사 로그·보고서에 기록되지 않습니다.

```dotenv
OPENAI_API_KEY=...
```

필요하면 `config.example.yaml`을 `config.yaml`로 복사해 경로와 동시성을 조정합니다. 모델과 이미지 설정 기본값은 제품 요구사항에 고정되어 있습니다.

## 입력

권장 위치는 `input/words.xlsx`, `input/character_sheet.png`입니다. 프로젝트 루트의 단일 후보도 자동 탐색합니다. 후보가 여러 개면 임의 선택하지 않고 `--excel`, `--character-sheet` 지정을 요구합니다.

Excel 첫 행에는 정확히 다음 필수 열이 있어야 합니다.

| 단어 | 뜻 |
|---|---|
| りんご | 사과 |
| 跳ぶ | 뛰다 |

바로 검사할 수 있는 예제는 `examples/words.example.xlsx`에 있습니다.

열 이름의 앞뒤 공백은 무시하지만 다른 열로 추측하지 않습니다. 빈 데이터 행은 무시하며, 단어 또는 뜻 하나만 빈 행은 `INPUT_ERROR`로 보고합니다. 원본 Excel과 캐릭터 시트는 수정하지 않습니다.

## 실행

먼저 무과금 입력 검증을 권장합니다. `--dry-run`은 OpenAI API를 전혀 호출하지 않습니다.

```powershell
python -m vocab_illustrator run --dry-run
python -m vocab_illustrator run
```

주요 옵션:

```text
--excel PATH
--character-sheet PATH
--output-dir PATH
--sheet NAME
--concurrency N
--resume | --no-resume
--overwrite
--dry-run
--limit N
--from-row N
--keep-rejected
```

기존 승인 이미지는 기본적으로 `EXISTING_SKIPPED`입니다. `--overwrite`일 때만 새 승인 결과로 원자적으로 대체합니다. `--no-resume`은 체크포인트를 읽지 않지만 원본 입력이나 기존 승인 파일을 삭제하지 않습니다.

## 출력과 복구

```text
output/images/{단어}.png
output/작업결과보고서.xlsx
work/attempts/{row_id}/attempt-{n}.png
work/run_state.json
logs/audit.jsonl
```

보고서는 `작업결과보고서` 워크시트 하나만 가지며, 실행 요약과 모든 처리 행을 포함합니다. 시간은 Asia/Seoul 기준입니다. 실패 후보는 최종 이미지 폴더에 들어가지 않습니다. `--keep-rejected`에서만 `work/rejected/`에 복사합니다.

체크포인트는 Excel, 캐릭터 시트, 운영 설정 해시가 모두 같을 때만 재사용합니다. 이미 성공한 항목은 최종 파일이 실제로 존재할 때만 재개 처리에서 건너뜁니다.

## 테스트와 정적 검사

테스트는 OpenAI 클라이언트를 모킹하며 과금 API를 호출하지 않습니다.

```powershell
pytest
ruff check .
ruff format --check .
mypy src
```

## 운영상 주의점

- `gpt-image-2`의 투명 배경은 공식 문서상 프리뷰 기능이므로 실제 알파 채널 기술 검사를 유지해야 합니다.
- 네트워크 타임아웃, 429, 5xx는 지수 백오프와 지터로 최대 5회 재시도합니다. 이는 품질 재생성 횟수에 포함되지 않습니다.
- 인증·권한·모델 오류는 즉시 중단합니다. 안전 정책 거절은 같은 요청을 반복하지 않습니다.
- 실제 생성은 비용과 외부 API 전송을 수반합니다. `--limit 1`로 운영 계정 접근 권한과 결과 품질을 먼저 확인하세요.

공식 호출 형식은 [GPT Image 2 안내](https://developers.openai.com/api/docs/guides/image-prompting), [구조화 출력](https://developers.openai.com/api/docs/guides/structured-outputs?api-mode=responses), [GPT-5.4 Mini](https://developers.openai.com/api/docs/models/gpt-5.4-mini)를 기준으로 구현했습니다.

