# Codex용 구현 프롬프트

아래 내용을 새 Codex 작업에 그대로 붙여 넣어 사용한다. 캐릭터 시트와 단어 Excel 파일은 실행 전에 프로젝트 폴더에 배치한다.

---

당신은 이 프로젝트를 끝까지 구현하고 검증하는 시니어 Python 엔지니어다. 단순한 예제나 설계 설명으로 끝내지 말고, 현재 프로젝트 폴더에서 실제로 실행 가능한 프로그램, 테스트, 설정 예시, README를 완성하라.

먼저 프로젝트 폴더의 모든 관련 파일과 `AGENTS.md`를 확인하라. 기존 사용자 파일과 변경 사항을 보존하라. 캐릭터 시트 이미지와 Excel 파일을 찾아 구조를 검사하되 원본을 수정하지 마라. 후보 파일이 여러 개여서 자동 결정할 수 없을 때만 사용자에게 물어라.

## 프로젝트 목적

Excel의 `단어`, `뜻`을 읽고 단어장 앱 카드용 삽화를 대량 생성한다. 삽화의 목적은 학습자가 그림을 보고 단어 의미를 연상해 기억하도록 돕는 것이다. 다람쥐는 브랜드 캐릭터이지만 무조건 넣지 않는다. 구체적 사물 단어는 사물만 그리고, 다람쥐는 행동·표정·상태를 전달하는 데 꼭 필요한 경우에만 사용한다. 다람쥐가 없는 삽화도 캐릭터 시트의 선, 색, 형태, 명암, 분위기를 그대로 따라야 한다.

## 절대 요구사항

1. Python 3.12 이상과 공식 OpenAI Python SDK를 사용하라.
2. `.env`의 `OPENAI_API_KEY`를 사용하라. 키를 코드, 로그, 보고서, 오류 메시지에 노출하지 마라.
3. 입력은 `.xlsx`이며 필수 열은 `단어`, `뜻`이다.
4. 단어 하나에 승인 삽화 하나만 만든다.
5. 최종 파일명은 Excel의 단어와 정확히 같은 `{단어}.png`여야 한다.
6. 파일명 금지문자, Windows 예약 이름, 같은 단어의 서로 다른 뜻 충돌이 있으면 이름을 몰래 바꾸지 말고 생성하지 않은 뒤 보고서에 사유를 기록하라.
7. 장면 설계와 생성 전 연상율 판단은 정확히 `gpt-5.6`으로 수행하라.
8. 연상율은 0~100이며 40 미만이면 이미지 API를 호출하지 마라.
9. 이미지 생성은 정확히 `gpt-image-2`, `quality="low"`, `size="1024x1024"`, `output_format="png"`, `background="transparent"`로 수행하라.
10. 캐릭터 시트를 모든 이미지 생성 요청의 스타일 참조 이미지로 제공하라. `gpt-image-2`에는 `input_fidelity`를 보내지 마라.
11. 사물 단어에는 사물만 그리고 다람쥐나 다른 캐릭터를 넣지 마라.
12. 다람쥐가 등장하는 경우 캐릭터 시트의 동일한 정체성, 체형, 귀, 얼굴, 꼬리, 색, 외곽선, 명암을 보존하라.
13. 다람쥐의 손과 발은 둥근 벙어리장갑 같은 덩어리 형태여야 하며, 손가락·발가락·발톱·손가락선·발가락선이 절대 분리되어 보이면 안 된다.
14. 이미지에는 글자, 숫자, 단어, 뜻, 캡션, 로고, 워터마크, 프레임을 넣지 마라.
15. 생성 결과 검수는 정확히 `gpt-5.4-mini`의 이미지 입력 기능으로 수행하라.
16. 검수에는 생성 이미지, 캐릭터 시트, 단어, 뜻, 장면 설계를 함께 전달하라.
17. 검수 실패 시 검수 피드백을 `gpt-5.6`에 전달해 장면과 프롬프트를 재설계한 뒤 재생성하라.
18. 재생성은 최대 2회다. 최초 생성을 포함한 총 이미지 생성 시도는 단어당 최대 3회다.
19. 3번째 결과도 실패하면 최종 이미지 폴더에 이미지를 남기지 말고 마지막 실패 사유와 시도 이력을 보고서에 기록하라.
20. 전체 실행 뒤 `작업결과보고서.xlsx`를 생성하라. 워크북에는 `작업결과보고서`라는 워크시트 하나만 있어야 한다.

## 상세 설계 문서

프로젝트에 `단어장_삽화_대량생성_프로그램_상세기획서.md`가 있으면 이를 제품 요구사항 문서로 사용하라. 문서와 이 프롬프트가 충돌하면 이 프롬프트의 절대 요구사항을 우선하라.

## 입력 파일 탐색

- 설정에 명시된 경로를 우선한다.
- 기본값은 `input/words.xlsx`와 `input/character_sheet.png`다.
- 캐릭터 시트 후보가 없으면 명확한 오류로 중단하라.
- 캐릭터 시트 후보가 여러 개면 임의 선택하지 마라.
- Excel 필수 열을 찾지 못하면 사용 가능한 열을 알려 주고 중단하라.
- 입력 파일을 절대 덮어쓰지 마라.

## 장면 설계 구조화 출력

Pydantic 모델과 엄격한 JSON Schema를 사용해 최소한 다음을 받으라.

- `category`: `CONCRETE_OBJECT`, `ACTION`, `EMOTION_EXPRESSION`, `STATE_PROPERTY`, `PLACE_NATURE`, `ABSTRACT_RELATION`, `OTHER`
- `use_squirrel`: boolean
- `scene_summary_ko`
- `main_subject`
- `supporting_elements`
- `forbidden_elements`
- `association_score`: 0~100 정수
- `association_reason_ko`
- `ambiguity_risks`
- `image_prompt_en`
- `decision`: `GENERATE` 또는 `SKIP_LOW_ASSOCIATION`

장면은 하나의 중심 의미만 표현하고 작은 단어카드에서도 읽혀야 한다. 연상율 점수를 부풀리지 마라. 표현하기 어려운 기능어·추상어는 40점 미만 판정을 허용하라.

## 이미지 프롬프트 규칙

다음 공통 내용을 코드의 프롬프트 템플릿에 포함하라.

```text
Create one clean vocabulary-card illustration on a transparent background.
Use the attached squirrel character sheet strictly as the visual style reference.
Match its line quality, shapes, proportions, color palette, shading, and friendly educational tone.
Do not add any text, letters, numbers, captions, logos, watermark, frame, or decorative background.
Keep one clear focal subject, centered and large enough to recognize at small card size.
Use a true transparent background with clean alpha edges.
```

다람쥐 포함 시:

```text
Depict the same squirrel identity shown in the reference sheet.
The paws must be simple mitten-like shapes with absolutely no separated fingers.
The feet must be simple rounded shapes with absolutely no separated toes.
Do not add claws, digits, finger lines, toe lines, or human-like hands and feet.
```

다람쥐 제외 시:

```text
Do not draw the squirrel or any other character.
Draw only the target object or the minimum non-character elements required to express the meaning.
Even without the squirrel, preserve the illustration style, palette, line quality, shape language, and shading language of the attached character sheet.
```

## 검수

먼저 Pillow로 다음 기술 검사를 하라.

- PNG 디코딩 가능
- 정확히 1024×1024
- 실제 알파 채널 존재
- 일부 픽셀이 투명함
- 이미지가 완전 투명 또는 완전 단색이 아님

그다음 `gpt-5.4-mini`에 구조화 비전 검수를 요청하라. 최소 검수 항목은 다음과 같다.

- 단어와 뜻에 대한 의미 적합도
- 캐릭터 시트 화풍 일치도
- 캐릭터 필요 여부 정책
- 캐릭터 정체성 일치
- 손가락 분리 없음
- 발가락 분리 없음
- 작은 카드에서의 명확한 구도
- 글자·숫자·로고·워터마크 없음
- 투명 배경에 적합한 가장자리

의미 적합도와 화풍 일치도는 각각 70점 이상이어야 한다. 손가락·발가락 규칙과 사물 단어의 불필요한 캐릭터 금지는 점수 보정이 불가능한 필수 조건이다. 실패 시 `failure_codes`, 한국어 단문 사유, 다음 시도에 적용할 구체적인 `redesign_instructions_ko`를 반환하게 하라.

## 상태와 재시도

다음 상태 코드를 구현하라.

`SUCCESS`, `SKIPPED_LOW_ASSOCIATION`, `REDESIGN_LOW_ASSOCIATION`, `QA_FAILED_FINAL`, `INPUT_ERROR`, `INVALID_FILENAME`, `DUPLICATE_SKIPPED`, `DUPLICATE_CONFLICT`, `EXISTING_SKIPPED`, `API_ERROR`, `SAFETY_BLOCKED`, `CANCELLED`.

통신 오류 재시도와 품질 실패 재생성을 분리하라. 429, 5xx, 타임아웃은 지수 백오프와 지터로 최대 5회 재시도하되 품질 재생성 횟수에는 포함하지 마라. 인증·권한·잘못된 모델 오류는 즉시 중단하라. 안전 정책 거절은 동일 요청을 반복하지 마라.

## 파일 저장

- 후보 이미지는 `work/attempts/{row_id}/attempt-{n}.png`에 저장하라.
- 검수를 통과한 파일만 `output/images/{단어}.png`로 원자적으로 이동하라.
- 기본적으로 실패 이미지는 최종 출력에 포함하지 마라.
- `--keep-rejected`일 때만 진단용 실패 이미지를 `work/rejected/`에 유지하라.
- 기존 파일은 기본적으로 덮어쓰지 말고 `--overwrite`에서만 대체하라.

## 작업결과보고서

`openpyxl`로 `output/작업결과보고서.xlsx`를 생성하라. 워크시트는 하나만 만들고 이름은 `작업결과보고서`로 하라.

같은 시트 상단에는 실행 ID, 시작·종료 시각, 입력 파일, 캐릭터 시트 해시, 총 건수, 상태별 건수, 이미지 호출 수, 재생성 수, 실제 모델 및 설정을 요약하라.

하단 목록에는 다음 열을 포함하라.

`번호`, `단어`, `뜻`, `최종상태`, `연상율`, `캐릭터필요`, `장면요약`, `총생성횟수`, `재생성횟수`, `최종검수`, `검수점수`, `실패사유`, `결과파일`, `처리시간초`, `완료시각`.

자동 필터, 틀 고정, 줄바꿈, 적절한 열 너비, 성공/주의/실패 색상, A4 가로 한 페이지 너비 맞춤을 적용하라. 모든 입력 행이 보고서에 정확히 한 번 나타나야 한다.

## 안정성

- 단어별 체크포인트를 저장해 중단 후 재개할 수 있게 하라.
- 입력 Excel, 캐릭터 시트, 설정의 해시가 달라지면 이전 실행 상태를 자동 재사용하지 마라.
- 서로 다른 단어는 기본 3개까지 병렬 처리할 수 있지만 같은 단어의 생성·검수·재설계는 직렬로 처리하라.
- 감사 로그는 JSONL로 기록하되 API 키와 base64 이미지를 넣지 마라.
- 모든 시간은 보고서에서 Asia/Seoul 기준으로 표시하라.

## CLI

최소한 다음을 제공하라.

```text
python -m vocab_illustrator run
python -m vocab_illustrator run --dry-run
```

옵션은 `--excel`, `--character-sheet`, `--output-dir`, `--sheet`, `--concurrency`, `--resume`, `--no-resume`, `--overwrite`, `--dry-run`, `--limit`, `--from-row`, `--keep-rejected`를 지원하라.

`--dry-run`에서는 어떤 OpenAI API도 호출하지 마라.

## 테스트

실제 과금 API 없이 실행되는 테스트를 충분히 작성하라. OpenAI 클라이언트는 모킹하라.

반드시 검증할 사례:

1. 연상율 39점이면 이미지 API가 호출되지 않는다.
2. 연상율 40점이면 생성 단계로 이동한다.
3. 최초 검수 성공.
4. 1회 재생성 후 성공.
5. 2회 재생성 후에도 실패하여 이미지가 최종 폴더에 남지 않는다.
6. 단어당 이미지 생성이 최대 3회를 넘지 않는다.
7. 손가락 또는 발가락 실패가 승인되지 않는다.
8. 사물 단어에 다람쥐가 있으면 승인되지 않는다.
9. 잘못된 파일명은 변형 없이 실패 처리된다.
10. 동일 단어·다른 뜻 충돌이 검출된다.
11. 보고서가 한 워크시트이고 모든 행을 포함한다.
12. PNG 크기와 알파 채널 검사가 작동한다.
13. 중단 후 재개 시 이미 성공한 단어를 다시 생성하지 않는다.

## 구현 절차

1. 현재 파일과 요구사항을 조사하고 간단한 구현 순서를 제시하라.
2. 공식 OpenAI 문서에서 현재 SDK 호출 형식과 요청 파라미터를 확인하라. 요청된 모델을 임의로 다른 모델로 대체하지 마라.
3. 프로젝트를 구현하라.
4. 포매터, 정적 검사, 테스트를 실행하라.
5. 실패를 수정하고 테스트를 다시 실행하라.
6. 가능하면 `--dry-run`으로 실제 입력 파일을 검증하라.
7. 실제 API 호출은 사용자가 명시적으로 허용하지 않았다면 수행하지 마라. 프로그램 구현과 모킹 테스트는 끝까지 수행하라.
8. 최종 답변에는 완성된 파일, 실행 방법, 테스트 결과, 실제 API를 아직 호출하지 않았다면 그 사실, 남은 운영상 주의점을 간결하게 보고하라.

작동하지 않는 의사코드, TODO, 비어 있는 함수, 임시 더미 구현을 남기지 마라. 다만 실제 API 호출이 필요한 최종 실생성 검증은 비용과 외부 상태 변경을 수반하므로 사용자 허가 없이 실행하지 마라.
