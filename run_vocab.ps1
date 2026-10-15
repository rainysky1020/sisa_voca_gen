$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

function Wait-ForUser {
    [void](Read-Host "계속하려면 Enter를 누르세요")
}

function Get-AppPython {
    if ($env:VOCAB_PYTHON -and (Test-Path -LiteralPath $env:VOCAB_PYTHON)) {
        return $env:VOCAB_PYTHON
    }

    $venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $venvPython) {
        return $venvPython
    }

    $systemPython = Get-Command python -ErrorAction SilentlyContinue
    if (-not $systemPython) {
        throw "Python을 찾지 못했습니다. https://www.python.org/downloads/ 에서 Python 3.12 이상을 설치하세요."
    }

    & $systemPython.Source -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 12) else 1)"
    if ($LASTEXITCODE -ne 0) {
        throw "Python 3.12 이상이 필요합니다."
    }

    Write-Host "처음 실행을 위한 환경을 준비합니다. 몇 분 걸릴 수 있습니다." -ForegroundColor Cyan
    & $systemPython.Source -m venv (Join-Path $PSScriptRoot ".venv")
    if ($LASTEXITCODE -ne 0) { throw "가상환경 생성에 실패했습니다." }
    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "pip 업데이트에 실패했습니다." }
    & $venvPython -m pip install -e $PSScriptRoot
    if ($LASTEXITCODE -ne 0) { throw "프로그램 설치에 실패했습니다." }
    return $venvPython
}

function Confirm-PaidRun([string]$Message) {
    Write-Host ""
    Write-Host $Message -ForegroundColor Yellow
    $answer = Read-Host "계속하려면 대문자 YES를 입력하세요"
    return $answer -ceq "YES"
}

function Get-InputFiles {
    $enteredPath = Read-Host "Excel 경로를 입력하세요. 그냥 Enter를 누르면 voca.xlsx 사용"
    if ([string]::IsNullOrWhiteSpace($enteredPath)) {
        $enteredPath = Join-Path $PSScriptRoot "voca.xlsx"
    }
    $excelPath = [System.IO.Path]::GetFullPath($enteredPath.Trim('"'))
    if (-not (Test-Path -LiteralPath $excelPath -PathType Leaf)) {
        throw "Excel 파일을 찾지 못했습니다: $excelPath"
    }

    $characterSheet = Join-Path $PSScriptRoot "다람쥐 캐릭터 시트.jpg"
    if (-not (Test-Path -LiteralPath $characterSheet -PathType Leaf)) {
        throw "캐릭터 시트를 찾지 못했습니다: $characterSheet"
    }
    return @($excelPath, $characterSheet)
}

function Invoke-VocabProgram([string[]]$Arguments) {
    Write-Host ""
    & $script:AppPython -m vocab_illustrator run @Arguments
    $exitCode = $LASTEXITCODE
    Write-Host ""
    if ($exitCode -eq 0) {
        Write-Host "작업이 끝났습니다." -ForegroundColor Green
        Write-Host "보고서: $(Join-Path $PSScriptRoot 'output\작업결과보고서.xlsx')"
    }
    else {
        Write-Host "작업이 오류로 끝났습니다. 위 메시지를 확인하세요. 종료 코드: $exitCode" -ForegroundColor Red
    }
    Wait-ForUser
}

try {
    $script:AppPython = Get-AppPython
    & $script:AppPython -c "import vocab_illustrator" *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "프로그램 패키지를 설치합니다." -ForegroundColor Cyan
        & $script:AppPython -m pip install -e $PSScriptRoot
        if ($LASTEXITCODE -ne 0) { throw "프로그램 설치에 실패했습니다." }
    }

    while ($true) {
        Clear-Host
        Write-Host "============================================================"
        Write-Host "                 단어장 삽화 대량 생성기"
        Write-Host "============================================================"
        Write-Host ""
        Write-Host "  1. 내 Excel 입력 검사 - 무료, API 호출 없음"
        Write-Host "  2. 이미지 1개 시험 생성 - API 비용 발생"
        Write-Host "  3. 전체 이미지 생성 - API 비용 발생"
        Write-Host "  4. 제공된 예제로 입력 검사 - 무료"
        Write-Host "  0. 종료"
        Write-Host ""
        $choice = Read-Host "번호를 입력하세요"

        try {
            switch ($choice) {
                "1" {
                    $files = Get-InputFiles
                    Invoke-VocabProgram @(
                        "--dry-run", "--excel", $files[0],
                        "--character-sheet", $files[1], "--no-resume"
                    )
                }
                "2" {
                    $files = Get-InputFiles
                    if (Confirm-PaidRun "1개 이미지를 실제로 생성합니다. OpenAI API 비용이 발생합니다.") {
                        Invoke-VocabProgram @(
                            "--excel", $files[0], "--character-sheet", $files[1], "--limit", "1"
                        )
                    }
                }
                "3" {
                    $files = Get-InputFiles
                    if (Confirm-PaidRun "전체 항목을 처리합니다. 여러 번의 OpenAI API 호출과 비용이 발생할 수 있습니다.") {
                        Invoke-VocabProgram @(
                            "--excel", $files[0], "--character-sheet", $files[1]
                        )
                    }
                }
                "4" {
                    Invoke-VocabProgram @(
                        "--dry-run",
                        "--excel", (Join-Path $PSScriptRoot "examples\words.example.xlsx"),
                        "--character-sheet", (Join-Path $PSScriptRoot "다람쥐 캐릭터 시트.jpg"),
                        "--no-resume"
                    )
                }
                "0" { exit 0 }
                default {
                    Write-Host "올바른 번호를 입력하세요." -ForegroundColor Yellow
                    Wait-ForUser
                }
            }
        }
        catch {
            Write-Host ""
            Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
            Wait-ForUser
        }
    }
}
catch {
    Write-Host ""
    Write-Host "실행 환경을 준비하지 못했습니다: $($_.Exception.Message)" -ForegroundColor Red
    Wait-ForUser
    exit 1
}
