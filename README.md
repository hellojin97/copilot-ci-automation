# 📊 자동 판매 데이터 분석 및 보고 시스템

이 프로젝트는 CSV 형식의 판매 데이터를 자동으로 분석하고, 요약 보고서를 Word 문서로 생성하여 지정된 이메일로 발송하는 파이썬 스크립트 및 자동화 워크플로우입니다.

## ✨ 주요 기능

- **데이터 로드 및 정제**: CSV 파일의 데이터를 로드하고, 날짜 형식 오류, 누락된 값 등을 자동으로 처리합니다.
- **심층 데이터 분석**: 카테고리, 지역, 영업사원, 제품별로 매출 및 판매량을 분석하여 다양한 인사이트를 도출합니다.
- **보고서 자동 생성**: 분석 결과를 바탕으로 체계적인 Word(.docx) 보고서를 생성합니다.
- **이메일 자동 발송**: 생성된 보고서를 첨부하여 지정된 수신자에게 이메일로 자동 발송합니다.
- **주간 자동 실행**: GitHub Actions를 통해 매주 월요일 오전 9시(UTC+9)에 전체 프로세스를 자동으로 실행합니다.

## 📂 프로젝트 구조

```text
.
├── .github/workflows/
│   └── weekly_sales_report.yml  # GitHub Actions 워크플로우
├── references/
│   └── cicd_data.csv            # 분석용 샘플 데이터
├── sales_data_analysis.py       # 메인 분석 및 리포팅 스크립트
├── pyproject.toml               # Poetry 의존성 관리 파일
└── README.md                    # 프로젝트 설명 파일
```

## ⚙️ 요구사항

- Python 3.9+
- `pandas`
- `python-docx`

## 🚀 사용 방법

### 1. 로컬 환경에서 실행

1. **저장소 복제**

   ```bash
   git clone https://github.com/hellojin97/copilot-ci-automation.git
   cd copilot-ci-automation
   ```

2. **의존성 설치 (Poetry 사용)**

   ```bash
   pip install poetry
   poetry install
   ```

3. **스크립트 실행**

   ```bash
   poetry run python sales_data_analysis.py
   ```

   - 스크립트가 실행되면 분석이 진행되고 Word 보고서가 생성됩니다.
   - 이메일 전송 여부를 묻는 메시지가 나타나면 `y`를 입력하고, 발신자/수신자 정보 및 Gmail 앱 비밀번호를 입력하여 이메일을 전송할 수 있습니다.

### 2. GitHub Actions를 통한 자동 실행

이 프로젝트는 매주 월요일 오전 9시(한국 시간)에 자동으로 보고서를 생성하고 이메일을 보내도록 설정되어 있습니다.

- **워크플로우 파일**: `.github/workflows/weekly_sales_report.yml`
- **수동 실행**: GitHub 저장소의 'Actions' 탭에서 `Weekly Sales Report Generation` 워크플로우를 선택하고 `Run workflow` 버튼을 클릭하여 수동으로 실행할 수도 있습니다.

#### 필요한 설정

자동 이메일 발송을 위해 GitHub 저장소에 다음 **Secrets**를 설정해야 합니다.

1. `Settings` > `Secrets and variables` > `Actions`로 이동합니다.
2. `New repository secret` 버튼을 클릭하여 아래의 시크릿들을 추가합니다.

- `SENDER_EMAIL`: 보고서를 발송할 Gmail 주소
- `EMAIL_PASSWORD`: 해당 Gmail 계정의 [앱 비밀번호](https://support.google.com/accounts/answer/185833)
- `RECIPIENT_EMAIL`: 보고서를 수신할 이메일 주소 (여러 명일 경우 쉼표로 구분)
