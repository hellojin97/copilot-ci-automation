# 📊 판매 데이터 분석 및 이메일 전송 가이드

## 🚀 기능 개요

이 스크립트는 CSV 판매 데이터를 분석하여 워드 보고서를 생성하고, 이메일로 전송하는 기능을 제공합니다.

## 📋 주요 기능

1. **데이터 분석**: CSV 파일의 판매 데이터 정제 및 분석
2. **워드 보고서 생성**: 전문적인 형태의 워드 문서 보고서 생성
3. **이메일 전송**: 생성된 보고서를 이메일로 첨부하여 전송

## 🛠️ 사용 방법

### 1. 기본 실행 (대화형)
```bash
uv run python sales_data_analysis.py
```

### 2. 코드로 이메일 전송
```python
from sales_data_analysis import SalesDataAnalyzer

# 분석기 생성 및 실행
analyzer = SalesDataAnalyzer("references/cicd_data.csv")
results = analyzer.run_full_analysis()

# 이메일 전송
success = analyzer.send_email_with_report(
    docx_file_path=results['word_report'],
    sender_email="your_email@gmail.com",
    sender_password="your_app_password",
    recipient_emails=["recipient1@example.com", "recipient2@example.com"],
    subject="📊 월간 판매 분석 보고서"
)
```

## 📧 이메일 설정

### Gmail 사용 시 (권장)
1. **2단계 인증 활성화**: Google 계정 > 보안 > 2단계 인증
2. **앱 비밀번호 생성**: 
   - Google 계정 > 보안 > 앱 비밀번호
   - "메일" 선택 후 생성
3. **생성된 16자리 비밀번호 사용**

### 다른 이메일 서비스
- **Outlook/Hotmail**: `smtp-mail.outlook.com`, 포트 587
- **Yahoo**: `smtp.mail.yahoo.com`, 포트 587
- **네이버**: `smtp.naver.com`, 포트 587

## 📊 생성되는 보고서 내용

### 1. 전체 요약
- 총 매출, 총 판매량, 평균 주문금액, 총 주문수

### 2. 상세 분석
- 카테고리별 분석
- 지역별 분석  
- 영업사원별 성과
- 상위 제품 (Top 10)

### 3. 데이터 품질 보고
- 처리된 데이터 이슈 목록

## 🔧 함수 매개변수

### `send_email_with_report()` 매개변수

| 매개변수 | 타입 | 필수 | 설명 |
|---------|------|------|------|
| `docx_file_path` | str | ✅ | 워드 파일 경로 |
| `sender_email` | str | ✅ | 발신자 이메일 |
| `sender_password` | str | ✅ | 발신자 비밀번호/앱 비밀번호 |
| `recipient_emails` | list | ✅ | 수신자 이메일 리스트 |
| `smtp_server` | str | ❌ | SMTP 서버 (기본: gmail) |
| `smtp_port` | int | ❌ | SMTP 포트 (기본: 587) |
| `subject` | str | ❌ | 이메일 제목 (기본: 자동 생성) |

## 🛡️ 보안 주의사항

1. **앱 비밀번호 사용**: 일반 비밀번호가 아닌 앱 비밀번호 사용
2. **환경변수 활용**: 비밀번호를 코드에 직접 입력하지 말고 환경변수 사용
3. **2단계 인증**: Gmail의 경우 2단계 인증 필수

### 환경변수 사용 예시
```python
import os

success = analyzer.send_email_with_report(
    docx_file_path=results['word_report'],
    sender_email=os.getenv('EMAIL_ADDRESS'),
    sender_password=os.getenv('EMAIL_PASSWORD'),
    recipient_emails=["recipient@example.com"]
)
```

## 🐛 문제 해결

### 일반적인 오류

1. **"Authentication failed"**
   - 앱 비밀번호 확인
   - 2단계 인증 활성화 확인

2. **"Connection refused"**
   - SMTP 서버/포트 확인
   - 방화벽 설정 확인

3. **"File not found"**
   - 워드 파일 경로 확인
   - 보고서 생성 완료 후 이메일 전송

### 디버깅 팁
- 먼저 보고서 생성이 정상 완료되는지 확인
- 단일 수신자로 테스트 후 여러 명으로 확장
- Gmail 외 다른 서비스 사용 시 SMTP 설정 확인

## 📁 파일 구조

```
110301-automation-csv/
├── references/
│   ├── cicd_data.csv                    # 원본 데이터
│   └── cicd_data_sales_report.docx      # 생성된 보고서
├── sales_data_analysis.py               # 메인 스크립트
├── test_email.py                        # 이메일 테스트
└── EMAIL_GUIDE.md                       # 이 가이드
```

## 📞 지원

문제가 발생하면 다음을 확인해주세요:
1. Python 환경 및 패키지 설치 상태
2. CSV 파일 형식 및 경로
3. 이메일 계정 설정 (2단계 인증, 앱 비밀번호)
4. 네트워크 연결 상태