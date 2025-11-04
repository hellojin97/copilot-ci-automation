#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
이메일 전송 기능 테스트 스크립트
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from sales_data_analysis import SalesDataAnalyzer

def test_email_function():
    """
    이메일 전송 기능 테스트
    """
    print("🧪 이메일 전송 기능 테스트")
    print("="*50)
    
    try:
        # 분석기 생성
        analyzer = SalesDataAnalyzer("references/cicd_data.csv")
        
        # 분석 실행 (워드 파일 생성)
        results = analyzer.run_full_analysis()
        
        print("\n📧 이메일 전송 기능 준비 완료!")
        print("실제 테스트를 위해서는 다음 정보가 필요합니다:")
        print("1. Gmail 계정")
        print("2. Gmail 앱 비밀번호 (2단계 인증 필요)")
        print("3. 수신자 이메일 주소")
        
        print("\n💡 Gmail 앱 비밀번호 생성 방법:")
        print("1. Google 계정 설정 > 보안")
        print("2. 2단계 인증 활성화")
        print("3. 앱 비밀번호 생성")
        
        # 테스트 예시 코드 출력
        print("\n📝 사용 예시:")
        print("""
# 이메일 전송 예시
success = analyzer.send_email_with_report(
    docx_file_path=results['word_report'],
    sender_email="ilhj1228@gmail.com",
    sender_password="clqq xbqj jbzg nzjy",
    recipient_emails=["recipient@example.com"],
    subject="📊 판매 분석 보고서"
)
        """)
        
        return True
        
    except Exception as e:
        print(f"❌ 테스트 실패: {str(e)}")
        return False

if __name__ == "__main__":
    test_email_function()