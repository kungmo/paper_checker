# 시험문제 검토 도우미

AI를 활용한 고등학교 시험문제 자동 검토 시스템

## 설치 방법
```bash
uv pip install -r requirements.txt
```

## 환경 설정

`.env` 파일 생성 후 설정:
```
# Google Gemini API
GOOGLE_API_KEY=your_api_key

# MariaDB 설정
MARIADB_HOST=localhost
MARIADB_USER=your_username
MARIADB_PASSWORD=your_password
MARIADB_DATABASE=your_db_name
MARIADB_PORT=3306
```

## MariaDB 설정 SQL 스크립트
```
CREATE TABLE IF NOT EXISTS paper_review_logs (
    id INT AUTO_INCREMENT PRIMARY KEY,
    session_id VARCHAR(255) NOT NULL,
    subject VARCHAR(50) NOT NULL,
    exam_content TEXT,
    llm_review TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_session (session_id),
    INDEX idx_subject (subject),
    INDEX idx_created (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
```

## 실행 방법
```bash
chainlit run paper_checker.py --host 0.0.0.0 --port 8000
```

## 배포 웹 사이트
https://acer2.snu.ac.kr/paper_checker

## 라이선스

듀얼 라이선스 (MIT for non-commercial / Commercial License)

# 생기부 검토 도우미

중학교와 고등학교 시험문제를 AI로 검토해주는 도구입니다. 이 도구는 국어, 수학, 영어, 사회, 과학 시험문제의 검토를 지원합니다.

## 주요 기능
- 시험문제를 docx 파일로 변환하여 업로드
- AI 기반 오타 및 문법 검토
- 문항별 상세 피드백 제공

## 기술 스택
- Python 3.14
- Chainlit
- LangChain
- Google Gemini API
- Pandoc
- Pillow

## 라이선스

### 📚 비상업적 사용 (무료)
다음 사용자는 이 프로젝트를 자유롭게 사용하실 수 있습니다:
- 개인 사용자
- 초중고 및 대학 등 교육기관
- 교육청, 정부기관 등 공공기관
- 비영리 단체

**MIT 라이선스** 조건으로 사용, 수정, 배포가 가능합니다.

### 🏢 상업적 사용 (허가 필요)
영리 목적의 기업이나 조직에서 사용하시려면 별도 계약이 필요합니다.

📧 문의: kungmo@snu.ac.kr

자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 기여하기
버그 리포트, 기능 제안, Pull Request 환영합니다!
