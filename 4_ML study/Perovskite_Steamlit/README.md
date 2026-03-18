🧪 Perovskite AI Lab (V6)

이 프로젝트는 머신러닝(XGBoost, Random Forest, Gaussian Process)과 설명 가능한 AI(SHAP)를 활용하여 페로브스카이트 태양전지의 공정 조건을 최적화하고 성능을 예측하는 웹 애플리케이션입니다.

📌 주요 기능

데이터 전처리:

CSV/Excel 파일 자동 병합

복합 조성('A + B')의 Multi-Label Binarization (MLB) 처리

XGBoost 호환을 위한 컬럼명 자동 정제 및 중복 처리

ML 모델링:

XGBoost: 대규모 정형 데이터에 강력한 성능 (GridSearchCV 최적화 포함)

Random Forest: 데이터가 적거나 노이즈가 많을 때 안정적

Gaussian Process: 데이터가 매우 적을 때(Small Data) 베이지안 최적화 수행

설명 가능한 AI (XAI):

SHAP 분석을 통해 공정 변수가 효율(PCE)에 미치는 양/음의 영향력 시각화

최적화 제안:

최고 성능 샘플과 중요 변수(Feature Importance)를 분석하여 구체적인 실험 방향 제안

🛠️ 설치 및 실행 방법

1. 환경 설정 (Prerequisites)

Python 3.8 이상이 설치되어 있어야 합니다. 가상환경 사용을 권장합니다.

# 가상환경 생성 (Windows)
python -m venv venv
# 가상환경 활성화
.\venv\Scripts\activate

# 가상환경 생성 (Mac/Linux)
python3 -m venv venv
# 가상환경 활성화
source venv/bin/activate


2. 라이브러리 설치

필요한 패키지들을 설치합니다.

pip install -r requirements.txt


3. 앱 실행

Streamlit 앱을 로컬 서버에서 실행합니다.

streamlit run app.py


📂 프로젝트 구조

perovskite-ai-lab/
├── app.py              # 메인 애플리케이션 코드
├── requirements.txt    # 의존성 라이브러리 목록
├── README.md           # 프로젝트 설명서
└── data/               # (Optional) 실험 데이터 폴더


📚 Reference

본 시스템은 Nature Energy (2024) 등 최신 페로브스카이트 ML 연구 방법론을 기반으로 개발되었습니다.
