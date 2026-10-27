# 쿠팡 파렛트 PPT 자동 생성 웹앱

## 기능
- `.xlsx` / `.xlsm` 발주표와 `.pptx` 양식 업로드
- 파렛트 번호별 행 그룹화
- 전체 파렛트 수-순번 형식으로 슬라이드 번호 생성 (예: `30-1` ~ `30-30`)
- 업체명 `주식회사 봉우리` 고정
- NO 자동 부여, 파렛트별 `box수량` 합계 계산
- 파렛트 개수 제한 없이 슬라이드 생성

## 로컬 실행
Python 3.10 이상 권장.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```
브라우저에서 `http://127.0.0.1:5000` 접속.

## 무료 웹 배포 개요 (Render 예시)
1. 이 폴더를 GitHub 저장소에 올립니다.
2. Render에서 New Web Service로 저장소를 연결합니다.
3. `render.yaml` 설정을 사용하거나 Build Command를 `pip install -r requirements.txt`, Start Command를 `gunicorn app:app`으로 설정합니다.
4. 배포 후 제공된 URL을 엽니다.

무료 호스팅의 제공 조건·절전 정책·용량 제한은 바뀔 수 있으며, 무료 서비스는 업무상 항상 이용 가능한 것을 보장하지 않습니다. 업로드된 발주 파일에는 거래처/물류 정보가 포함될 수 있으므로 회사의 보안 정책을 확인하고 공개 저장소나 공개 서버에 민감한 파일을 올리지 마세요.

## 주의
- 엑셀 표의 열 이름은 `파렛트`, `SKU No`, `SKU NAME`, `box수량`, `수량`, `제조일자`여야 합니다. 공백은 일부 허용합니다.
- 기본 정보는 시트 상단에 `입고예정일자`, `납품센터명`, `발주센터명`, `발주번호` 이름으로 있어야 합니다.
- 템플릿 첫 번째 슬라이드를 기본 디자인으로 복제합니다.
- 매우 많은 상품 행이 한 파렛트에 포함되면 표가 슬라이드 밖으로 넘칠 수 있으므로 실제 업무 데이터로 검수 후 사용하세요.
