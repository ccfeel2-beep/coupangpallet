# 쿠팡 파렛트 PPT 생성기

## 업데이트 내용
- 앱 시작 시 GitHub 공개 저장소 `ccfeel2-beep/coupang`의 `main/products.json`을 자동으로 읽습니다.
- 상품 관리 탭의 JSON 업로드 기능을 제거했습니다.
- GitHub 상품 목록 새로고침 버튼을 추가했습니다.
- 상품 추가 및 사용/중지 관리 기능을 유지했습니다.
- 현재 상품 목록 JSON 다운로드 기능을 유지했습니다.
- GitHub 연결에 실패하면 앱에 포함된 로컬 `products.json`이 있을 경우 예비 데이터로 사용합니다.

## GitHub에 적용하는 방법
1. 이 ZIP의 `app.py`, `requirements.txt`, `쿠팡 파렛트 양식.pptx`를 기존 저장소 최상위에 업로드해 교체합니다.
2. 기존 저장소의 `products.json`은 **교체하거나 삭제하지 말고 그대로 유지**합니다.
3. Commit changes를 누릅니다.
4. Streamlit Cloud가 재배포된 후 앱을 새로 열어 상품 목록이 자동으로 표시되는지 확인합니다.

## 상품 목록 수정 방법
1. 앱의 상품 관리 탭에서 상품을 추가하거나 사용/중지 상태를 변경합니다.
2. `현재 상품 목록 JSON 다운로드`를 누릅니다.
3. 다운로드한 `products.json`으로 GitHub 저장소 최상위의 기존 `products.json`을 교체하고 Commit changes를 누릅니다.
4. 앱의 `GitHub 상품 목록 새로고침` 버튼으로 반영 여부를 확인할 수 있습니다.

주의: GitHub 새로고침은 GitHub에 저장된 목록으로 현재 세션의 상품 목록을 덮어씁니다. 아직 GitHub에 저장하지 않은 변경사항이 있다면 먼저 JSON을 다운로드하세요.
