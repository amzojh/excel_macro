# excel_macro
User defined Excel Macro 
필요한 매크로들을 만들기
<사용법>

requirement

xlwings
pandas
numpy
pywin32
matplotlib
seaborn

1. 위 모듈을 설치한다.

2. https://github.com/xlwings/xlwings/releases 에서 xlwings.xlam파일을 다운로드 받은 후,
excel : Files -> Options -> Trust Center -> Trusted Location -> Excel default location: Excel StartUp로 Description이 써진 경로에 저 xlwings.xlam파일을 넣는다.

3. Files -> Options -> Trust Center -> Macro settings에서 Trust Aaccess to the VBA project object model을 활성화한다.

4. alt + f11 을 눌러서 엑셀 vba 에디터를 켠 후, refrence -> xlwings활성화

5. cmd창에 xlwings quickstart 친다.
