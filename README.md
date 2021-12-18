
문서 자동화

temp.docx 파일 설정 temp파일에 해당과 같이 적을경우
자동 변환됨.
--------------------------------------------------
{yyyy}     =>  현재 연도 지정

{mm}       =>  현재 월 지정

{dd}       =>  현재 일 지정

{COMPANY_NM}   =>  회사명 지정

{CUSTOMER_NM}  =>  고객사명 지정



powershell에서 아래 명령어로 파라미터 지정
--------------------------------------------------
[Environment]::SetEnvironmentVariable('COMPANY_NM', '좋은회사', 'User');

[Environment]::SetEnvironmentVariable('CUSTOMER_NM', '좋은고객사', 'User');



cmd에서 아래 명령어로 문서생성
--------------------------------------------------
run rcc

