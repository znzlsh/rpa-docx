*** Settings ***
Documentation           make word file rpa
Library                 RPA.Word.Application
Library                 DateTime
Task Setup              Open Application
Suite Teardown          Quit Application

*** Variables ***
${COMPANY_NM}     %{COMPANY_NM}
${CUSTOMER_NM}    %{CUSTOMER_NM}

*** Tasks ***
Minimal task
    Open File    .//template//temp.docx
    ${date}=    Get Current Date    result_format=datetime
    FOR    ${i}    IN RANGE   50  
        Replace Text    {COMPANY_NM}     ${COMPANY_NM} 
        Replace Text    {CUSTOMER_NM}    ${CUSTOMER_NM}
        Replace Text    {yyyy}           ${date.year}
        Replace Text    {mm}             ${date.month}
        Replace Text    {dd}             ${date.day}
    END
    
    Set Header    ${COMPANY_NM} ( ${date.month} )월 문서
    Export To Pdf       .//output//${date.year}-${date.month}-${date.day}
    Save Document As    .//output//${date.year}-${date.month}-${date.day}