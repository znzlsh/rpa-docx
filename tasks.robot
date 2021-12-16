*** Settings ***
Documentation           make word file rpa
Library                 RPA.Word.Application
Library                 DateTime
Task Setup              Open Application
Suite Teardown          Quit Application

*** Variables ***
${회사명}       ROBOBLOCK
${고객사명}     좋은고객사


*** Tasks ***
Minimal task
    Open File    .//template//temp.docx
    ${date}=    Get Current Date    result_format=datetime
    FOR    ${i}    IN RANGE   50  
        Replace Text    {회사명}    ${회사명} 
        Replace Text    {고객사명}  ${고객사명}
        Replace Text    {yyyy}      ${date.year}
        Replace Text    {mm}        ${date.month}
        Replace Text    {dd}        ${date.day}
    END
    
    Set Header    ${회사명} ( ${date.month} )월 정기 점검표
    Export To Pdf       .//output//${date.year}-${date.month}-${date.day}
    Save Document As    .//output//${date.year}-${date.month}-${date.day}