*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log In
    Downloading Excel
    Fill and Submit data using Excel
    Collecting Results
    Creating PDF from HTML
    Log Out


*** Keywords ***
Open the intranet website
    Open available Browser    https://robotsparebinindustries.com/

Log In
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Fill and Submit Form for one person
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Downloading Excel
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Fill and Submit data using Excel
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and Submit Form for one person    ${sales_rep}
    END

Collecting Results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales-summary.png

Creating PDF from HTML
   Wait Until Element Is Visible    id:sales-results
   ${sales_results_html}=   Get Element Attribute    id:sales-results    outerHTML
   Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log Out
   Click Button    Log out
   Close Browser
