*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
# Library             RPA.Desktop


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the Excel file
    Fill the form using the data from the Excel file
    Collect the results
    Export the table as a PDF
    [Teardown]    Log out and close the browser


*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Fill and submit the form
    Input Text    firstname    John
    Input Text    lastname    Smith
    Input Text    salesresult    123
    Select From List By Value    salestarget    10000
    Click Button    Submit

Fill and submit the form for one person
    [Arguments]    ${sales_data}
    Input Text    firstname    ${sales_data}[First Name]
    Input Text    lastname    ${sales_data}[Last Name]
    Input Text    salesresult    ${sales_data}[Sales]
    Select From List By Value    salestarget    ${sales_data}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel file
    Open Workbook    SalesData.xlsx
    ${sales_data}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_data}    IN    @{sales_data}
        Fill and submit the form for one person    ${sales_data}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png
    # Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_data}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_data}    ${OUTPUT_DIR}${/}sales_results_data.Pdf

Log out and close the browser
    Click Button    Log out
    Close Browser
