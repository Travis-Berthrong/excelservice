# Microserivces Project - Excel API (Travis Berthrong)
This project is a microservice that allows users to interact with an excel workbook stored in OneDrive.
The API provides endpoints to create, delete, and get sheets in the workbook, create tables in the sheets, and add data to the tables.
The API also provides endpoints to authenticate with OneDrive and get the access token.
</br> 
</br>
This service is intended to be used in conjunction with the data analysis service (Junior Chi Emmanuel) which uses the data in the excel workbook to perform data analysis and visualization. A core service (Ilia Tiskin) is used to provide a common public-facing interface for the two services.


# Usage

To run this project, follow these steps:

1. Create an excel workbook in OneDrive with the name: `microservice_workbook.xlsx`
2. Run the following command to install the required dependencies:
    ```
    npm install
    ```
3. Make sure to set all the environment variables specified in the `.env.template` file.
4. Ensure that PostgreSQL is running.
5. Start the project by running the following command:
    ```
    npm start
    ```
6. Use the [Postman collection](https://speeding-shuttle-145414.postman.co/workspace/New-Team-Workspace~9bbc6a62-0def-40d9-bad3-56959c01b44b/collection/32573845-2bccb303-05c6-40f4-a20c-69393dea7322?action=share&creator=32573845) to test the API.

</br>
</br>

# Routes

## excel_auth

- GET `/excel_auth/get_auth_url` - Get the URL to authenticate with OneDrive
- GET `/excel_auth/get_access_token` - Get the access token from OneDrive and save it in the database

## excel_sheets 

- POST `/excel_sheets/create_session` - Create a session to work with the excel workbook (Must be done before any other operation)

     Query Params: 
     - `email` - Email of the user who has access to the workbook

- GET `/excel_sheets/` - Get all the sheets in the workbook
- POST `/excel_sheets/` - Create a new sheet in the workbook

     Body: 
     - `sheetName` - Name of the sheet to be added

- DELETE `/excel_sheets/:sheetName` - Delete a sheet from the workbook

- POST `/excel_sheets/table` - Create a table in a sheet

     Body:
     - `sheetName` - Name of the sheet to add the table
     - `tableAddress` - Address of the table in the sheet in the format: `A1:B2` (Start cell:End cell)
     - (optional) `tableHasHeaders` - Boolean value to specify if the table has headers 

- POST `/excel_sheets/table/:tableName/` - Add CSV data to a table

     Query Params:
     - `sheetName` - Name of the sheet that contains the table

     Body:
     - `file` - CSV file to be uploaded

- GET `/excel_sheets/table/:tableName/` - Get the data of a table
     
     Query Params:
     - `sheetName` - Name of the sheet that contains the table

# ROUTES

## excel_auth

- GET `/excel_auth/get_auth_url` - Get the URL to authenticate with OneDrive
- GET `/excel_auth/get_access_token` - Get the access token from OneDrive and save it in the database

## excel_sheets 

- POST `/excel_sheets/create_session` - Create a session to work with the excel workbook (Must be done before any other operation)

    Query Params: 
    - `email` - Email of the user who has access to the workbook
- GET `/excel_sheets/` - Get all the sheets in the workbook
- POST `/excel_sheets/` - Create a new sheet in the workbook

    Body: 
    - `sheetName` - Name of the sheet to be added

- DELETE `/excel_sheets/:sheetName` - Delete a sheet from the workbook

- POST `/excel_sheets/table`

    Body:
    - `sheetName` - Name of the sheet to add the table
    - `tableAddress` - Address of the table in the sheet in the format: `A1:B2` (Start cell:End cell)
    - (optional) `tableHasHeaders` - Boolean value to specify if the table has headers 

- POST `/excel_sheets/table/:tableName/` - Add csv data to the table

    Query Params:
    - `sheetName` - Name of the sheet that contains the table

    Body:
    - `file` - CSV file to be uploaded

- GET `/excel_sheets/table/:tableName/` - Get the data of the table
    
        Query Params:
        - `sheetName` - Name of the sheet that contains the table




