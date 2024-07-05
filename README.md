# Microserivces Project - Excel API (Travis Berthrong)
This project is a microservice that allows users to interact with an excel workbook stored in OneDrive.
The API provides endpoints to create, delete, and get sheets in the workbook, create tables in the sheets, and add data to the tables.
The API also provides endpoints to authenticate with OneDrive and get the access token.
</br> 
</br>
This service is intended to be used in conjunction with the [data analysis service](https://github.com/chiemmanuel/analyticservice) (Junior Chi Emmanuel) which uses the data in the excel workbook to perform data analysis and visualization. A core service (Ilia Tiskin) is used to provide a common public-facing interface for the two services.


# Usage

To run this project, follow these steps:

1. Create an excel workbook in OneDrive with the name: `microservice_workbook.xlsx`
2. Run the following command to install the required dependencies:
    ```
    npm install
    npm i -D
    ```
3. Login to [Microsoft Entra](https://entra.microsoft.com/) and create a new application to get the `CLIENT_ID` and `SECRET_KEY` for OneDrive authentication. Upon registering the application, then ensure that the new application has a redirect URL such as `https://google.com/redirect` and the delegated permissions are set to `offline_access`, `profile`, `User.Read`, `Files.Read`, and `Files.readWrite`.
4. Make sure to set all the environment variables specified in the `.env.template` file within a local `.env` file.
5. Ensure that PostgreSQL is running.
6. Start the project by running the following command:
    ```
    npm start
    ```
7. Use the [Postman collection](https://speeding-shuttle-145414.postman.co/workspace/New-Team-Workspace~9bbc6a62-0def-40d9-bad3-56959c01b44b/collection/32573845-2bccb303-05c6-40f4-a20c-69393dea7322?action=share&creator=32573845) to test the API.

8. The `gen_test_csv.py` script can be used to generate a test CSV file to upload to the API. The script requires the `pandas` and `numpy` libraries to be installed. The script can be run with the following command:
    ```
    python gen_test_csv.py
    ```
    The script will generate a CSV file named `data_analysis_test.csv` in the root directory of the project.

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
     - `file` - CSV file to be uploaded. NOTE: the number of columns in the CSV file must match the number of columns in the table

- GET `/excel_sheets/table/:tableName/` - Get the data of a table
     
     Query Params:
     - `sheetName` - Name of the sheet that contains the table





