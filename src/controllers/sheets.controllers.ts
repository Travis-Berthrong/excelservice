import axios, { AxiosResponse } from 'axios';
import { MicrosoftAccount } from '../entity/MicrosoftAccount';
import { IExcelSheet } from './interfaces/ExcelSheet.interface';
import { ICustomRequestError } from './interfaces/CustomRequestError.interface';
import { config } from 'dotenv';
import { sendAuthTokenRequest } from './auth.controllers';

config();

class SheetsController {
    private MicrosoftAccount: MicrosoftAccount;
    private workbookSessionId: string;
    private sheets: Array<IExcelSheet>;

    private async getSessionId(avoid_stack_overflow=false): Promise<Boolean> {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/createSession`;
        const data = {
            persistChanges: true
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`
        };
        try {
        const response = await axios.post(url, data, { headers });
        this.workbookSessionId = response.data.id;
        return true
        } catch (error) {
                if (error.response.status === 401 && !avoid_stack_overflow) {
                    console.log('request failed with 401, refreshing tokens')
                    const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
                    this.MicrosoftAccount.access_token = access_token;
                    this.MicrosoftAccount.refresh_token = refresh_token;
                    return this.getSessionId(true)
                }
                console.log(error.data)
                return false
        }
    }

    private async fetchSheets(avoid_stack_overflow = false): Promise<Array<IExcelSheet>> {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets`;
        const headers = {
            'Accept': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
        const response = await axios.get(url, { headers });
        return response.data.value;
        } catch (error) {
            if (!error.response) {
                console.log(error)
                return;
            }
            if (error.response.status == 404 && !avoid_stack_overflow) {
                await this.getSessionId();
                return this.fetchSheets(true);
            }
            if (error.response.status === 401 && !avoid_stack_overflow) {
                const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
                this.MicrosoftAccount.access_token = access_token;
                this.MicrosoftAccount.refresh_token = refresh_token;
                return this.fetchSheets(true);
            }
            console.log(error.code)
            console.log(error.response)
            return;
        }

    }

    private getSheetId(sheetName: string) {
        const sheet = this.sheets.find(sheet => sheet.name === sheetName)
        if (!sheet) return null;
        return encodeURIComponent(sheet.id);
    }

    private async fetchTables(avoid_stack_overflow = false): Promise<void> {
        try {
            const updatedSheets = this.sheets.map(async (sheet): Promise<IExcelSheet> => {
                const sheetId = encodeURIComponent(sheet.id)
                const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/${sheetId}/tables`;
                const headers = {
                    'Accept': 'application/json',
                    'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
                    'workbook-session-id': this.workbookSessionId
                };
                const response = await axios.get(url, { headers });
                const tables = response.data.value.reduce((acc, table) => {
                    return [...acc, { id: table.id, name: table.name }]
                }, [])
                sheet = {...sheet, tables: tables}
                return sheet;
            })
            this.sheets = await Promise.all(updatedSheets)
        } catch (error) {
            if (!error.response) {
                console.log(error)
                return;
            }
            if (error.response.status == 404 && !avoid_stack_overflow) {
                await this.getSessionId();
                return this.fetchTables(true);
            }
            if (error.response.status === 401 && !avoid_stack_overflow) {
                const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
                this.MicrosoftAccount.access_token = access_token;
                this.MicrosoftAccount.refresh_token = refresh_token;
                return this.fetchTables(true);
            }
            console.log(error.code)
            console.log(error.response)
            return;
        }
    }

    private async handleRequestError(error: any, avoid_stack_overflow: Boolean, retryFunction: Function, ...args: any[]) {
        if (!error.response) {
            console.log(error)
            return null;
        }
        if (error.response.status == 404 && !avoid_stack_overflow) {
            await this.getSessionId();
            return retryFunction(...args, true);
        }
        if (error.response.status === 401 && !avoid_stack_overflow) {
            const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
            this.MicrosoftAccount.access_token = access_token;
            this.MicrosoftAccount.refresh_token = refresh_token;
            return retryFunction(...args, true);
        }
        console.log(error.code)
        console.log(error.response)
        return error.response.data;
    }


    public async Init(MicrosoftAccount: MicrosoftAccount) {
        this.MicrosoftAccount = MicrosoftAccount;
        await this.getSessionId();
        this.sheets = await this.fetchSheets();
        await this.fetchTables();
    }

    public GetSheets() {
        return this.sheets;
    }

    public GetMicrosoftAccount() {
        return this.MicrosoftAccount;
    }

    public async AddSheet (sheetName: string, avoid_stack_overflow = false): Promise<AxiosResponse | ICustomRequestError> {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/add`;
        const data = {
            name: sheetName
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
        const response = await axios.post(url, data, { headers });
        this.sheets = await this.fetchSheets()
        await this.fetchTables()
        return response.data
        } catch (error) {
            try {
                return await this.handleRequestError(error, avoid_stack_overflow, this.AddSheet, sheetName);
            } catch (error) {
                console.log(error)
                return;
            }
        }
    }

    public async DeleteSheet(sheetName: string, avoid_stack_overflow = false): Promise<Boolean> {
        const sheetId: string = this.getSheetId(sheetName);
        if (!sheetId) throw new Error('Sheet not found')
        const url: string = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets('${sheetId}')`;
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
        await axios.delete(url, { headers });
        this.sheets = this.sheets.filter(sheet => sheet.name !== sheetName);
        return true
        } catch (error) {
            const result = await this.handleRequestError(error, avoid_stack_overflow, this.DeleteSheet, sheetName);
            if('error' in result) {
                return false
            }
        }
    }

   public async AddTable(sheetName: string, tableAddress: string, tableHasHeaders: Boolean = true, avoid_stack_overflow = false): Promise<AxiosResponse | ICustomRequestError> {
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error("Sheet not found")
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets('${sheetId}')/tables/add`;
        const data = {
            address: tableAddress,
            hasHeaders: tableHasHeaders,
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
        const response = await axios.post(url, data, { headers });
        this.sheets.find(sheet => sheet.name === sheetName).tables.push({ id: response.data.id, name: response.data.name });
        return response.data;
        } catch (error) {
            try {
                return await this.handleRequestError(error, avoid_stack_overflow, this.AddTable, sheetName, tableAddress, tableHasHeaders);
            }
            catch (error) {
                console.log(error)
                return;
            }
        }
    }

    //DELETE https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/tables/{id|name}
    public async DeleteTable(sheetName: string, tableName: string, avoid_stack_overflow = false): Promise<Boolean> {
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error("Sheet not found")
        if (!this.sheets.find(sheet => sheet.tables.find(table => table.name === tableName))) {
            throw new Error("Table not found")
        }
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/${sheetId}/tables/${tableName}`;
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
            await axios.delete(url, { headers });
            this.sheets.find(sheet => sheet.name === sheetName).tables = this.sheets.find(sheet => sheet.name === sheetName).tables.filter(table => table.name !== tableName);
            return true
        } catch (error) {
            try {
                const result = await this.handleRequestError(error, avoid_stack_overflow, this.DeleteTable, sheetName, tableName);
                if('error' in result) {
                    return false
                }
            } catch (error) {
                console.log(error)
                return false;
            }
        }
    }

    public async AddTableRows (sheetName: string, tableName: string, tableData: Array<Array<string>>, avoid_stack_overflow = false): Promise<AxiosResponse | ICustomRequestError> {
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error("Sheet not found")
        if (!this.sheets.find(sheet => sheet.tables.find(table => table.name === tableName))) {
            throw new Error("Table not found")
        }
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/${sheetId}/tables/${tableName}/rows`
        const data = {
            values: tableData
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
            const response = await axios.post(url, data, { headers })
            return response
        } catch (error) {
            try {
                return await this.handleRequestError(error, avoid_stack_overflow, this.AddTableRows, sheetName, tableName, tableData);
            } catch (error) {
                console.log(error)
                return;
            }
        }
    }

    //DELETE /me/drive/items/{id}/workbook/worksheets/{id|name}/tables/{id|name}/rows/{index}
    public async DeleteTableRows (sheetName: string, tableName: string, rowIndexs: Array<number>, avoid_stack_overflow = false): Promise<Boolean> {
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error("Sheet not found")
        if (!this.sheets.find(sheet => sheet.tables.find(table => table.name === tableName))) {
            throw new Error("Table not found")
        }
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/${sheetId}/tables/${tableName}/rows`
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
            for (const rowIndex of rowIndexs) {
                await axios.delete(`${url}/${rowIndex}`, { headers })
            }
            return true
        } catch (error) {
            try {
                const result = await this.handleRequestError(error, avoid_stack_overflow, this.DeleteTableRows, sheetName, tableName, rowIndexs);
                if('error' in result) {
                    return false
                }
            } catch (error) {
                console.log(error)
                return false;
            }
        }
    }

    public async GetTableRows (sheetName: string, tableName: string, avoid_stack_overflow = false): Promise<Array<Array<string>> | ICustomRequestError> {
        //GET https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/tables/{id|name}/rows?$top=5&$skip=5
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error("Sheet not found")
        if (this.sheets.find(sheet => sheet.tables.find(table => table.name === tableName)) ? false : true) {
            throw new Error("Table not found")
        }
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets/${sheetId}/tables/${tableName}/rows`
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        try {
            const response = await axios.get(url, { headers })
            const tableData = response.data.value.map((row: any) => { return { row_index: row.index, values: row.values[0] } })
            console.log(tableData)
            return tableData
        } catch (error) {
            try {
                return await this.handleRequestError(error, avoid_stack_overflow, this.GetTableRows, sheetName, tableName);
            } catch (error) {
                console.log(error)
                return;
            }
        } 
    }
}



export default SheetsController;
