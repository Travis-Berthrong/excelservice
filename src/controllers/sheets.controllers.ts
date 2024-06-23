import axios from 'axios';
import { MicrosoftAccount } from '../entity/MicrosoftAccount';
import { IExcelSheet } from './interfaces/ExcelSheet.interface';
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

    private async fetchTables(avoid_stack_overflow = false) {
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

    public async Init(MicrosoftAccount: MicrosoftAccount) {
        this.MicrosoftAccount = MicrosoftAccount;
        await this.getSessionId();
        this.sheets = await this.fetchSheets();
        await this.fetchTables();
    }

    public GetSheets() {
        return this.sheets;
    }

    public async AddSheet (sheetName: string, avoid_stack_overflow = false) {
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
        if (!error.response) {
            console.log(error)
            return null;
        }
        if (error.response.status == 400 && error.response.data.error.code == 'ItemAlreadyExists') {
            throw new Error(`A sheet with this name already exists`)
        }
        if (error.response.status == 404 && !avoid_stack_overflow) {
            console.log('refreshing session id...')
            await this.getSessionId();
            return this.AddSheet(sheetName, true);
        }
        if (error.response.status == 401 && !avoid_stack_overflow) {
            console.log("refreshing tokens...")
            const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
            this.MicrosoftAccount.access_token = access_token;
            this.MicrosoftAccount.refresh_token = refresh_token;
            return this.AddSheet(sheetName, true);
        }
        console.log(error.code)
        console.log(error.response)
        throw new Error(error.response.data.error.message)
    }
    }

    public async DeleteSheet(sheetName: string, avoid_stack_overflow = false) {
        const sheetId = this.getSheetId(sheetName);
        if (!sheetId) throw new Error('Sheet not found')
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets('${sheetId}')`;
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
        if (!error.response) {
            console.log(error)
            return false;
        }
        if (error.response.status == 404 && !avoid_stack_overflow) {
            console.log('refreshing session id...')
            await this.getSessionId();
            return this.DeleteSheet(sheetName, true);
        }
        if (error.response.status == 401 && !avoid_stack_overflow) {
            console.log("refreshing tokens...")
            const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
            this.MicrosoftAccount.access_token = access_token;
            this.MicrosoftAccount.refresh_token = refresh_token;
            return this.DeleteSheet(sheetName, true);
        }
        console.log(error.code)
        console.log(error.response)
        throw new Error(error.response.data.error.message)

    }
    }

   public async AddTable(sheetName: string, tableAddress: string, tableHasHeaders: Boolean = true, avoid_stack_overflow = false): Promise<any> {
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
        if (!error.response) {
            console.log(error)
            return null;
        }
        if (error.response.status == 404 && !avoid_stack_overflow) {
            await this.getSessionId();
            return this.AddTable(sheetName, tableAddress, tableHasHeaders, true);
        }
        if (error.response.status === 401 && !avoid_stack_overflow) {
            const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
            this.MicrosoftAccount.access_token = access_token;
            this.MicrosoftAccount.refresh_token = refresh_token;
            return this.AddTable(sheetName, tableAddress, tableHasHeaders, true);
        }
        console.log(error.code)
        console.log(error.response)
        throw new Error(error.response.data.error.message)
        
    }
    }

    //POST /me/drive/items/{id}/workbook/worksheets/{id|name}/tables/{id|name}/rows
    public async AddTableRows (sheetName: string, tableName: string, tableData: Array<Array<string>>, avoid_stack_overflow = false) {
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
            if (!error.response) {
                console.log(error)
                return null;
            }
            if (error.response.status == 404 && !avoid_stack_overflow) {
                await this.getSessionId();
                return this.AddTableRows(sheetName, tableName, tableData, true);
            }
            if (error.response.status === 401 && !avoid_stack_overflow) {
                const { access_token, refresh_token } = await sendAuthTokenRequest(this.MicrosoftAccount.refresh_token, true);
                this.MicrosoftAccount.access_token = access_token;
                this.MicrosoftAccount.refresh_token = refresh_token;
                return this.AddTableRows(sheetName, tableName, tableData, true);
            }
            if (error.response.status === 504 && !avoid_stack_overflow) {
                return this.AddTableRows(sheetName, tableName, tableData, true);
            }
            console.log(error.code)
            console.log(error.response)
            throw new Error(error.response.data.error.message)
        }
    }
}

export default SheetsController;
