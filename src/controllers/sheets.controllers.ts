import axios from 'axios';
import { MicrosoftAccount } from '../entity/MicrosoftAccount';
import { IExcelSheet } from './interfaces/ExcelSheet.interface';
import { config } from 'dotenv';

config();

class SheetsController {
    private MicrosoftAccount: MicrosoftAccount;
    private workbookSessionId: string;
    private sheets: Array<IExcelSheet>;
    constructor(MicrosoftAccount: MicrosoftAccount) {
        this.MicrosoftAccount = MicrosoftAccount;
        this.getSessionId();
        this.fetchSheets();
    }

    private async getSessionId() {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/createSession`;
        const data = {
            persistChanges: true
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`
        };
        const response = await axios.post(url, data, { headers });
        if (response.status !== 201) {
            throw new Error(response.data);
        }
        this.workbookSessionId = response.data.id;
    }

    private async fetchSheets(avoid_stack_overflow = false): Promise<Array<IExcelSheet>> {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets`;
        const headers = {
            'Accept': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        const response = await axios.get(url, { headers });
        if (response.status == 404 && !avoid_stack_overflow) {
            await this.getSessionId();
            return this.fetchSheets(true);
        }
        if (response.status !== 200) {
            throw new Error(response.data);
        }
        return response.data;
    }

    public async getSheets() {
        return this.sheets;
    }

    public async addSheet (sheetName: string, avoid_stack_overflow = false) {
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/workbook/worksheets`;
        const data = {
            name: sheetName
        };
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        const response = await axios.post(url, data, { headers });
        if (response.status == 404 && !avoid_stack_overflow) {
            await this.getSessionId();
            return this.addSheet(sheetName, true);
        }
        if (response.status !== 201) {
            throw new Error(response.data);
        }
        this.sheets.push(response.data);
    }

    public async deleteSheet(sheetId: string, avoid_stack_overflow = false) {
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.MicrosoftAccount.workbook_id}/workbook/worksheets('${sheetId}')`;
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${this.MicrosoftAccount.access_token}`,
            'workbook-session-id': this.workbookSessionId
        };
        const response = await axios.delete(url, { headers });
        if (response.status == 404 && !avoid_stack_overflow) {
            await this.getSessionId();
            return this.deleteSheet(sheetId, true);
        }
        if (response.status !== 204) {
            throw new Error(response.data);
        }
        this.sheets = this.sheets.filter(sheet => sheet.id !== sheetId);
        return true
    }
}

export default SheetsController;
