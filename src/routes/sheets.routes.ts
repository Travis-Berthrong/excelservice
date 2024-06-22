import express, { Router, Request, Response } from 'express';
import SheetsController from '../controllers/sheets.controllers';
import { MicrosoftAccount } from '../entity/MicrosoftAccount';
import { AppDataSource } from '../data-source';

const router: Router = express.Router();

const sheetsController = new SheetsController();

router.post('/create_session', async (req: Request, res: Response) => {
    try {
        const email = req.query.email as string;
        if (!email) {
            return res.status(400).json({ message: 'Email is required' });
        }
        const microsoftAccount = await AppDataSource.getRepository(MicrosoftAccount).findOne({ where : { email }});
        if (!microsoftAccount) {
            return res.status(404).json({ message: 'Account not found' });
        }
        await sheetsController.init(microsoftAccount);
        res.status(200).json({ message: 'Session created successfully' });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.get('/', async (req: Request, res: Response) => {
    try {
        const sheets = await sheetsController.getSheets();
        res.json(sheets);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.post('/', async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.body;
        const result = await sheetsController.addSheet(sheetName);
        if (!result) {
            res.status(500).json({ message: "Failed to create worksheet"})
        }
        res.status(201).json({ message: 'Sheet created', new_sheet: result });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.delete('/:sheetName', async (req: Request, res: Response) => {
    try {
        const sheetName = req.params.sheetName;
        await sheetsController.deleteSheet(sheetName);
        res.json({ message: 'Sheet deleted' });
    } catch (error) {
        if (error.message === "Sheet not found") return res.status(404).json({ message: error.message})
        res.status(500).json({ message: error.message });
    }
});

router.post('/table', async (req: Request, res: Response) => {
    try {
        const { sheetName, tableAddress, tableHasHeaders } = req.body;
        const response = await sheetsController.addTable(sheetName, tableAddress, tableHasHeaders)
        if (response) return res.status(201).json(response)
        return res.status(500).json({ message: 'Failed to create new table'})
    } catch (error) {
        if (error.message === "Sheet not found") return res.status(404).json({ message: error.message})
        res.status(500).json({ message: error.message });
    }
});

export default router;

