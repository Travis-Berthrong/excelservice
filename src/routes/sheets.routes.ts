import express, { Router, Request, Response } from 'express';
import SheetsController from '../controllers/sheets.controllers';
import { MicrosoftAccount } from '../entity/MicrosoftAccount';
import { AppDataSource } from '../data-source';
import multer from 'multer';
import * as fs from 'fs';
import * as fastCsv from 'fast-csv';
import path from 'path';

const router: Router = express.Router();
const uploadPath = path.join(__dirname, '/temp_uploads/')
const upload = multer({ dest: uploadPath });

const parseCsvData = async (csvFile: Express.Multer.File): Promise<string[][] | null> => {
  const filePath = `./src/routes/temp_uploads/${csvFile.filename}`;
  return new Promise((resolve, reject) => {
    const results: string[][] = [];
    fs.createReadStream(filePath)
      .pipe(fastCsv.parse({ headers: true }))
      .on('data', (row) => {
        row = [...Object.values(row)].filter((value: string) => !isNaN(parseFloat(value)))
        console.log('Row:', row);
        results.push(row);
      })
      .on('end', () => {
        fs.unlink(filePath, (err) => {
          if (err) {
            console.error('Error deleting file:', err);
            reject(err);
          } else {
            resolve(results);
          }
        });
      })
      .on('error', (error) => {
        console.error('Error processing file:', error);
        reject(error);
      });
  });
};

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
        await sheetsController.Init(microsoftAccount);
        res.status(200).json({ message: 'Session created successfully' });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.get('/', async (req: Request, res: Response) => {
    try {
        const sheets = await sheetsController.GetSheets();
        res.json(sheets);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.post('/', async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.body;
        const result = await sheetsController.AddSheet(sheetName);
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
        await sheetsController.DeleteSheet(sheetName);
        res.json({ message: 'Sheet deleted' });
    } catch (error) {
        if (error.message === "Sheet not found") return res.status(404).json({ message: error.message});
        res.status(500).json({ message: error.message });
    }
});

router.post('/table', async (req: Request, res: Response) => {
    try {
        const { sheetName, tableAddress, tableHasHeaders } = req.body;
        const response = await sheetsController.AddTable(sheetName, tableAddress, tableHasHeaders);
        if (response) return res.status(201).json(response);
        return res.status(500).json({ message: 'Failed to create new table'});
    } catch (error) {
        if (error.message === "Sheet not found") return res.status(404).json({ message: error.message});
        res.status(500).json({ message: error.message });
    }
});

router.post('/table/:tableName', upload.single('file'), async (req: Request, res: Response) => {
    try {
        let { sheetName } = req.query;
        sheetName = sheetName.toString();
        const { tableName } = req.params;
        if (!sheetName) return res.status(400).json('Invalid sheet name');
        if (!req.file) {
            return res.status(400).send('No file uploaded.');
        }
        const tableData = await parseCsvData(req.file);
        if (!tableData) return res.status(500).json("Error parsing csv");
        const response = await sheetsController.AddTableRows(sheetName, tableName, tableData);
        if (!response) return res.status(500).json('Failed to add table data');
        return res.status(201).json({ message: 'Data added successfully!'})
    } catch (error) {
        if (error.message === "Sheet not found" || error.message === "Table not found") {
            return res.status(404).json(error.message);
        }
        return res.status(500).json({ message: error.message });
    }
})

export default router;

