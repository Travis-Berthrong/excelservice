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
        row = [...Object.values(row)].map((value: string) => !isNaN(parseFloat(value))? value : 0)
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

const validateSession = (req, res, next) => {
    if(!sheetsController.GetMicrosoftAccount()) {
        return res.status(401).json({ message: "Invalid session"})
    }
    next();
}

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

router.get('/', validateSession, async (req: Request, res: Response) => {
    try {
        const sheets = await sheetsController.GetSheets();
        res.json(sheets);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.post('/', validateSession, async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.body;
        const result = await sheetsController.AddSheet(sheetName);
        console.log('Result:', result);
        if (!result) {
            return res.status(500).json({ message: "Failed to create worksheet"})
        }
        if ('error' in result && result.error.code === 'ItemAlreadyExists') {
            return res.status(409).json({ message: 'Sheet already exists' });
        }
        if ('error' in result) {
            return res.status(500).json({ message: 'Error creating sheet' });
        }
        res.status(201).json({ message: 'Sheet created', new_sheet: result });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.delete('/:sheetName', validateSession, async (req: Request, res: Response) => {
    try {
        const sheetName = req.params.sheetName;
        await sheetsController.DeleteSheet(sheetName);
        res.json({ message: 'Sheet deleted' });
    } catch (error) {
        if (error.message === "Sheet not found") return res.status(404).json({ message: error.message});
        res.status(500).json({ message: error.message });
    }
});

router.post('/table', validateSession, async (req: Request, res: Response) => {
    try {
        const { sheetName, tableAddress, tableHasHeaders } = req.body;
        const response = await sheetsController.AddTable(sheetName, tableAddress, tableHasHeaders);
        if ('error' in response) {
            if (response.error.code === 'InvalidArgument') {
                return res.status(400).json({ message: 'Invalid table address' });
            }
            if (response.error.code === 'ItemAlreadyExists') {
                return res.status(409).json({ message: 'Table already exists' });
            }
            return res.status(500).json({ message: 'Failed to create new table'});
        }
        if (response) return res.status(201).json(response);
        return res.status(500).json({ message: 'Failed to create new table'});
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

router.delete('/table/:tableName', validateSession, async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.query;
        const { tableName } = req.params;
        if (!sheetName) return res.status(400).json('Invalid sheet name');
        await sheetsController.DeleteTable(sheetName.toString(), tableName);
        return res.json({ message: 'Table deleted' });
    } catch (error) {
        if (error.message === "Sheet not found" || error.message === "Table not found") {
            return res.status(404).json(error.message);
        }
        return res.status(500).json({ message: error.message });
    }
});

router.post('/table/:tableName', validateSession, upload.single('file'), async (req: Request, res: Response) => {
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
        if ('error' in response) {
            if (response.error.code === 'ItemNotFound') {
                return res.status(404).json({ message: 'Table not found' });
            }
            return res.status(500).json({ message: 'Failed to add table data'});
        }
        return res.status(201).json({ message: 'Data added successfully!'})
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
});

router.delete('/table/:tablename/rows', validateSession, async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.query;
        const { tableName } = req.params;
        const { rowIds } = req.body;
        if (!sheetName) return res.status(400).json('Invalid sheet name');
        if (!rowIds || rowIds.length === 0) return res.status(400).json('Invalid row ids');
        await sheetsController.DeleteTableRows(sheetName.toString(), tableName, rowIds);
        return res.json({ message: `${rowIds.length} rows deleted from ${tableName}` });
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
});

router.get('/table/:tableName', validateSession, async (req: Request, res: Response) => {
    try {
        const { sheetName } = req.query;
        const { tableName } = req.params;
        if (!sheetName) return res.status(400).json('Invalid sheet name');
        const tableData = await sheetsController.GetTableRows(sheetName.toString(), tableName);
        if (!tableData) return res.status(404).json('Table data not found');
        if ('error' in tableData) {
            if (tableData.error.code === 'ItemNotFound') {
                return res.status(404).json({ message: 'Table not found' });
            }
            return res.status(500).json({ message: 'Failed to get table data'});
        }
        return res.json(tableData);
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
});

export default router;

