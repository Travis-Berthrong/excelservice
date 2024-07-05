import express from 'express';
import cors from 'cors';

import authRoutes from './routes/auth.routes';
import sheetRoutes from './routes/sheets.routes';

const app = express();
app.use(express.json());
app.use(cors());

app.use('/excel_auth', authRoutes);
app.use('/excel_sheets', sheetRoutes)

export default app;