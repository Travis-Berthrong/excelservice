import express from 'express';
import cors from 'cors';

import authRoutes from './routes/auth.routes';

const app = express();
app.use(express.json());
app.use(cors());

app.use('/excel_auth', authRoutes);

export default app;