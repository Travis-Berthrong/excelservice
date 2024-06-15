import { AppDataSource } from "./data-source"
import { config } from 'dotenv';
config();
import app from './app';

AppDataSource.initialize().then(async () => {


    const port = process.env.PORT || 8001;

    app.listen(port, () => {
    console.log(`Server running on port ${port}`);
    });
}).catch(error => console.log(error))
