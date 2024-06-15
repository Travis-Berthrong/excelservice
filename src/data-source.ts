import "reflect-metadata"
import { DataSource } from "typeorm"
import { MicrosoftAccount } from "./entity/MicrosoftAccount"

export const AppDataSource = new DataSource({
    type: "postgres",
    host: "localhost",
    port: 5432,
    username: "postgres",
    password: "password",
    database: "excelservice",
    synchronize: true,
    logging: false,
    entities: [MicrosoftAccount],
    migrations: [],
    subscribers: [],
})
