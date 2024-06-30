import "reflect-metadata"
import { DataSource } from "typeorm"
import { MicrosoftAccount } from "./entity/MicrosoftAccount"
import { config } from "dotenv"
config()

if (!process.env.PG_USER || !process.env.PG_PASSWORD || !process.env.PG_DATABASE) {
    console.error("Missing environment variables")
    process.exit(1)
}

export const AppDataSource = new DataSource({
    type: "postgres",
    host: process.env.PG_HOST || "localhost",
    port: process.env.PG_PORT ? parseInt(process.env.PG_PORT) : 5432,
    username: process.env.PG_USER,
    password: process.env.PG_PASSWORD,
    database: process.env.PG_DATABASE,
    synchronize: true,
    logging: false,
    entities: [MicrosoftAccount],
    migrations: [],
    subscribers: [],
})
