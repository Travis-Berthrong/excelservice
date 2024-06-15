import { Entity, PrimaryGeneratedColumn, Column } from "typeorm"

@Entity()
export class MicrosoftAccount {

    @PrimaryGeneratedColumn()
    id: number

    @Column()
    email: string

    @Column()
    access_token: string

    @Column()
    refresh_token: string

    @Column("text", { array: true })
    excel_sheets: string[][]

}