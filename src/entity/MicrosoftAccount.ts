import { Entity, Column, PrimaryColumn } from "typeorm"

@Entity()
export class MicrosoftAccount {

    @PrimaryColumn()
    email: string

    @Column()
    access_token: string

    @Column()
    refresh_token: string

    @Column()
    workbook_id: string

}