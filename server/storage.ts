
import { db } from "./db";
import {
  names,
  type Name,
  type InsertName,
} from "@shared/schema";
import { eq } from "drizzle-orm";

export interface IStorage {
  // Not used much in this app since it's an Excel Add-in,
  // but we keep it for potential future features like saving names to DB.
  getNames(): Promise<Name[]>;
  createName(name: InsertName): Promise<Name>;
}

export class DatabaseStorage implements IStorage {
  async getNames(): Promise<Name[]> {
    return await db.select().from(names);
  }

  async createName(insertName: InsertName): Promise<Name> {
    const [name] = await db.insert(names).values(insertName).returning();
    return name;
  }
}

export const storage = new DatabaseStorage();
