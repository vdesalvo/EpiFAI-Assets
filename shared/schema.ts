
import { pgTable, text, serial, integer, boolean, timestamp } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const names = pgTable("names", {
  id: serial("id").primaryKey(),
  name: text("name").notNull(),
  formula: text("formula"),
  comment: text("comment"),
  scope: text("scope"),
  status: text("status"),
});

export const insertNameSchema = createInsertSchema(names).omit({ id: true });

export type Name = typeof names.$inferSelect;
export type InsertName = z.infer<typeof insertNameSchema>;
