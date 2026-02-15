
import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import path from "path";
import fs from "fs";

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  // Allow Excel to load resources from this origin (CORS)
  app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
  });

  // Serve the manifest dynamically replacing REPLIT_URL with the actual host
  app.get("/manifest.xml", (req, res) => {
    const host = req.headers.host;
    const protocol = req.headers["x-forwarded-proto"] || "https";
    const baseUrl = `${protocol}://${host}`;

    const manifestPath = path.join(__dirname, "manifest.xml");
    
    fs.readFile(manifestPath, "utf8", (err, data) => {
      if (err) {
        console.error("Error reading manifest:", err);
        return res.status(500).send("Error generating manifest");
      }

      const manifest = data.replace(/REPLIT_URL/g, baseUrl);
      res.header("Content-Type", "application/xml");
      res.send(manifest);
    });
  });

  // Basic API routes if needed (though mostly handled by Excel JS API)
  app.get("/api/names", async (req, res) => {
    const names = await storage.getNames();
    res.json(names);
  });

  return httpServer;
}
