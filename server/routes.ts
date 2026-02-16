import type { Express } from "express";
import type { Server } from "http";
import { storage } from "./storage";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const MANIFEST_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">

  <Id>a1b2c3d4-e5f6-7890-abcd-123456789abc</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Epifai</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Epifai Name Manager"/>
  <Description DefaultValue="A better way to manage Named Ranges and Charts in Excel"/>
  <IconUrl DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE_URL}}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="{{BASE_URL}}"/>

  <AppDomains>
    <AppDomain>{{BASE_URL}}</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="{{BASE_URL}}/taskpane.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

</OfficeApp>`;

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    res.removeHeader("X-Frame-Options");
    res.removeHeader("Content-Security-Policy");
    if (req.method === "OPTIONS") {
      return res.sendStatus(200);
    }
    next();
  });

  app.get("/manifest.xml", (req, res) => {
    const host = req.headers.host;
    const protocol = req.headers["x-forwarded-proto"] || "https";
    const baseUrl = `${protocol}://${host}`;

    const manifest = MANIFEST_TEMPLATE.replace(/\{\{BASE_URL\}\}/g, baseUrl);
    res.header("Content-Type", "application/xml");
    res.send(manifest);
  });

  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", timestamp: new Date().toISOString() });
  });

  app.get("/api/names", async (req, res) => {
    const names = await storage.getNames();
    res.json(names);
  });

  app.get("/taskpane-test", (req, res) => {
    res.header("Content-Type", "text/html");
    res.send(`<!DOCTYPE html>
<html><head>
<meta charset="UTF-8"/>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
<title>Epifai Test</title>
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head><body>
<h1>Epifai Name Manager</h1>
<p id="status">Loading Office.js...</p>
<script>
Office.onReady(function(info) {
  document.getElementById('status').textContent = 'Office.js ready! Host: ' + info.host + ', Platform: ' + info.platform;
});
setTimeout(function() {
  if (document.getElementById('status').textContent.indexOf('Loading') === 0) {
    document.getElementById('status').textContent = 'Office.js did not initialize. This page may not be running inside Excel.';
  }
}, 5000);
</script>
</body></html>`);
  });

  return httpServer;
}
