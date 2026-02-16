import type { Express } from "express";
import type { Server } from "http";
import { storage } from "./storage";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const SERVER_START_TIME = new Date().toISOString();

const MANIFEST_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">

  <Id>9f95c57c-a36f-4f05-b486-e01f61f3cbaa</Id>
  <Version>1.3.0.0</Version>
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
    <SourceLocation DefaultValue="{{BASE_URL}}"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{{BASE_URL}}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{{BASE_URL}}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="{{BASE_URL}}"/>
        <bt:Url id="Commands.Url" DefaultValue="{{BASE_URL}}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{{BASE_URL}}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Epifai Name Manager"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Epifai"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Name Manager"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Click Name Manager on the Home tab to manage your named ranges and charts."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open Epifai Name Manager to manage named ranges and charts"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>

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

  app.get("/manifest-simple.xml", (req, res) => {
    const host = req.headers.host;
    const protocol = req.headers["x-forwarded-proto"] || "https";
    const baseUrl = `${protocol}://${host}`;

    const simpleManifest = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>9f95c57c-a36f-4f05-b486-e01f61f3cbaa</Id>
  <Version>1.3.0.0</Version>
  <ProviderName>Epifai</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Epifai Name Manager"/>
  <Description DefaultValue="A better way to manage Named Ranges and Charts in Excel"/>
  <IconUrl DefaultValue="${baseUrl}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${baseUrl}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="${baseUrl}"/>
  <AppDomains>
    <AppDomain>${baseUrl}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${baseUrl}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>`;
    res.header("Content-Type", "application/xml");
    res.send(simpleManifest);
  });

  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", timestamp: new Date().toISOString() });
  });

  app.get("/api/names", async (req, res) => {
    const names = await storage.getNames();
    res.json(names);
  });

  app.get("/api/build-info", (req, res) => {
    res.json({ buildTime: SERVER_START_TIME, version: "1.3" });
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
