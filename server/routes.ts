import type { Express } from "express";
import type { Server } from "http";
import { storage } from "./storage";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const MANIFEST_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">

  <Id>a1b2c3d4-e5f6-7890-abcd-123456789abc</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Epifai</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Epifai Name Manager"/>
  <Description DefaultValue="A better way to manage Named Ranges and Charts in Excel"/>
  <IconUrl DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE_URL}}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://epifai.com"/>

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

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStartedTitle"/>
            <Description resid="GetStartedDesc"/>
            <LearnMoreUrl resid="LearnMoreUrl"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="EpifaiGroup">
                <Label resid="GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon16"/>
                  <bt:Image size="32" resid="Icon32"/>
                  <bt:Image size="80" resid="Icon80"/>
                </Icon>
                <Control xsi:type="Button" id="EpifaiButton">
                  <Label resid="ButtonLabel"/>
                  <Supertip>
                    <Title resid="ButtonLabel"/>
                    <Description resid="ButtonTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon16"/>
                    <bt:Image size="32" resid="Icon32"/>
                    <bt:Image size="80" resid="Icon80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>EpifaiPane</TaskpaneId>
                    <SourceLocation resid="TaskpaneUrl"/>
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
        <bt:Image id="Icon16" DefaultValue="{{BASE_URL}}/assets/icon-16.png"/>
        <bt:Image id="Icon32" DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
        <bt:Image id="Icon80" DefaultValue="{{BASE_URL}}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="TaskpaneUrl" DefaultValue="{{BASE_URL}}"/>
        <bt:Url id="LearnMoreUrl" DefaultValue="https://epifai.com"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Epifai"/>
        <bt:String id="ButtonLabel" DefaultValue="Name Manager"/>
        <bt:String id="ButtonTooltip" DefaultValue="Open Epifai Name Manager"/>
        <bt:String id="GetStartedTitle" DefaultValue="Epifai Name Manager"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStartedDesc" DefaultValue="Click Name Manager on the Home tab to manage your named ranges and charts visually."/>
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

  app.get("/api/names", async (req, res) => {
    const names = await storage.getNames();
    res.json(names);
  });

  return httpServer;
}
