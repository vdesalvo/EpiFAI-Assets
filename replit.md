# Epifai Name Manager - Excel Add-in

## Overview

This is an **Excel Office Add-in** (Task Pane app) called "Epifai Name Manager" that provides a better way to manage Named Ranges and Charts in Excel. It runs as a web application that loads inside Excel's task pane via Office.js, allowing users to view, create, edit, delete, and navigate to named ranges, as well as manage charts across worksheets.

The app is built as a full-stack TypeScript application with a React frontend and Express backend. The primary interaction model is client-side — the frontend communicates directly with Excel via the Office.js/Excel.js API. The backend serves the app, provides the Office Add-in manifest XML, and has a minimal database layer for potential future features (like persisting named ranges externally).

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend (client/)
- **Framework**: React 18 with TypeScript
- **Routing**: Wouter (lightweight client-side router)
- **State/Data**: TanStack React Query for async state management. Excel data (names, charts) is fetched via Office.js API calls wrapped in custom hooks (`use-excel.ts`), not from the backend API
- **UI Components**: shadcn/ui component library (new-york style) built on Radix UI primitives
- **Styling**: Tailwind CSS with CSS variables for theming (green primary color matching Excel branding)
- **Animations**: Framer Motion for smooth transitions
- **Icons**: Lucide React
- **Build**: Vite with React plugin

### Key Frontend Files
- `client/src/lib/excel-names.ts` — Service layer wrapping Excel.js API for named range CRUD operations
- `client/src/lib/excel-charts.ts` — Service layer wrapping Excel.js API for chart operations
- `client/src/lib/excel-types.d.ts` — Minimal TypeScript declarations for the Office.js/Excel namespace
- `client/src/hooks/use-excel.ts` — React Query hooks wrapping the Excel service functions
- `client/src/components/NameList.tsx` — Main list view for named ranges with search/filter
- `client/src/components/NameEditor.tsx` — Form for creating/editing named ranges with range picker
- `client/src/pages/Home.tsx` — Main page with tabbed interface (Names tab + Charts tab)

### Office.js Integration
- The Office.js script is loaded via `<script>` tag in `client/index.html`
- All Excel interactions use `Excel.run()` context pattern
- In development without Excel, mock data is returned as fallback (check `import.meta.env.DEV`)
- The app requests `ReadWriteDocument` permissions in the manifest

### Backend (server/)
- **Framework**: Express 5 on Node.js
- **Language**: TypeScript, executed via tsx
- **Database**: PostgreSQL via `node-postgres` (pg), with Drizzle ORM
- **Schema**: Defined in `shared/schema.ts` using Drizzle's `pgTable`
- **Storage**: `server/storage.ts` implements a `DatabaseStorage` class (currently minimal — stores names but the app primarily reads from Excel directly)
- **API Routes**: Minimal — mainly serves the Office Add-in XML manifest at `/manifest.xml` with dynamic base URL substitution
- **Dev Server**: Vite dev server middleware integrated into Express for HMR
- **Production**: Static files served from `dist/public`

### Build System
- **Dev**: `tsx server/index.ts` runs the server with Vite middleware
- **Build**: Custom `script/build.ts` — Vite builds the client, esbuild bundles the server. Server dependencies are selectively bundled vs externalized for cold start optimization
- **Database migrations**: `drizzle-kit push` for schema sync

### Database Schema
Single table currently:
- **names** — `id` (serial PK), `name` (text), `formula` (text), `comment` (text), `scope` (text), `status` (text)

This table exists for potential future use (persisting names outside Excel). The primary data source for names is the Excel workbook itself via Office.js.

### Project Structure
```
client/          — React frontend
  src/
    components/  — App components + shadcn ui/ folder
    hooks/       — Custom React hooks (use-excel, use-toast, use-mobile)
    lib/         — Excel service layers, query client, utils
    pages/       — Route pages (Home, NotFound)
server/          — Express backend
  index.ts       — Server entry point
  routes.ts      — API routes (manifest serving)
  storage.ts     — Database storage interface
  db.ts          — Database connection
  vite.ts        — Vite dev middleware setup
  static.ts      — Production static file serving
shared/          — Shared between client and server
  schema.ts      — Drizzle database schema
  routes.ts      — API route type definitions
migrations/      — Drizzle migration files
```

## External Dependencies

### Database
- **PostgreSQL** — Required. Connection via `DATABASE_URL` environment variable. Uses Drizzle ORM for queries and `drizzle-kit` for schema management.

### Office.js (Microsoft)
- Loaded via CDN: `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`
- Provides the `Excel` global namespace for workbook interaction
- The app generates an Office Add-in XML manifest dynamically at `/manifest.xml`

### Key NPM Packages
- `drizzle-orm` + `drizzle-zod` — ORM and schema validation
- `@tanstack/react-query` — Async state management
- `framer-motion` — Animations
- `wouter` — Client-side routing
- `zod` — Schema validation
- Full suite of `@radix-ui/*` packages — Accessible UI primitives
- `express` v5 — HTTP server
- `connect-pg-simple` — PostgreSQL session store (available but sessions not heavily used currently)

### Fonts (Google Fonts)
- Inter, JetBrains Mono (via CSS import)
- DM Sans, Architects Daughter, Fira Code, Geist Mono (via HTML link)