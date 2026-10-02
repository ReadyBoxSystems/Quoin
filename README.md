# Quoin

Quoin is a spreadsheet-first tool for turning working knowledge into structured, repeatable workflows.

Many teams rely on spreadsheet calculators, reference documents, and experienced people who know which assumptions matter. Quoin starts with the familiar spreadsheet grid instead of asking those teams to design an application first. Build or import the calculator, prove the math, then name and configure the important cells so they can power a cleaner Runner Preview.

Quoin is currently an internal prototype. It is ready for guided evaluation, but it is not yet a production system or shared cloud service.

## What You Can Do

- Enter normal values and formulas in an Excel-familiar grid.
- Use coordinate references, ranges, `IF`, `COUNT`, `ROUNDUP`, exact-match `VLOOKUP`, and exact/default `XLOOKUP`.
- Import a local `.xlsx` workbook while preserving its Sheets, values, formulas, and supported dropdowns.
- Name important cells to promote them into Smart Cells.
- Configure Smart Cell labels, input controls, annotations, Runner visibility, and conditional sections.
- Render dropdown and checkbox inputs in both the Sheet and Runner Preview.
- Generate a controlled Runner Preview from surfaced Smart Cells across the workbook.
- Experiment with Runner values without changing the defaults stored on the Sheet.
- Save multiple configurations locally in the browser.

The core Smart Cell roles are:

- **Input:** a value supplied by the author or runner.
- **Formula:** a calculated value used inside the workbook.
- **Output:** a result that can be shown in Runner Preview.

Lookup, action, validation, and compliance roles are available as advanced prototype concepts, but they are not the recommended starting point for evaluating the core workflow.

## Before You Start

Quoin currently runs locally on Windows. You will need:

- [GitHub Desktop](https://desktop.github.com/) or [Git for Windows](https://git-scm.com/download/win) to clone the public repository and receive updates.
- [Node.js LTS](https://nodejs.org/) installed with its standard options.
- A modern browser such as Microsoft Edge or Google Chrome.

No database, Quoin account, or cloud deployment is required.

> **Use a local folder:** Clone Quoin onto the computer's local drive, such as somewhere under `C:\Users\YourName` or another approved local development folder. Avoid network shares, mapped company drives, and cloud-synced folders unless your IT environment specifically supports Node.js development there. Network permissions, file locking, and synchronization can make dependency installation or the development server unreliable.

## Installation

### Option 1: GitHub Desktop — Recommended

1. Open GitHub Desktop. Signing in is optional for cloning this public repository, although your company may prefer that you use your GitHub account.
2. Choose **File → Clone repository**.
3. Select the **URL** tab and enter:

   ```text
   https://github.com/ReadyBoxSystems/Quoin.git
   ```

4. Choose a local path on the computer's own drive—not a network or mapped drive.
5. Select **Clone**.
6. In the cloned folder, double-click `Quoin.bat`.

GitHub Desktop provides the clearest update experience for someone who does not regularly use Git from a terminal.

### Option 2: Terminal

If Git for Windows and Node.js LTS are installed, open PowerShell in the local folder where Quoin should live and run:

```powershell
git clone https://github.com/ReadyBoxSystems/Quoin.git
cd Quoin
.\Quoin.bat
```

Cloning creates a real Git checkout that can receive future updates. Downloading the repository as a ZIP may run, but it does not provide the same straightforward update path and is not the recommended installation method.

### First Launch

On the first launch, Quoin installs its software dependencies. This can take a few minutes and requires an internet connection. Later launches are faster.

The launcher opens Quoin at:

```text
http://localhost:3000
```

Keep the Quoin command window open while using the application. Closing that window stops the local Quoin server.

## Five-Minute Walkthrough

The included demo uses fake, non-proprietary data. It demonstrates the product direction and is not an engineering calculator.

### 1. Start with the Sheet

- Click a normal cell and enter a value.
- Enter a formula such as `=2+2` or reference another cell.
- Change an input and confirm that its dependent formulas recalculate.

Normal spreadsheet work comes first. You do not have to configure Smart Cells before proving the calculator.

### 2. Inspect a Smart Cell

- Select one of the named cells in the demo.
- Review its Smart Cell name, role, display label, and Runner settings in the Inspector.
- Notice that the Smart Cell name is formula-safe while the display label can use natural wording.

Naming a cell is what promotes it into a Smart Cell.

### 3. Open Runner Preview

- Select **Runner** from the top workspace controls.
- Change a Runner input.
- Confirm that the outputs recalculate.
- Return to the Sheet and notice that the Sheet default was not overwritten.
- Use **Reset Preview** to clear the temporary Runner values.

### 4. Try a Controlled Input

- Find a dropdown or checkbox Smart Cell.
- Change it in the Sheet and Runner Preview.
- Try the demo's conditional section and observe which Runner fields appear or disappear.

### 5. Import a Workbook

- Open **File → Import Workbook**.
- Choose a disposable `.xlsx` calculator.
- Review the detected Sheets and any import review items.
- Import it as a new local configuration and switch among its Sheets.

Quoin preserves unsupported or risky formulas for review instead of silently deleting them. It does not promise complete spreadsheet compatibility.

## The Quoin Workflow

The intended authoring sequence is:

1. Build a calculator in Quoin or import an existing workbook.
2. Verify that the spreadsheet math behaves correctly.
3. Name the important cells.
4. Configure their roles, labels, controls, notes, and Runner visibility.
5. Use Runner Preview to test the controlled workflow.
6. Save the configuration locally.

The goal is not to hide the spreadsheet. Quoin keeps the calculation surface available to the author while giving the runner a focused view of the inputs and results that matter.

## Where Your Data Is Stored

Quoin is local-first at this prototype stage:

- Configurations are stored in the current browser profile on the current computer.
- Runner Preview values are temporary and do not rewrite Sheet defaults.
- Imported workbooks are processed locally in the browser; Quoin does not upload them to a Quoin service.
- Configurations do not sync between computers or users.
- Clearing browser site data can delete saved Quoin configurations.
- Private or incognito browser windows may discard configurations when closed.

Use test or non-sensitive workbook data during prototype evaluation. Quoin does not yet provide native configuration export, shared persistence, permissions, or backups.

## Getting Updates

The recommended update workflow is:

1. Close the Quoin command window if it is running.
2. Open the Quoin repository in GitHub Desktop.
3. Choose **Fetch origin**, then **Pull origin** if an update is available.
4. Double-click `Quoin.bat` again.

The launcher also attempts a Git update when command-line Git is available. GitHub Desktop remains the clearest way to see whether the update succeeded.

If you use the terminal instead, close Quoin and run the following from inside the cloned `Quoin` folder:

```powershell
git pull --ff-only
.\Quoin.bat
```

## Current Prototype Boundaries

Quoin currently includes:

- A Next.js, React, and TypeScript application.
- A deterministic formula engine powered by `mathjs`.
- Multi-Sheet local configurations.
- Local `.xlsx` workbook import.
- Browser-local persistence.
- Sheet and workbook error reporting.
- Undo and redo for workbook authoring changes.

Quoin does not yet include:

- A backend database or shared configuration storage.
- Authentication, user accounts, departments, or permissions.
- Published and immutable workflow versions.
- Saved Runner executions or audit reports.
- A desktop installer or native application package.
- Complete Excel compatibility, formatting, charts, pivots, macros, or VBA.
- First-class reference tables or live CSV-backed option lists.

This repository contains the runnable prototype application. Internal planning notes, local tests, fixtures, logs, and reference documents are intentionally kept outside the GitHub-tracked app surface.

## Troubleshooting

### “Node.js is not installed”

Install the current Node.js LTS release from [nodejs.org](https://nodejs.org/), accept the standard installation options, and then reopen `Quoin.bat`.

### Dependency installation fails

Confirm that the computer has internet access and that company security software permits Node.js and npm to download development packages. Then close the command window and run `Quoin.bat` again.

### The browser opens before Quoin is ready

Wait for the command window to report that the server is ready, then refresh `http://localhost:3000`.

### Port 3000 is already in use

Close other Quoin or Next.js command windows and launch `Quoin.bat` again.

### The page shows stale or missing application files

Close the Quoin command window and browser tab, then launch Quoin again. If the problem persists, delete the generated `.next` folder inside the Quoin directory and restart `Quoin.bat`. Do not delete `node_modules` unless dependency installation itself is broken.

### A saved configuration is missing

Confirm that you are using the same browser, browser profile, computer, and `http://localhost:3000` address used when it was saved. Configurations currently live only in that browser's local storage.

## Developer Commands

To run Quoin without the Windows launcher:

```bash
npm install
npm run dev
```

Then open `http://localhost:3000`.

TypeScript verification:

```bash
npm run typecheck
```

Avoid running a production build while someone is actively testing the development server; it can disturb the generated Next.js development files.
