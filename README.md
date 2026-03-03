# TOA Companion — Setup Guide

## First-Time Setup

### 1. Install Node.js
Download and install from: https://nodejs.org (LTS version)

### 2. Add your credentials file
Place your Google API credentials file in the root of this folder:
```
toa-companion/
  credentials.json   ← paste your JSON file here
  main.js
  ...
```

### 3. Install dependencies
Open a terminal/command prompt in this folder and run:
```
npm install
```

### 4. Run the app
```
npm start
```

---

## Build an installer (.exe) for company PCs

Once everything is working:
```
npm run build
```
This creates a `dist/` folder with a Windows installer (.exe) you can distribute.

---

## Google Sheet Requirements

Your sheet must have:
- A tab called **Used** where consumed codes get moved with columns in this order:
  - A: Code
  - B: Employee Name
  - C: Customer Name
  - D: Date
  - E: Time
  - F: Code Type
  - G: Notes
- The service account email must have **Editor** access to the sheet

Behavior:
- Clicking **Generate Code** immediately logs the code to `Used` and removes it from the source tab.
- **Employee Name** and **Customer Name** are required and must each be at least 4 characters.

Service account email:
`toa-app@toa-appliction-api.iam.gserviceaccount.com`

---

## Troubleshooting

**"Failed to load tabs"** — Check that credentials.json is in the root folder and the sheet is shared with the service account email.

**Codes not moving to Used tab** — Make sure the service account has Editor (not Viewer) access.
