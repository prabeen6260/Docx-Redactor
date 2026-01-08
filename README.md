# Docx Redactor

A Microsoft Word Add-in that redacts sensitive information from documents, adds a confidentiality header, and enables Track Changes.

## Features

- **Redacts sensitive information:**
  - Email addresses
  - Social Security Numbers (SSN)
  - Phone numbers
  - Credit card numbers
  - Dates of birth (MM/DD/YYYY)
  - ID patterns (EMP, MRN, INS prefixed IDs)

- **Adds "CONFIDENTIAL DOCUMENT" header** at the top of the document

- **Enables Track Changes** to log all modifications

## To run the code

1. Install dependencies:
   ```
   npm install
   ```

2. Start the development server:
   ```
   npm start
   ```
   The program runs on local server at https://localhost:3000.

3. To add the add-in in Word:
   - **Word on the Web:** Go to Insert → Add-ins → Upload My Add-in → Select `manifest.xml`
   - **Word Desktop:** Follow [Microsoft's sideloading guide](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

4. Open your test document and click **"Redact & Protect"** in the add-in taskpane.

## Tech Stack

- React 19 compiler + TypeScript
- Vite
- Office.js (Word JavaScript API)
