## Clasp and Google Apps Script (GAS) Usage

## About Google Apps Script (GAS)
- `Google Apps Script` (GAS) is a JavaScript-based scripting language that allows users to automate tasks across Google products and third-party services.

## About Clasp
- `Clasp` is an open-source tool provided by Google to manage GAS projects from the terminal. ([Github](https://github.com/google/clasp))
- By using `Clasp`, we can develop GAS projects using **TypeScript** and transpile `.ts` files to valid Apps Script files. (See [this doc](https://github.com/google/clasp/blob/master/docs/typescript.md) for details)

## How to Use Clasp and GAS
1. Install the dependencies:
    ```bash
    pnpm install
    ```

2. Create a `.clasp.json` file in the root directory as follows:
    ```json
    {
      "scriptId": "1zvBy5qk03iopFICaLC7oN5nUsMZ-3XU11EEv9nPTZ4ElU4T1_NFAgogm",  // scriptId is described in the Apps Script Editor under Project Settings > Script ID
      "rootDir": "./gas-src"
    }
    ```
    *scriptId is the ID of the GAS project associated with [this sandbox Google Sheet](https://docs.google.com/spreadsheets/d/1izSNNE7SyY4mimvQbxWG--siekWxlsnqZ_4Y67SIbo4/edit?gid=125002228#gid=125002228).

3. Use `clasp login` to log in. Otherwise, when using other clasp commands, you will encounter the following error:
    ```
    Could not read API credentials. Are you logged in globally?
    ```
***When you login, you are required to allow `clasp` to access your Google Account. Permission to do this action has been granted by the Security Department (Mimura-san).**

4. Create or update `.ts` files in the `src` directory.
    - Types for Google Apps Script are available.
    ```typescript
    example
    function checkContactDataSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
    // GoogleAppsScript type is used for sheet params
    ```

5. Use `clasp push` to push changes to the GAS project.

6. Use `clasp deploy` to deploy changes to the GAS project.

## Features Created by this GAS Project
Features created by this GAS project is used when the General Affairs department to check the contact list for nenga email. For more details on the features, please refer to [the manual slide about contact list check](https://docs.google.com/presentation/d/1upv99HlV6TRcbfl0xvPbFOiDCi5gGt_nm_HnPvipM-c/edit#slide=id.p).