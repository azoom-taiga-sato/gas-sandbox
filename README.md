# Nenga Mail - Contact List Check
A repository to store Google Apps Script (GAS) project which checks contact list for New Year Greeting Emails

## About clasp
- clasp is an open-source tool provided by Google which allows developers to develop and manage GAS projects from one's terminal. ([Official Doc](https://github.com/google/clasp))
- Files in `gas-src` is linked to GAS project of a designated Google Spreadsheet
  - To associate this repository with a Google Spreadsheet, set `.clasp.json` in the root folder as follow
  ```
  {
   "scriptId":"",  //scriptId can be found in Google Apps Script Editor > Project Settings > Script ID
   "rootDir": "./gas-src"
  }
  ```