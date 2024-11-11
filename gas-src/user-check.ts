const USER_LIST_SHEET_HEADER = {
  firstRow: ['送信先メールアドレス', '送信先会社名', '送信先担当者名', '送信者メールアドレス'],
  secondRow: ['to', 'companyName', 'staffName', 'from']
}

const MULTIPLE_SENDERS_CHECK_SHEET_INFO = {
  title: '確認1: 複数のAZメンバーが送信予定',
  header: ['送信先メールアドレス', '送信先会社名', '送信先担当者名', '送信者メールアドレス'],
}

const MULTIPLE_SENDERS_CHECK_UI = {
  threadhold: {
    uiTitle: '複数送信者数の設定',
    uiDescription: '例: 4人以上が送信予定の宛先リストを表示 => 4',
    uiAlertMessage: '無効な入力です。数値を入力してください。',
  },
  noData: {
    alertDescription: '指定人数以上が送信予定の宛先リストはありません',
  },
  dataExisted: {
    alertDescription: `指定人数以上が送信予定の宛先リストがあります。${MULTIPLE_SENDERS_CHECK_SHEET_INFO.title} シートを確認してください。`,
  }
}

const MULTIPLE_COMPANY_NAME_CHECK_SHEET_INFO = {
  title: '確認3: 送信先会社名のパターンが複数あり',
  header: ['送信先メールアドレス', '送信先会社名', '送信先担当者名', '送信者メールアドレス(送信先会社名ごと)'],
}

const MULTIPLE__COMPANY_NAME_CHECK_UI = {
  noData: {
    alertDescription: '全宛先の送信先会社名は同じです',
  },
  dataExisted: {
    alertDescription: `送信先会社名の記載パターンが複数存在する宛先リストがあります。${MULTIPLE_COMPANY_NAME_CHECK_SHEET_INFO.title} シートを確認してください。`,
  }
}

const MULTIPLE_STAFF_NAME_CHECK_SHEET_INFO = {
  title: '確認2: 送信先担当者名のパターンが複数あり',
  header: ['送信先メールアドレス', '送信先会社名', '送信先担当者名', '送信者メールアドレス(送信先担当者名ごと)'],
}

const MULTIPLE_STAFF_NAME_CHECK_UI = {
  noData: {
    alertDescription: '全宛先の送信先担当者名は同じです',
  },
  dataExisted: {
    alertDescription: `送信先担当者名の記載パターンが複数存在する宛先リストがあります。${MULTIPLE_STAFF_NAME_CHECK_SHEET_INFO.title} シートを確認してください。`,
  }
}

const REQUIRED_DATA_TYPE = {
  toBasedData: 'toBasedData',
  staffNameBasedData: 'staffNameBasedData',
  companyBasedData: 'companyBasedData'
}

const CHECK_NAME_TYPE = {
  staffName: 'staffName',
  companyName: 'companyName'
}

const SHEET_CLEAR_UI = {
  filePath: 'clear-sheets',
  title: {
    beforeClear: 'シート内データの削除',
    afterClear: '削除完了'
  }
}

const ADMIN_USER_EMAILS = [
  'tiger.tiger.1223@gmail.com',
]

type ContactData = {
  to: string,
  companyName: string,
  staffName: string,
  from: string,
}

type ContactDetails<T extends string> = {
  [to: string]: {
    [key in T]?: ContactData[]
  } | ContactData[]
}

type CheckSheetInfo = {
  title: string
  header: string[]
}

function onOpen(): void {
  const ui = SpreadsheetApp.getUi()
  const activeUserEmail = Session.getActiveUser().getEmail()
  if (!ADMIN_USER_EMAILS.includes(activeUserEmail)) {
    return
  }
  ui.createMenu('リスト確認')
    .addItem(`${MULTIPLE_SENDERS_CHECK_SHEET_INFO.title}`, 'checkMultipleSenderRecords')
    .addSeparator()
    .addItem(`${MULTIPLE_COMPANY_NAME_CHECK_SHEET_INFO.title}`, 'checkMultipleCompanyNames')
    .addSeparator()
    .addItem(`${MULTIPLE_STAFF_NAME_CHECK_SHEET_INFO.title}`, 'checkMultipleStaffNames')
    .addSeparator()
    .addItem(`${SHEET_CLEAR_UI.title.beforeClear}`, 'showSheetSelectionDialog')
    .addToUi()
}

function checkMultipleSenderRecords(): void {
  const targetNum = getUserForThreshold()
    if (targetNum === null) {
      return
    }
  processAndOutputData(
    MULTIPLE_SENDERS_CHECK_SHEET_INFO,
    REQUIRED_DATA_TYPE.toBasedData,
    targetNum
  )
}

function checkMultipleCompanyNames(): void {
  const targetNum = 2
  processAndOutputData(
    MULTIPLE_COMPANY_NAME_CHECK_SHEET_INFO,
    REQUIRED_DATA_TYPE.companyBasedData,
    targetNum
  )
}

function checkMultipleStaffNames(): void {
  const targetNum = 2
  processAndOutputData(
    MULTIPLE_STAFF_NAME_CHECK_SHEET_INFO,
    REQUIRED_DATA_TYPE.staffNameBasedData,
    targetNum
  )
}

function processAndOutputData(
  sheetInfo: CheckSheetInfo,
  requiredDataType: string,
  targetNum: number
): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadsheet.getSheets()
  const contactDetails: ContactDetails<string> = {}
  const excludedSheetNames: string[] = []

  processSheets(
    sheets,
    3,
    'A',
    'D',
    excludedSheetNames,
    contactDetails,
    requiredDataType
  )

  const targetData = getFilteredData(
    contactDetails,
    requiredDataType,
    targetNum
  )

  outputTargetDataToSheet(
    spreadsheet,
    sheetInfo,
    excludedSheetNames,
    targetData,
    requiredDataType
  )
}

function getUserForThreshold(): number | null {
  const ui = SpreadsheetApp.getUi()
  const response = ui.prompt(
    `${MULTIPLE_SENDERS_CHECK_UI.threadhold.uiTitle}`,
    `${MULTIPLE_SENDERS_CHECK_UI.threadhold.uiDescription}`,
    ui.ButtonSet.OK_CANCEL
  )

  if (response.getSelectedButton() !== ui.Button.OK) return null
  const threshold = Number(response.getResponseText())
  if (isNaN(threshold)) {
    ui.alert(`${MULTIPLE_SENDERS_CHECK_UI.threadhold.uiAlertMessage}`)
    return null
  }

  return threshold
}

function processSheets<T extends string>(
  sheets: GoogleAppsScript.Spreadsheet.Sheet[],
  startRow: number,
  startCol: string,
  endCol: string,
  excludedSheetNames: string[],
  contactDetails: ContactDetails<T>,
  requiredDataType: string
): void {
  sheets.forEach(sheet => {
    const sheetName = sheet.getName()
    if (!checkContactDataSheet(sheet)) {
      excludedSheetNames.push(sheetName)
      return
    }
    const range = sheet.getRange(`${startCol}${startRow}:${endCol}${sheet.getLastRow()}`)
    const rawDatas = range.getValues()
    
    if (requiredDataType === REQUIRED_DATA_TYPE.toBasedData) {
      addToBasedData(rawDatas, contactDetails)
    }
    if (requiredDataType === REQUIRED_DATA_TYPE.staffNameBasedData) {
      addNameBasedData(rawDatas, contactDetails, CHECK_NAME_TYPE.staffName)
    }
    if (requiredDataType === REQUIRED_DATA_TYPE.companyBasedData) {
      addNameBasedData(rawDatas, contactDetails, CHECK_NAME_TYPE.companyName)
    }

  })
}

function checkContactDataSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): boolean {
  const firstHeader = getHeader(sheet, 1, 'A', 'D')
  const secondHeader = getHeader(sheet, 2, 'A', 'D')
  return checkHeaderValidation(firstHeader, USER_LIST_SHEET_HEADER.firstRow) &&
    checkHeaderValidation(secondHeader, USER_LIST_SHEET_HEADER.secondRow)
}

function getHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet, rowNumber: number, startCol: string, endCol: string): string[] {
  const range = sheet.getRange(`${startCol}${rowNumber}:${endCol}${rowNumber}`)
  return range.getValues()[0]
}

function checkHeaderValidation(sheetHeader: string[], expectedHeader: string[]): boolean {
  return sheetHeader.length === expectedHeader.length &&
         sheetHeader.every((value, index) => value === expectedHeader[index])
}

function addToBasedData(
  rawDatas: any[],
  contactDetails: ContactDetails<string>
) {
  rawDatas.forEach(rawData => {
    addContactDetail(
      rawData,
      contactDetails
    )
  })
}

function addNameBasedData(
  rawDatas: any[][],
  contactDetails: ContactDetails<string>,
  checkName: string
) {
  rawDatas.forEach(rawData => {
    addContactDetail(
      rawData,
      contactDetails,
      checkName
    )
  })
}

function addContactDetail(
  rawData: any[],
  contactDetails: ContactDetails<string>,
  checkName?: string,
) {
  const to = rawData[0].trim()
  if (!to) return
  const companyName = rawData[1]
  const staffName = rawData[2].trim()
  const from = rawData[3].trim()

  if (!contactDetails[to]) {
    contactDetails[to] = []
  }
  
  if(!checkName) {
    (contactDetails[to] as ContactData[]).push({
      to,
      companyName,
      staffName,
      from
    })
  }
  if(checkName === CHECK_NAME_TYPE.companyName) {
    if (!contactDetails[to][companyName]) {
      contactDetails[to][companyName] = []
    }
    contactDetails[to][companyName].push({
      to,
      companyName,
      staffName,
      from
    })
  }
  if(checkName === CHECK_NAME_TYPE.staffName) {
    if (!contactDetails[to][staffName]) {
      contactDetails[to][staffName] = []
    }
    contactDetails[to][staffName].push({
      to,
      companyName,
      staffName,
      from
    })
  }
}

function getFilteredData<T extends string>(
  contactDetails: ContactDetails<T>,
  requiredDataType: string,
  targetNum: number
): ContactData[] {
  const targetDatas: ContactData[] = []
  switch (requiredDataType) {
    case REQUIRED_DATA_TYPE.toBasedData:
      getFilteredToBasedData(
        contactDetails,
        targetNum,
        targetDatas
      )
      break
      case REQUIRED_DATA_TYPE.companyBasedData:
        getFilteredNameBasedData(
          contactDetails,
          targetNum,
          CHECK_NAME_TYPE.companyName,
          targetDatas
        )
        break
      case REQUIRED_DATA_TYPE.staffNameBasedData:
        getFilteredNameBasedData(
          contactDetails,
          targetNum,
          CHECK_NAME_TYPE.staffName,
          targetDatas,
        )
        break
  }
  return targetDatas
}

function getFilteredToBasedData<T extends string>(
  contactDetails: ContactDetails<T>,
  targetNum: number,
  toBasedDatas: ContactData[],
) {
  for (let to in contactDetails) {
    const toBasedData = contactDetails[to] as ContactData[]
    if (toBasedData.length >= targetNum) {
      const fromValues: string = toBasedData.map(record => `・${record.from}`).join('\n')

      toBasedDatas.push({
        to: to,
        companyName: toBasedData[0].companyName,
        staffName: toBasedData[0].staffName,
        from: fromValues
      })
    }
  }
}

function getFilteredNameBasedData<T extends string>(
  contactDetails: ContactDetails<T>,
  targetNum: number,
  checkName: string,
  nameBasedDatas: ContactData[],
) {
  for (let to in contactDetails) {
    const names = Object.keys(contactDetails[to])
    if(names.length < targetNum) {
      return
    }

    const trimmedNames = names.map(name => name.replace(/\s+/g, '').trim())
    const originalNames = new Set(trimmedNames)
    
    if (originalNames.size === 1) {
      return
    }

    const relevantRecords: ContactData[] = []
    const fromGroup: { [key: string]: string[] } = {}

    names.forEach(name => {
      contactDetails[to][name].forEach(record => {
        relevantRecords.push({
          to: record.to,
          companyName: record.companyName,
          staffName: record.staffName,
          from: record.from
        })
        if (!fromGroup[name]) fromGroup[name] = []
        fromGroup[name].push(record.from)
      })
    })

    const uniqueNames = Array.from(new Set(relevantRecords.map(record => record[checkName] as string)))
    const fromDetails = uniqueNames.map((name, index) => {
      const fromList = fromGroup[name].join(', ')
      const suffix = index === uniqueNames.length - 1 ? '' : '\n'
      return `・${name}: ${fromList}${suffix}`
    }).join('')

    nameBasedDatas.push({
      to,
      companyName: checkName === CHECK_NAME_TYPE.companyName ? uniqueNames.join(', ') : relevantRecords[0].companyName,
      staffName: checkName === CHECK_NAME_TYPE.staffName ? uniqueNames.join(', ') : relevantRecords[0].staffName,
      from: fromDetails,
    })

  }
}

function outputTargetDataToSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  checkShInfo: CheckSheetInfo,
  excludedSheetNames: string[],
  targetDatas: ContactData[],
  requiredDataType: string
) {
  const ui = SpreadsheetApp.getUi()
  let resultSheet = spreadsheet.getSheetByName(checkShInfo.title)
  if (!resultSheet) {
    resultSheet = spreadsheet.insertSheet(checkShInfo.title)
  } else {
    resultSheet.clear()
  }

  resultSheet.appendRow(['対象外シート名', '', ...checkShInfo.header])

  if (excludedSheetNames.length > 0) {
    resultSheet.getRange(2, 1).setValue(excludedSheetNames.join('\n'))
  }
  console.log(targetDatas)

  if(targetDatas.length === 0) {
    console.log(targetDatas.length)
    switch (requiredDataType) {
      case REQUIRED_DATA_TYPE.toBasedData:
        ui.alert(MULTIPLE_SENDERS_CHECK_UI.noData.alertDescription)
        break
      case REQUIRED_DATA_TYPE.staffNameBasedData:
        ui.alert(MULTIPLE_STAFF_NAME_CHECK_UI.noData.alertDescription)
        break
      case REQUIRED_DATA_TYPE.companyBasedData:
        ui.alert(MULTIPLE__COMPANY_NAME_CHECK_UI.noData.alertDescription)
        break
    }
    return
  }

  targetDatas.forEach((result, index) => {
    const row = 2 + index
    resultSheet.getRange(row, 3, 1, 4)
      .setValues([[
        result.to,
        result.companyName,
        result.staffName,
        result.from
      ]])
  })
  switch (requiredDataType) {
    case REQUIRED_DATA_TYPE.toBasedData:
      ui.alert(`${MULTIPLE_SENDERS_CHECK_UI.dataExisted.alertDescription}`)
      break
    case REQUIRED_DATA_TYPE.staffNameBasedData:
      ui.alert(MULTIPLE_STAFF_NAME_CHECK_UI.dataExisted.alertDescription)
      break
    case REQUIRED_DATA_TYPE.companyBasedData:
      ui.alert(MULTIPLE__COMPANY_NAME_CHECK_UI.dataExisted.alertDescription)
      break
  }
}

function showSheetSelectionDialog(): void {
  const htmlOutput = HtmlService.createHtmlOutputFromFile(SHEET_CLEAR_UI.filePath)
    .setWidth(300)
    .setHeight(500)
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, SHEET_CLEAR_UI.title.beforeClear)
}

function getSheetsList(): string {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  let html = ''
  sheets.forEach(sheet => {
    html += `<input type="checkbox" name="${sheet.getName()}"> ${sheet.getName()}<br>`
  })
  return html
}

function clearSheetsData(sheetNames: string[]): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  sheetNames.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName)
    if (!sheet) {
      return
    }
    if (checkContactDataSheet(sheet)) {
      const lastRow = sheet.getLastRow()
      if (lastRow > 2) {
        sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent()
      }
    } else {
      sheet.clear()
    }
  })
  const ui = SpreadsheetApp.getUi()
  ui.alert(SHEET_CLEAR_UI.title.afterClear)
}