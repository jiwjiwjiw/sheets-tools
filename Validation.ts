import { rangeIntersect } from './helpers'

export class Validation {
  constructor (
    readonly validatedSheetName: string,
    readonly validatedRangeName: string,
    readonly validatingSheetName: string,
    readonly validatingRangeName: string,
    readonly allowInvalid: boolean = false,
    readonly additionalValidationValues: string[] = []
  ) {}

  update (
    modifiedRange: GoogleAppsScript.Spreadsheet.Range | undefined = undefined
  ): void {
    const validatedSheet = SpreadsheetApp.getActive().getSheetByName(
      this.validatedSheetName
    )
    if (!validatedSheet) {
      SpreadsheetApp.getUi().alert(
        `Tentative d'accès à la feuille inexistante "${this.validatedSheetName}"`
      )
      return
    }
    const validatedRange = validatedSheet.getRange(this.validatedRangeName)
    if (!validatedRange) {
      SpreadsheetApp.getUi().alert(
        `Tentative d'accès à la plage inexistante "${this.validatedRangeName}"`
      )
      return
    }
    const validatingSheet = SpreadsheetApp.getActive().getSheetByName(
      this.validatingSheetName
    )
    let validatingRange:
      | GoogleAppsScript.Spreadsheet.Range
      | undefined = undefined
    if (validatingSheet) {
      validatingRange = validatingSheet.getRange(this.validatingRangeName)
    }
    if (!modifiedRange || rangeIntersect(modifiedRange, validatingRange)) {
      let validationValues: string[] = []
      if (validatingRange) {
        validationValues = validationValues.concat(
          ...validatingRange.getDisplayValues()
        )
      }
      validationValues = validationValues.concat(
        ...this.additionalValidationValues
      )
      const rules = SpreadsheetApp.newDataValidation()
        .setAllowInvalid(this.allowInvalid)
        .requireValueInList(validationValues)
        .build()
      validatedRange.setDataValidation(rules)
    }
  }
}
