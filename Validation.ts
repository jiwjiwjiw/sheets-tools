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

  /**
   * Applies the dropdown to the validated range.
   *
   * When the allowed values come from a range only, the rule references that
   * range (`requireValueInRange`): it follows the sheet without any refresh
   * and has no item limit. A static list, with or without a range, becomes a
   * list rule (`requireValueInList`, at most 500 items).
   *
   * Nothing is written when a sheet is missing: a missing validating sheet
   * would otherwise turn the dropdown into an empty strict list that rejects
   * every input. No UI call either, as this may run from a trigger.
   */
  update (
    modifiedRange: GoogleAppsScript.Spreadsheet.Range | undefined = undefined
  ): void {
    const spreadsheet = SpreadsheetApp.getActive()
    const validatedSheet = spreadsheet.getSheetByName(this.validatedSheetName)
    if (!validatedSheet) {
      console.warn(`Feuille "${this.validatedSheetName}" introuvable`)
      return
    }
    const validatedRange = validatedSheet.getRange(this.validatedRangeName)

    let validatingRange: GoogleAppsScript.Spreadsheet.Range | undefined
    if (this.validatingSheetName !== '') {
      const validatingSheet = spreadsheet.getSheetByName(
        this.validatingSheetName
      )
      if (!validatingSheet) {
        console.warn(`Feuille "${this.validatingSheetName}" introuvable`)
        return
      }
      validatingRange = validatingSheet.getRange(this.validatingRangeName)
    }

    if (modifiedRange && !rangeIntersect(modifiedRange, validatingRange)) {
      return
    }

    const builder = SpreadsheetApp.newDataValidation().setAllowInvalid(
      this.allowInvalid
    )
    if (validatingRange && this.additionalValidationValues.length === 0) {
      builder.requireValueInRange(validatingRange, true)
    } else {
      let validationValues: string[] = []
      if (validatingRange) {
        validationValues = validationValues.concat(
          ...validatingRange.getDisplayValues()
        )
      }
      validationValues = validationValues.concat(
        ...this.additionalValidationValues
      )
      // Empty cells of the validating column would otherwise become empty
      // entries of the list, and duplicates inflate it towards the 500-item
      // limit of list validations.
      validationValues = validationValues.filter(
        (value, index, all) => value !== '' && all.indexOf(value) === index
      )
      builder.requireValueInList(validationValues, true)
    }
    validatedRange.setDataValidation(builder.build())
  }
}
