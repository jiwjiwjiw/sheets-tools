import { Validation } from './Validation'

export class ValidationHandler {
  private static _instance: ValidationHandler
  validations: Validation[] = []

  private constructor () {}

  public static getInstance (): ValidationHandler {
    if (!ValidationHandler._instance) {
      ValidationHandler._instance = new ValidationHandler()
    }
    return ValidationHandler._instance
  }

  public add (validation: Validation) {
    this.validations.push(validation)
  }

  public update (
    modifiedRange: GoogleAppsScript.Spreadsheet.Range | undefined = undefined
  ): void {
    for (let validation of this.validations) {
      validation.update(modifiedRange)
    }
  }
}
