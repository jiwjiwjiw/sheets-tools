import { Validation } from './Validation'
import { ValidationHandler } from './ValidationHandler'

export interface EmailTemplateParams {
  sheetName: string
  conditions: string[]
  insertData(html: string, data: any): string
  evaluateCondition(condition: string, data: any): boolean
}

export class EmailTemplate {
  private textEntries: Array<{
    type: string
    richText: GoogleAppsScript.Spreadsheet.RichTextValue | null
    condition: string
  }> = []
  private subject: string = ''
  private html: string = ''

  constructor (private params: EmailTemplateParams) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(params.sheetName)
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        `Tentative d'accès à la feuille inexistante "${params.sheetName}"`
      )
      return
    }
    this.textEntries = new Array()
    for (let index = 2; index < sheet.getLastRow(); index++) {
      this.textEntries.push({
        type: sheet.getRange(index, 1).getValue(),
        condition: sheet.getRange(index, 2).getValue(),
        richText: sheet.getRange(index, 3).getRichTextValue()
      })
    }
  }

  public addValidation () {
    const validationHandler = ValidationHandler.getInstance()
    validationHandler.add(
      new Validation(this.params.sheetName, 'A2:A', '', '', false, [
        'sujet',
        'titre',
        'sous-titre',
        'paragraphe',
        'élément de liste',
        'aucun'
      ])
    )
    validationHandler.add(
      new Validation(
        this.params.sheetName,
        'B2:B',
        '',
        '',
        false,
        this.params.conditions
      )
    )
  }

  public constructHtml (data: any) {
    let html = ''
    let subject = ''
    let listContext = false
    for (const entry of this.textEntries) {
      const conditionOk = this.params.evaluateCondition(entry.condition, data)
      switch (entry.type) {
        case 'sujet':
          subject = entry.richText?.getText() ?? '' // no rich text handling for subject
          break
        case 'aucun':
          html += listContext ? '</ul>' : ''
          listContext = false
          if (conditionOk) html += entry.richText?.getText() // no rich text handling for 'aucun'
          break
        case 'paragraphe':
          html += listContext ? '</ul>' : ''
          listContext = false
          if (conditionOk)
            html += `<p>${
              entry.richText
                ? this.richTextToHtml(entry.richText)
                : 'Paragraphe'
            }</p>`
          break
        case 'titre':
          html += listContext ? '</ul>' : ''
          listContext = false
          if (conditionOk)
            html += `<h1>${entry.richText?.getText() ?? 'Titre 1'}</h1>` // no rich text handling for title
          break
        case 'sous-titre':
          html += listContext ? '</ul>' : ''
          listContext = false
          if (conditionOk)
            html += `<h2>${entry.richText?.getText() ?? 'Titre 2'}</h2>` // no rich text handling for subtitle
          break
        case 'élément de liste':
          if (conditionOk) {
            html += listContext ? '' : '<ul>'
            listContext = true
            html += `<li>${
              entry.richText
                ? this.richTextToHtml(entry.richText)
                : 'Elément de liste'
            }</li>`
          }
          break
        default:
          break
      }
    }
    return { subject, html }
  }

  public insertData (html: string, data: any): string {
    return this.params.insertData(html, data)
  }

  private richTextToHtml (
    richText: GoogleAppsScript.Spreadsheet.RichTextValue
  ): string {
    const getRunAsHtml = (
      richTextRun: GoogleAppsScript.Spreadsheet.RichTextValue
    ) => {
      const richText = richTextRun.getText()

      // Returns the rendered style of text in a cell.
      const style = richTextRun.getTextStyle()

      // Returns the link URL, or null if there is no link
      // or if there are multiple different links.
      const url = richTextRun.getLinkUrl()

      const styles: any = {
        color: style.getForegroundColor(),
        'font-family': style.getFontFamily(),
        'font-size': `${style.getFontSize()}pt`,
        'font-weight': style.isBold() ? 'bold' : '',
        'font-style': style.isItalic() ? 'italic' : '',
        'text-decoration': style.isUnderline() ? 'underline' : ''
      }

      // Gets whether or not the cell has strike-through.
      if (style.isStrikethrough()) {
        styles['text-decoration'] = `${styles['text-decoration']} line-through`
      }

      const css = Object.keys(styles)
        .filter(attr => styles[attr])
        .map(attr => [attr, styles[attr]].join(':'))
        .join(';')

      const styledText = `<span style='${css}'>${richText}</span>`
      return url ? `<a href='${url}'>${styledText}</a>` : styledText
    }

    /* Returns the Rich Text string split into an array of runs,
        wherein each run is the longest possible
        substring having a consistent text style. */
    const runs = richText.getRuns()

    return runs.map(run => getRunAsHtml(run)).join('')
  }
}
