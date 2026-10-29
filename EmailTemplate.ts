import { escapeHtml, textToHtml } from './helpers'
import { Validation } from './Validation'
import { ValidationHandler } from './ValidationHandler'

export interface EmailTemplateParams {
  sheetName: string
  conditions: string[]
  /** Substitutes the markers in the HTML body. Values must be HTML-escaped. */
  insertData(html: string, data: any): string
  /** Substitutes the markers in the plain-text subject (optional). */
  insertSubjectData?(subject: string, data: any): string
  evaluateCondition(condition: string, data: any): boolean
}

/** One row of a template sheet. */
export interface TemplateEntry {
  type: string
  condition: string
  /** The cell text, whatever the cell type (text, number, date…). */
  text: string
  /** The rich text, or null when the cell does not hold text. */
  richText: GoogleAppsScript.Spreadsheet.RichTextValue | null
}

export const TEMPLATE_TYPES = [
  'sujet',
  'titre',
  'sous-titre',
  'paragraphe',
  'élément de liste',
  'aucun'
]

export class EmailTemplate {
  private textEntries: TemplateEntry[] = []

  constructor (private params: EmailTemplateParams) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(params.sheetName)
    if (!sheet) {
      // No UI call here: this may run from a trigger. Callers check isEmpty().
      console.warn(`Feuille modèle "${params.sheetName}" introuvable`)
      return
    }
    // Rows 2..lastRow hold the template; read them in two calls instead of
    // three per row.
    const numRows = sheet.getLastRow() - 1
    if (numRows <= 0) return
    const range = sheet.getRange(2, 1, numRows, 3)
    const values = range.getValues()
    const richTexts = range.getRichTextValues()
    for (let index = 0; index < numRows; index++) {
      this.textEntries.push({
        type: String(values[index][0]).trim(),
        condition: String(values[index][1]).trim(),
        text: String(values[index][2]),
        richText: richTexts[index][2]
      })
    }
  }

  /** True when no template row could be read (missing sheet or only a header). */
  public isEmpty (): boolean {
    return this.textEntries.length === 0
  }

  /** Registers the dropdowns of a template sheet without reading it. */
  public static addValidation (params: EmailTemplateParams) {
    const validationHandler = ValidationHandler.getInstance()
    validationHandler.add(
      new Validation(params.sheetName, 'A2:A', '', '', false, TEMPLATE_TYPES)
    )
    validationHandler.add(
      new Validation(params.sheetName, 'B2:B', '', '', false, params.conditions)
    )
  }

  public addValidation () {
    EmailTemplate.addValidation(this.params)
  }

  public constructHtml (data: any) {
    let html = ''
    let subject = ''
    let listContext = false
    const closeList = () => {
      html += listContext ? '</ul>' : ''
      listContext = false
    }
    for (const entry of this.textEntries) {
      const conditionOk = this.params.evaluateCondition(entry.condition, data)
      switch (entry.type) {
        case 'sujet':
          if (conditionOk) subject = entry.text // plain text, no rich text handling
          break
        case 'aucun':
          closeList()
          // Raw pass-through: the only way to put custom HTML in the mail.
          if (conditionOk) html += entry.text
          break
        case 'paragraphe':
          closeList()
          if (conditionOk) html += `<p>${this.entryToHtml(entry)}</p>`
          break
        case 'titre':
          closeList()
          if (conditionOk) html += `<h1>${textToHtml(entry.text)}</h1>`
          break
        case 'sous-titre':
          closeList()
          if (conditionOk) html += `<h2>${textToHtml(entry.text)}</h2>`
          break
        case 'élément de liste':
          if (conditionOk) {
            html += listContext ? '' : '<ul>'
            listContext = true
            html += `<li>${this.entryToHtml(entry)}</li>`
          }
          break
        default:
          break
      }
    }
    closeList()
    return { subject, html }
  }

  public insertData (html: string, data: any): string {
    return this.params.insertData(html, data)
  }

  public insertSubjectData (subject: string, data: any): string {
    return this.params.insertSubjectData
      ? this.params.insertSubjectData(subject, data)
      : subject
  }

  /** Rich text when the cell holds text, escaped plain text otherwise. */
  private entryToHtml (entry: TemplateEntry): string {
    return entry.richText
      ? this.richTextToHtml(entry.richText)
      : textToHtml(entry.text)
  }

  private richTextToHtml (
    richText: GoogleAppsScript.Spreadsheet.RichTextValue
  ): string {
    const getRunAsHtml = (
      richTextRun: GoogleAppsScript.Spreadsheet.RichTextValue
    ) => {
      const text = textToHtml(richTextRun.getText())

      // Returns the rendered style of text in a cell.
      const style = richTextRun.getTextStyle()

      // Returns the link URL, or null if there is no link
      // or if there are multiple different links.
      const url = richTextRun.getLinkUrl()

      const styles: any = {
        color: style.getForegroundColor(),
        'font-family': style.getFontFamily(),
        // getFontSize() is null when the run has no explicit size.
        'font-size': style.getFontSize() ? `${style.getFontSize()}pt` : '',
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

      const styledText = `<span style='${css}'>${text}</span>`
      return url ? `<a href="${escapeHtml(url)}">${styledText}</a>` : styledText
    }

    /* Returns the Rich Text string split into an array of runs,
        wherein each run is the longest possible
        substring having a consistent text style. */
    const runs = richText.getRuns()

    return runs.map(run => getRunAsHtml(run)).join('')
  }
}
