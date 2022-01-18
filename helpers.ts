export function rangeIntersect (
  r1: GoogleAppsScript.Spreadsheet.Range | undefined,
  r2: GoogleAppsScript.Spreadsheet.Range | undefined
): boolean {
  if (!r1 || !r2) return false
  let sheetMatches = r1.getSheet().getName() == r2.getSheet().getName()
  let rangeIntersects =
    r1.getLastRow() >= r2.getRow() &&
    r2.getLastRow() >= r1.getRow() &&
    r1.getLastColumn() >= r2.getColumn() &&
    r2.getLastColumn() >= r1.getColumn()
  return sheetMatches && rangeIntersects
}

export function rowHasContent (row: Array<string>) {
  return row.join('').length > 0
}

export function rowHasContentInColumn (index: number) {
  return (row: Array<string>) => row[index].length > 0
}

export function compareRowsOnColumn (index: number) {
  return (a: Array<string>, b: Array<string>) => (a[index] > b[index] ? 1 : -1)
}

export function rowHasValue (index: number, value: string) {
  return (row: Array<string>) => row[index] === value
}

export function getColumnAsRow (index: number) {
  return (row: Array<string>) => row[index]
}

export function getColumn (index: number) {
  return (row: Array<string>) => [row[index]]
}

export function searchReplace (oldValue: string, newValue: string) {
  return (row: Array<string>) => row.map(x => (x === oldValue ? newValue : x))
}

export function getDriveId (sharingLink: string): string {
  const result = sharingLink.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
  if (!result) throw new Error(`No drive ID could be found in ${sharingLink}`)
  return result[0]
}

// function durationToDecimalHours(d: Duration): number {
//   if (d.years) throw new Error("Function durationToDecimalHours cannot be used if duration contain years (number of days in year varies)!")
//   const daysInWeek = 7
//   const hoursInDay = 24
//   const minutesInHour = 60
//   const secondsInMinute = 60
//   const weeksHours = d.weeks ? d.weeks * daysInWeek * hoursInDay : 0
//   const daysHours = d.days ? d.days * hoursInDay : 0
//   const hoursHours = d.hours ? d.hours : 0
//   const minutesHours = d.minutes ? d.minutes / minutesInHour : 0
//   const secondsHours = d.seconds ? d.seconds / minutesInHour / secondsInMinute : 0
//   return weeksHours + daysHours + hoursHours + minutesHours + secondsHours
// }
