XLSX = require('xlsx')
fs = require('fs')

function readSheet(sheet) {
const rows = []

let mattRow

let dateStr = sheet['B1'].v
// dateStr = getDateFromString()
console.log(dateStr)
// Get month and first number

let [ month, dayRange ] = dateStr
dayRange = dayRange.split('-')

const colToDay = {
C: { name: 'Monday' },
D: { name: 'Tuesday' },
E: { name: 'Wednesday' },
F: { name: 'Thursday' },
G: { name: 'Friday' },
H: { name: 'Saturday' },
I: { name: 'Sunday' }
}

Object.values(colToDay).forEach((day, i) => {
day.date = dateStr
// day.date = i + parseInt(dayRange[0])
})

Object.keys(sheet).forEach((key) => {
let cell = sheet[key].v
let col = key.replace(/[0-9]+/, '')
let row = key.replace(/[A-Z]+/, '')

if (cell !== 'x') {
if (/Matt(\s\(\$\))?/.test(cell)) {
if (mattRow) throw new Error('More than one Matt exists.')
mattRow = row
} else if (row === mattRow) {
console.log(colToDay[col], cell)
}
}
})
}

let sheets

fs.readdir(`${__dirname}/attachments`, (err, files) => {
if (err) throw new Error(err)
if (files.length < 1) console.log('No files to read!')
files = files.filter((v) => v !== '.DS_Store')
files.forEach((fileName) => {
sheet = XLSX.readFile(`${__dirname}/attachments/${fileName}`).Sheets.Sheet1
readSheet(sheet)
})
})
