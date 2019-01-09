var fs = require('fs')
var XLSX = require('xlsx')
var _ = require('lodash')
var ical = require('ical-generator')

var cal = ical({ domain: 'road hammers', name: 'running' })

var buf = fs.readFileSync('input.xlsx')
var wb = XLSX.read(buf, { type: 'buffer' })

var workoutSheet = wb.Sheets['2019']
console.log(workoutSheet)

var planRows = [9, 10, 11, 12, 13, 14, 15]
var dateMap = {
  'Jan': 0,
  'Feb': 1,
  'Mar': 2,
  'Apr': 3,
  'May': 4,
  'Jun': 5,
  'Jul': 6,
  'Aug': 7,
  'Sep': 8,
  'Oct': 9,
  'Nov': 10,
  'Dec': 11
}

var events = []

_.each(planRows, (row) => {
  var date = workoutSheet['A' + row]
  var dateRegex = /([a-zA-Z]+)\s+(\d+)\s?\-.*/g

  var match = dateRegex.exec(date.w)
  var month = match[1]
  var startDate = match[2]

  var days = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
  _.each(days, (day, idx) => {
    var kms = workoutSheet[day + row]
    cal.createEvent({
      start: new Date(2019, dateMap[month], parseInt(startDate, 10) + parseInt(idx, 10)),
      allDay: true,
      summary: kms.w
    })
  })
})

fs.writeFileSync('plan.ics', cal.toString())
