import ExcelJS from "exceljs"

//will refactor someday... someday...
import { grantham_start_year } from "./main.js"
import { dveTeacherMarkingSchema } from "./schema.js"

const _singleBorder = { style: "medium", color: { argb: "FF000000" } }

const excelStyle = {
  align: {
    vertical: "middle",
    horizontal: "center",
  },
  border: {
    full: ["top", "left", "bottom", "right"].reduce((acc, key) => {
      acc[key] = _singleBorder
      return acc
    }, {}),
    thin: ["top", "left", "bottom", "right"].reduce((acc, key) => {
      acc[key] = { style: "thin", color: { argb: "808080" } }
      return acc
    }, {}),
  },
  fontSize: {
    table: 10,
  },
}

const bodyStyle = (cell, thin) => {
  cell.border = thin ? excelStyle.border.thin : excelStyle.border.full
  cell.font = {
    size: excelStyle.fontSize.table,
  }
  cell.alignment = {
    wrapText: true,
    vertical: "top",
  }
}

const bodyWarningStyle1 = (cell, thin) => {
  bodyStyle(cell, thin)
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFE066" },
  }
}
const bodyWarningStyle2 = (cell, thin) => {
  bodyStyle(cell, thin)
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFC107" },
  }
}

const headStyle = (cell, thin) => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" },
  }
  cell.alignment = {
    ...cell.alignment,
    wrapText: true,
  }
  cell.font = {
    size: excelStyle.fontSize.table,
  }
  cell.border = thin ? excelStyle.border.thin : excelStyle.border.full
}

const systemStyle = (cell, header, thin) => {
  headStyle(cell, thin)
  cell.fill.fgColor = { argb: "FFDCE6F1" }
  if (!header)
    cell.alignment = {
      wrapText: true,
      vertical: "top",
    }
}

const warningStyle = (cell, header, thin) => {
  systemStyle(cell, header, thin)
  cell.fill.fgColor = { argb: "FFFFFFCC" }
}

const errorStyle = (cell, header, thin) => {
  systemStyle(cell, header, thin)
  cell.fill.fgColor = { argb: "FFFCD5B4" }
}

//#D8BFD8

excelStyle.table = {
  head: (cell, header, thin) => headStyle(cell, thin),
  error: (cell, header, thin) => errorStyle(cell, header, thin),
  warning: (cell, header, thin) => warningStyle(cell, header, thin),
  system: (cell, header, thin) => systemStyle(cell, header, thin),
  body: (cell, header, thin) => bodyStyle(cell, thin),
  warning1: (cell, header, thin) => bodyWarningStyle1(cell, thin),
  warning2: (cell, header, thin) => bodyWarningStyle2(cell, thin),
}

function sanitizeWorksheetName(name) {
  return name
    .replace(/[\\/]/g, "-")
    .replace(/[*?:\[\]]/g, "")
    .substring(0, 31)
}

export function useExcelGenerator(
  schema,
  staticStyle,
  freezeX = 0,
  thin = false,
  rowHeight = 15
) {
  const columns = schema.map((column) => ({ ...column }))
  const stylesIn = columns.map((col) => col._style || null)
  let totalRecords = 0
  // let recordNum = 0

  function stylizeRow(row, style, height, header = false) {
    totalRecords++

    row.height = height
    row.alignment = { vertical: "middle", horizontal: "center" }
    let rowStyle = style || "body"

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      let colStyle = rowStyle
      if (header || !style) colStyle = stylesIn?.[colNumber - 1] || colStyle
      if (colNumber <= columns.length) {
        excelStyle.table[colStyle](cell, header, thin)
      }
    })
  }

  function parseRecord(worksheet, row, style) {
    // let corrected = false
    const rowValues = columns.reduce((acc, col, index) => {
      let value = col.get(row)
      acc[col.key] = value
      //   if (col.__style && value) {
      //     styles[index] = col._style
      //     corrected = true
      //   } else {
      //     styles[index] = stylesIn?.[index] || undefined
      //   }
      return acc
    }, {})
    // if (corrected) console.log('rowValues', styles)
    // corrected = false
    stylizeRow(worksheet.addRow(rowValues), staticStyle || style, rowHeight)
  }

  let warning = true
  function parseRecords(worksheet, rows) {
    for (const [key, records] of rows) {
      // console.log('key', key, records.length)
      if (records.length > 1) {
        for (const r of records) {
          parseRecord(
            worksheet,
            r,
            key == "no id" ? undefined : warning ? "warning1" : "warning2"
          )
        }
        warning = !warning
        // recordNum += records.length
        continue
      }
      // recordNum++
      let record = records[0] || records
      parseRecord(
        worksheet,
        record,
        record.__overwritten
          ? key == "no id"
            ? undefined
            : warning
            ? "warning1"
            : "warning2"
          : undefined
      )
      if (record.__overwritten) {
        warning = !warning
      }
    }
  }

  return (data) => {
    const workbook = new ExcelJS.Workbook()
    totalRecords = 0

    Object.keys(data).map((key) => {
      const sanitizedKey = sanitizeWorksheetName(key)
      // console.log('sanitizedKey', sanitizedKey)
      const records = data[key]
      const worksheet = workbook.addWorksheet(sanitizedKey)

      worksheet.columns = columns

      let header = worksheet.getRow(1)
      stylizeRow(header, "head", thin ? 40 : 50, true)

      //console.log('columns', columns, records)
      worksheet.views = [
        { state: "frozen", ySplit: 1, ...(freezeX ? { xSplit: freezeX } : {}) },
      ]

      const lastColumn = worksheet.getColumn(worksheet.columnCount).letter
      worksheet.autoFilter = {
        from: "A1",
        to: `${lastColumn}1`,
      }

      parseRecords(worksheet, records)
      // console.log('TotalRecords', totalRecords, recordNum, records.size)
    })
    // console.log(stylesIn)
    return workbook
  }
}

export function generateTeacherWorkbook(param, rows) {
  const workbook = new ExcelJS.Workbook()

  const granthamYear = `${grantham_start_year + 4}-${grantham_start_year + 5}`
  const title = `葛量洪獎學金 ${granthamYear} \n學生表現評核表`

  const startColumn = 2
  const worksheet = workbook.addWorksheet(title)
  worksheet.properties.defaultRowHeight = 20
  worksheet.columns = dveTeacherMarkingSchema.map((v) => ({
    ...v,
    key: v.name,
  }))

  const centered = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  }

  const fillRow = (i) => {
    const row = worksheet.getRow(i)
    const cell = row.getCell(1)

    //to solve weird behavior
    row.fill = cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF" },
    }
  }
  const fillRows = (i, num) => {
    for (let k = 0; k < num; k++) fillRow(i + k)
  }

  const fill = (blue) => ({
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: blue ? "44B3E1" : "D86DCD" },
  })

  const borderStyle = {
    top: { style: "medium", color: { argb: "00000000" } },
    left: { style: "medium", color: { argb: "00000000" } },
    bottom: { style: "medium", color: { argb: "00000000" } },
    right: { style: "medium", color: { argb: "00000000" } },
  }

  const contact = [
    "",
    "職業發展計劃辦事處",
    param.address,
    `電話：${param.phone}   傳真：${param.fax}`,
    "",
  ]

  let i = 0,
    cell,
    row

  for (; i < contact.length; i++) {
    row = worksheet.getRow(i + 1)
    row.height = 15
    row.alignment = {
      vertical: "middle",
    }
    row.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "D9D9D9" },
    }

    cell = row.getCell(startColumn)
    cell.value = contact[i] ?? ""
    cell.font = {
      size: 8,
      bold: true,
      color: { argb: "595959" },
    }
  }

  i++

  row = worksheet.getRow(i)
  row.height = 40

  i++

  worksheet.mergeCells(i, startColumn, (i += 3), startColumn + 3)

  cell = worksheet.getRow(i).getCell(startColumn)
  cell.value = title
  cell.font = {
    size: 20,
    bold: true,
  }
  cell.alignment = { wrapText: true, vertical: "middle" }

  i += 2

  worksheet.mergeCells(i, startColumn, (i += 2), startColumn + 3)

  cell = worksheet.getRow(i).getCell(startColumn)
  cell.value =
    "為鼓勵「Teen才再現」畢業同學繼續努力學習，本辦事處特申請葛量洪獎學金，以鼓勵升學後保持良好學習態度的畢業同學。現懇請　閣下為班內之「Teen才再現」畢業生給予意見，以便評核委員會作評選。"
  cell.alignment = {
    vertical: "top",
    horizontal: "left",
    wrapText: true,
  }

  i += 2

  worksheet.mergeCells(i, startColumn, i, startColumn + 3)

  cell = worksheet.getRow(i).getCell(startColumn)
  cell.value = "*請導師為每位學生填寫下表，謝謝!"
  cell.font = {
    bold: true,
    size: 8,
  }
  cell.alignment = {
    vertical: "top",
    horizontal: "left",
    wrapText: false,
  }

  i += 2

  row = worksheet.getRow(i)
  row.height = 140
  row.alignment = centered

  let markingCellIndex, excelEndColumn

  fillRows(contact.length + 1, 15)

  // worksheet.getRow(11).fill = {
  //   type: "pattern",
  //   pattern: "solid",
  //   fgColor: { argb: "FFFF00" },
  // }

  dveTeacherMarkingSchema.forEach((value, x) => {
    if (!x) return

    excelEndColumn = startColumn + x - 1 //for the spacer
    cell = row.getCell(excelEndColumn)

    if (value.blue && markingCellIndex === undefined) {
      markingCellIndex = excelEndColumn
      worksheet.mergeCells(i - 1, markingCellIndex, i, markingCellIndex)
      cell.alignment = centered
    }

    cell.value = value.name
    cell.width = value.width

    cell.border = borderStyle
    cell.fill = fill(value.blue)
  })

  const recordStartRow = i + 1
  const markingItemCount = excelEndColumn - markingCellIndex

  i--

  row = worksheet.getRow(i)
  row.height = 20
  row.alignment = centered

  for (let y = 0; y < markingItemCount; y++) {
    cell = row.getCell(y + markingCellIndex + 1)

    cell.value = `項目${y + 1}`
    cell.border = borderStyle
    cell.fill = fill(true)
  }

  i--

  row = worksheet.getRow(i)
  row.height = 40

  cell = row.getCell(markingCellIndex)
  cell.value = "100%為上限\n( 由開學直至今)"
  cell.border = borderStyle
  cell.fill = fill(true)
  cell.alignment = centered

  worksheet.mergeCells(i, markingCellIndex + 1, i, excelEndColumn)

  cell = row.getCell(markingCellIndex + 1)
  cell.value =
    "學生表現\n(10=優             7=良             5=常              3=可              1=劣)"
  cell.border = borderStyle
  cell.fill = fill(true)
  cell.alignment = centered

  i--

  row = worksheet.getRow(i)
  row.height = 20

  worksheet.mergeCells(i, markingCellIndex, i, excelEndColumn)

  cell = row.getCell(markingCellIndex)
  cell.value = "學生表現核表"
  cell.border = borderStyle
  cell.fill = fill(true)
  cell.alignment = centered

  i = recordStartRow

  const records = [
    ...samples,
    ...rows,
    ...Array.from({ length: 1 }).map(() => ({ 0: "_" })),
  ]

  fillRows(i, records.length + 1)

  for (const record of records) {
    row = worksheet.getRow(i)

    const isLastRow = i - recordStartRow === records.length - 1

    row.height = i == recordStartRow + 1 ? 40 : isLastRow ? 40 : 30

    row.alignment = {
      vertical: "middle",
    }

    let isblank = false
    let isSample = false

    for (let x = startColumn; x < excelEndColumn + 1; x++) {
      cell = row.getCell(x)

      if (record[x - startColumn] === "_") {
        cell.value = ""
        isblank = true
      } else {
        cell.value = record[x - startColumn] ?? ""
      }

      if (`${cell.value}`.startsWith("(例子")) isSample = true

      if (x >= markingCellIndex - 1) {
        cell.alignment = centered
      }

      cell.fill =
        x >= markingCellIndex
          ? isLastRow
            ? { type: "pattern", pattern: "solid", fgColor: { argb: "B7DEE8" } }
            : fill(true)
          : isblank
          ? null
          : {
              type: "pattern",
              pattern: "solid",
              fgColor: {
                argb: isSample ? "D9D9D9" : "F2F2F2",
              },
            }

      cell.border = {
        ...borderStyle,
        top: isLastRow
          ? null
          : i == recordStartRow || x >= markingCellIndex
          ? borderStyle.top
          : { style: "thin", color: { argb: "808080" } },
        bottom:
          i - recordStartRow === records.length - 2
            ? null
            : i - recordStartRow === records.length - 1 || x >= markingCellIndex
            ? borderStyle.bottom
            : { style: "thin", color: { argb: "808080" } },
      }
    }

    i++
  }

  i++

  const footNote = ["", `*請於${param.deadline}或之前用電郵直接繳交。`, "", ""]

  fillRows(i, footNote.length + 10)

  for (let x = 0; x < footNote.length; x++) {
    row = worksheet.getRow(i + 1 + x)
    row.height = footNote[x] ? 20 : 15
    row.alignment = {
      vertical: "middle",
    }

    cell = row.getCell(startColumn)
    cell.value = footNote[x] ?? ""
    cell.font = {
      bold: true,
      color: { argb: "262626" },
    }
  }

  row.border = {
    bottom: borderStyle.bottom,
  }

  return workbook
}

const sampleRichText_No = {
  richText: [
    {
      text: "(不推薦→不用填寫)",
      font: { color: { argb: "7E350E" } },
    },
  ],
}

const samples = [
  {
    0: "(例子1) 陳大文",
    1: "FSXXXXXX",
    2: "3X 班",
    3: "XXXXX",
    4: "XXXXXX先生",
    5: "XXXXXXXX",
    6: "XXXXX@vtc.edu.hk",
    7: {
      richText: [
        { text: "推薦 ", font: { color: { argb: "000000" } } },
        {
          text: "(請繼續填寫籃色部份內容)",
          font: { color: { argb: "FF0000" } },
        },
      ],
    },
    8: 80.56,
    9: 8,
    10: 9,
    11: 10,
    12: 8,
    13: 7,
    14: 6,
  },
  {
    0: "(例子2) 張大明",
    1: "FSXXXXXX",
    2: "B1X 班",
    3: "XXXXX",
    4: "XXXXXX小姐",
    5: "XXXXXXXX",
    6: "XXXXX@vtc.edu.hk",
    7: "不推薦, 原因：學生態度懶散",
    8: sampleRichText_No,
    9: sampleRichText_No,
    10: sampleRichText_No,
    11: sampleRichText_No,
    12: sampleRichText_No,
    13: sampleRichText_No,
    14: sampleRichText_No,
  },
]
