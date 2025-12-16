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

const generateDate = (date = new Date()) => {
  const chiNumbers = ["日", "一", "二", "三", "四", "五", "六"]

  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, "0")
  const day = String(date.getDate()).padStart(2, "0")
  const weekday = chiNumbers[date.getDay()]

  return `${year}年${month}月${day}日（星期${weekday}）`
}

const _argb = (hexC) => ({ argb: hexC.replace("#", "") })
const _fill = (color) => ({
  type: "pattern",
  pattern: "solid",
  fgColor: _argb(color),
})

export function generateTeacherWorkbook(param, rows) {
  const workbook = new ExcelJS.Workbook()

  const granthamYear = `${grantham_start_year + 4}-${grantham_start_year + 5}`
  const title = `葛量洪獎學金 ${granthamYear} \n學生表現評核表`

  const startColumn = 2
  const worksheet = workbook.addWorksheet(title, {
    views: [
      {
        zoomScale: 80,
      },
    ],
  })

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
  const bottomCentered = {
    vertical: "bottom",
    horizontal: "center",
    wrapText: true,
  }
  const TopCentered = {
    vertical: "top",
    horizontal: "center",
    wrapText: true,
  }
  const _noWrap = (t) => ({ ...t, wrapText: false })

  const bgFill = _fill("#FFFFFF")
  const contactFill = _fill("#D9D9D9")

  const noHeaderFill = _fill("#963634")
  const yesHeaderFill = _fill("#76933C")
  const yesSubHeaderNCellFill = _fill("#C4D79B")

  const normalRowFill = _fill("#F2F2F2")
  const sampleRowFill = _fill("#D9D9D9")
  const teacherInfoHeaderFill = _fill("#404040")
  const prefilledHeaderFill = _fill("#C5D9F1")

  const headerFont = {
    size: 12,
    bold: true,
    color: _argb("#000000"),
  }
  const headerFontReverse = { ...headerFont, color: _argb("#FFFFFF") }
  const headerFontSmall = { ...headerFont, size: 11 }
  const yesSubHeaderItemFont = { ...headerFontSmall, color: _argb("#4F6228") }
  const headerFontReverseBig = { ...headerFontReverse, size: 20 }

  const contactFont = {
    size: 8,
    bold: true,
    color: _argb("#595959"),
  }

  const fillRow = (i) => {
    const row = worksheet.getRow(i)
    const cell = row.getCell(1)

    //to solve weird behavior
    row.fill = cell.fill = bgFill
  }

  const fillRows = (i, num) => {
    for (let k = 0; k < num; k++) fillRow(i + k)
  }

  const bs = {
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
    cell.font = contactFont
  }

  i++

  row = worksheet.getRow(i)
  row.height = 40

  i++

  worksheet.mergeCells(i, startColumn, (i += 3), startColumn + 3)
  for (let j = i - 4; j < i; j++) worksheet.getRow(i).height = 25 //!check row 7-10

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

  i += 1

  row = worksheet.getRow(i)
  row.height = 55 //!check row 17
  const nextRow = worksheet.getRow(i + 1)
  nextRow.height = 100 //!check row 18

  let markingCellIndex, excelEndColumn

  fillRows(contact.length + 1, 15)

  //filling header

  const headerStyles = {
    blue: {
      font: headerFont,
      fill: prefilledHeaderFill,
      alignment: centered,
    },
    gray: {
      font: headerFontReverse,
      fill: teacherInfoHeaderFill,
      alignment: centered,
    },
    no: {
      font: headerFontReverseBig,
      fill: noHeaderFill,
      alignment: centered,
    },
    yes: {
      font: headerFontSmall,
      fill: yesSubHeaderNCellFill,
      alignment: centered,
    },
    yesItem: {
      font: yesSubHeaderItemFont,
      fill: yesSubHeaderNCellFill,
      alignment: bottomCentered,
    },
    yesBanner: {
      font: headerFontReverseBig,
      fill: yesHeaderFill,
      alignment: centered,
    },
  }

  dveTeacherMarkingSchema.forEach((value, x) => {
    if (!x) return

    excelEndColumn = startColumn + x - 1 //for the spacer
    cell = row.getCell(excelEndColumn)

    if (value.type == "yes") {
      if (markingCellIndex === undefined) {
        markingCellIndex = excelEndColumn
        worksheet.mergeCells(i - 1, markingCellIndex, i, markingCellIndex)

        cell.value = "100%為上限\n(由開學直至今)"
        cell.border = { ...bs, top: undefined, bottom: undefined }
      } else {
        cell.value = `項目${excelEndColumn - markingCellIndex}`
        cell.border = { ...bs, bottom: undefined }

        if (x === dveTeacherMarkingSchema.length - 1) {
          worksheet.mergeCells(
            i - 1,
            markingCellIndex + 1,
            i - 1,
            excelEndColumn
          )

          const tempRow = worksheet.getRow(i - 1)
          tempRow.height = 55 //!check row 16

          let tempCell = tempRow.getCell(markingCellIndex + 1)
          tempCell.value =
            "學生表現\n(10=優             7=良             5=常              3=可              1=劣)"

          Object.assign(tempCell, headerStyles.yes)
          tempCell.border = { ...bs, top: undefined }

          worksheet.mergeCells(
            i - 1 - 5,
            markingCellIndex,
            i - 2,
            excelEndColumn
          )
          tempCell = worksheet.getRow(i - 1 - 5).getCell(markingCellIndex)
          tempCell.value = `推薦 (請填寫以下部份，包括出席率及項目1-${
            excelEndColumn - markingCellIndex
          })`
          Object.assign(tempCell, headerStyles.yesBanner)
          tempCell.border = { ...bs, bottom: undefined }
        }
      }
      Object.assign(
        cell,
        markingCellIndex == excelEndColumn
          ? headerStyles.yes
          : headerStyles.yesItem
      )

      cell = nextRow.getCell(excelEndColumn)
      cell.border = { ...bs, top: undefined }
    } else {
      worksheet.mergeCells(i, excelEndColumn, i + 1, excelEndColumn)
      cell.border = bs
    }

    cell.value = value.name

    Object.assign(cell, headerStyles[value.type] || {})
  })

  i += 1

  const recordStartRow = i + 1

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

      cell.fill = isSample
        ? sampleRowFill
        : x >= markingCellIndex
        ? yesSubHeaderNCellFill
        : isblank
        ? null
        : normalRowFill

      cell.border = {
        ...bs,
        top: isLastRow
          ? null
          : i == recordStartRow // || x >= markingCellIndex
          ? bs.top
          : { style: "thin", color: { argb: "808080" } },
        bottom:
          i - recordStartRow === records.length - 2
            ? null
            : i - recordStartRow === records.length - 1 // || x >= markingCellIndex
            ? bs.bottom
            : { style: "thin", color: { argb: "808080" } },
      }
    }

    i++
  }

  const deadline = generateDate(param.deadline)
  const footNote = ["", `*請於${deadline}或之前用電郵直接繳交。`, "", "", ""]

  fillRows(i, footNote.length + 1)

  for (let x = 0; x < footNote.length; x++) {
    row = worksheet.getRow(i + 1 + x)
    // row.height = footNote[x] ? 20 : 15
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
    bottom: bs.bottom,
  }

  return workbook
}

const cellImportantFont = {
  size: 11,
  bold: true,
  color: _argb("#C00000"), //FF0000
}

const sampleRichText_No = {
  richText: [
    {
      text: "(不推薦→不用填寫)",
      font: cellImportantFont,
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
      richText: [{ text: "(推薦→不用填寫) ", font: cellImportantFont }],
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
