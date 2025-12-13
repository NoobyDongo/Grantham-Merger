import ExcelJS from "exceljs"

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
  },
  fontSize: {
    table: 10,
  },
}

const bodyStyle = (cell) => {
  cell.border = excelStyle.border.full
  cell.font = {
    size: excelStyle.fontSize.table,
  }
  cell.alignment = {
    wrapText: true,
    vertical: "top",
  }
}

const bodyWarningStyle1 = (cell) => {
  bodyStyle(cell)
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFE066" },
  }
}
const bodyWarningStyle2 = (cell) => {
  bodyStyle(cell)
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFC107" },
  }
}

const headStyle = (cell) => {
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
  cell.border = excelStyle.border.full
}

const systemStyle = (cell, header) => {
  headStyle(cell)
  cell.fill.fgColor = { argb: "FFDCE6F1" }
  if (!header)
    cell.alignment = {
      wrapText: true,
      vertical: "top",
    }
}

const warningStyle = (cell, header) => {
  systemStyle(cell, header)
  cell.fill.fgColor = { argb: "FFFFFFCC" }
}

const errorStyle = (cell, header) => {
  systemStyle(cell, header)
  cell.fill.fgColor = { argb: "FFFCD5B4" }
}

//#D8BFD8

excelStyle.table = {
  head: (cell) => headStyle(cell),
  error: (cell, header) => errorStyle(cell, header),
  warning: (cell, header) => warningStyle(cell, header),
  system: (cell, header) => systemStyle(cell, header),
  body: (cell) => bodyStyle(cell),
  warning1: (cell) => bodyWarningStyle1(cell),
  warning2: (cell) => bodyWarningStyle2(cell),
}

function sanitizeWorksheetName(name) {
  return name
    .replace(/[\\/]/g, "-")
    .replace(/[*?:\[\]]/g, "")
    .substring(0, 31)
}

export function useExcelGenerator(schema, staticStyle) {
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
        excelStyle.table[colStyle](cell, header)
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
    stylizeRow(worksheet.addRow(rowValues), staticStyle || style, 15)
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
      stylizeRow(header, "head", 50, true)

      //console.log('columns', columns, records)
      worksheet.views = [{ state: "frozen", ySplit: 1 }]

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
