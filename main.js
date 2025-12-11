import * as readline from "readline"
import * as xlsx from "xlsx"
import fs from "fs"
import path from "path"
import { fileURLToPath } from "url"
import {
  checker,
  exportSchema,
  header,
  headerRemakeNames,
  remarkRegex,
  solver,
} from "./schema.js"
import { useExcelGenerator } from "./assembler.js"

///=================================================================================================
//
// log utils
//
///=================================================================================================

export const customColor = (colorCode) => (text) => {
  return `\x1b[38;5;${colorCode}m${text}\x1b[0m`
}

const green = customColor(2)
const red = customColor(1)
const orange = customColor(208)
const lightBlue = customColor(12)

///=================================================================================================
//
// params
//
///=================================================================================================

//========//exe config
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const isPkg = typeof process.pkg !== "undefined"
export const baseDir = isPkg ? path.dirname(process.execPath) : __dirname

//========//excel config
const excelDir = path.join(baseDir, "_put your excels here")
if (!fs.existsSync(excelDir)) {
  fs.mkdirSync(excelDir)
}

const outputDir = path.join(baseDir, "_output")
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir)
}

const backupDirName = "__backup"
const campusDirName = "_campus"

const dveDateRowName = "面試批次"

// const badDirName = '_missing data'

const ycInputDirConfig = {
  yc: {
    name: "from YC",
  },
}

const inputDirConfig = {
  award: {
    name: "award",
  },
  wayout: {
    name: "wayout",
  },
  master: {
    name: "master",
  },
  dve: {
    name: "dve interview",
    /**
     * only for folder generation
     */
    content: {
      yc: {
        name: "YC",
      },
      cci: {
        name: "HTICCI",
      },
    },
  },
}
const reverseInputDirConfig = Object.keys(inputDirConfig).reduce((acc, key) => {
  const config = inputDirConfig[key]
  acc[key] = config.name
  return acc
}, {})

const generatePath = (configs, res = {}, basepath = excelDir) =>
  Object.keys(configs).reduce((acc, key) => {
    const config = configs[key]

    acc[key] = path.join(basepath, config.name)

    if (config.content) return generatePath(config.content, res, acc[key])

    return acc
  }, res)

const ycInputDir = generatePath(ycInputDirConfig)
const inputDir = generatePath(inputDirConfig)

// Object.keys(inputDirConfig).reduce((acc, key) => {
//   const config = inputDirConfig[key]
//   acc[key] = path.join(excelDir, config.name)
//   return acc
// }, {})

//========//grantham config

const __config_file_name = "config.json"

const __grantham_start_year = "Grantham Start Date"
const __grantham_start_month = "Grantham Start Month"
const __config_copyToBackup = "Copy Input to Output"
const __config_removeInput = "Remove Input On Each Run"
const __config_individualSummery = "Make Console Summery For Each Excel"

var _configFile, _config

try {
  var _configFile = fs.readFileSync(path.join(baseDir, __config_file_name))
  _config = JSON.parse(_configFile)

  console.log("Config Found, loaded with these values:")
  console.table(_config)
  console.log()
} catch (e) {
  console.log("Config Not Found, creating a new one with default values\n")

  _config = {
    [__grantham_start_year]: 2021,
    [__grantham_start_month]: 6,
    [__config_copyToBackup]: true,
    [__config_removeInput]: false,
    [__config_individualSummery]: false,
  }

  fs.writeFileSync(
    path.join(baseDir, __config_file_name),
    JSON.stringify(_config, null, "\t")
  )
}

export const grantham_start_year = _config[__grantham_start_year]
export const grantham_start_month = _config[__grantham_start_month]

const copyToBackup = _config[__config_copyToBackup]
const removeInput = _config[__config_removeInput]
const individualSummery = _config[__config_individualSummery]

//========//solvers
let _base_headers = header
let _base_solver = solver

///=================================================================================================
//
// utils
//
///=================================================================================================

function syncProp(obj1, obj2, prop) {
  if (obj1?.[prop] && !obj2?.[prop]) {
    obj2[prop] = obj1[prop]
  } else if (!obj1?.[prop] && obj2?.[prop]) {
    obj1[prop] = obj2[prop]
  }
}

function addToCollection(collection, record) {
  const id = record.id
  const hkid = record.hkId

  if (!collection.__hkId) collection.__hkId = new Map()
  if (!collection.__hkId_found) collection.__hkId_found = new Set()

  let arr

  if (id && collection.has(id)) {
    arr = collection.get(id)
  } else if (hkid && collection.__hkId.has(hkid)) {
    arr = collection.__hkId.get(hkid)
  } else {
    arr = []
  }

  arr.push(record)

  let firstRecord = arr[0]

  syncProp(firstRecord, record, "id")

  if (id) collection.set(id, arr)
  if (hkid) collection.__hkId.set(hkid, arr)
  if (id && hkid) collection.__hkId_found.add(hkid)

  if (!id && !hkid) {
    if (!collection.has("no id")) collection.set("no id", [])
    let arr = collection.get("no id")
    arr.push(record)
  }
}

function extractYear(fileName) {
  const currentYear = new Date().getFullYear()
  const yearPrefix = Math.floor(currentYear / 100)

  let match,
    year = 0
  if ((match = fileName.match(/AY\d{4}/))) {
    year = yearPrefix + match[0].substring(4)
  } else if ((match = fileName.match(/\d{2,4}/))) {
    year = match[0]
    if (year.length === 2) year = yearPrefix + year
  }
  return year
}

//this is so funny
const remakeHeaders = (headers) => {
  let email
  let hasDveEntry
  let teensClassRowId, confirmedTeensHeader
  const newHeaders = headers.map((rawheader, i) => {
    const header = rawheader.trim()

    // if (/dve entry/i.test(header)) {
    //   if (hasDveEntry) return headerRemakeNames.dveEntry
    //   hasDveEntry = true
    // }
    if (/course/i.test(header) && teensClassRowId != undefined) {
      confirmedTeensHeader = true
    } else if (/^programme$/i.test(header)) {
      if (teensClassRowId == undefined) teensClassRowId = i
      else return headerRemakeNames.diplomaName
    }

    if (email) {
      const temp = email
      email = null
      return temp
    }
    if (remarkRegex.test(header)) {
      return "__remark"
    } else if (/trade/i.test(header)) {
      email = "Email (Trade)"
      return "Class Tutor (Trade, Fullname)"
    } else if (/generic|genric/i.test(header)) {
      email = "Email (Generic)"
      return "Class Tutor (Generic, Fullname)"
    }
    return header
  })

  if (teensClassRowId != undefined && confirmedTeensHeader) {
    newHeaders[teensClassRowId] = headerRemakeNames.tsEntry
  }

  return newHeaders
}

const findDveYear = (rows) => {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]
    if (row.some((cell) => `${cell}`.includes(dveDateRowName))) {
      const interviewDateIndex = row.findIndex((cell) =>
        `${cell}`.includes(dveDateRowName)
      )
      for (let j = interviewDateIndex + 1; j < row.length; j++) {
        if (row[j] && `${row[j]}`.trim() !== "") {
          let split = row[j].split("-")
          if (split.length === 3) {
            console.log(
              `Found DVE "${dveDateRowName}" in row: ${i + 1}, ${split[1]}`
            )
            return split[1]
          }
        }
      }
    }
  }
  console.log(`DVE "${dveDateRowName}" not found`)
  return null
}

const findHeader = (rows) => {
  const headers = _base_headers.reduce((acc, curr) => {
    if (!curr) return acc
    acc[
      curr
        .trim()
        .replace(" ", "")
        .replace("（", "(")
        .replace("）", ")")
        .toLowerCase()
    ] = curr
    return acc
  }, {})

  //have to convert the matched headers to the same casing lol
  const matchedHeaders = {}

  for (let i = 0; i < rows.length; i++) {
    const row = remakeHeaders(rows[i])
    const columns = new Set()

    const matchingHeaders = row.filter((header) => {
      if (!header) return false

      const raw = header
        .trim()
        .toLowerCase()
        .replace(" ", "")
        .replace("（", "(")
        .replace("）", ")")

      // if (columns.has(raw))
      //   throw new Error(
      //     `There are multiple columns with the same name: [${header}], please refer to the document for the acceptable column headers`
      //   )
      // else columns.add(raw)

      const match = Boolean(headers[raw])
      if (match) {
        matchedHeaders[header] = headers[raw]
      }
      return match
    })

    if (matchingHeaders.length >= 4) {
      console.log(`Confirmed header: ${row.slice(0, 7).join(", ")}, ...`)
      return [row.map((col) => matchedHeaders[col] || col), i]
    }
  }
  throw new Error(
    "No valid header found, please refer to the document for the acceptable column headers"
  )
}

///=================================================================================================
//
// main
//
///=================================================================================================

const readExcel = async (file, type, name, parentCollectionManager, id) => {
  console.log(`Reading file: ${name}, year: ${extractYear(name)}`)
  const fileBuffer = fs.readFileSync(file)
  const workbook = xlsx.read(fileBuffer)

  if (workbook.SheetNames.length > 0) {
    console.log(`Worksheets:`, workbook.SheetNames)
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]

    const useAward = type === reverseInputDirConfig.award
    const useDVE = type === reverseInputDirConfig.dve

    const rawRows = xlsx.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
    })

    const year = useDVE ? findDveYear(rawRows) : extractYear(name)

    const headerRow = findHeader(rawRows)

    if (!headerRow) {
      return
    }

    const [header, headerIndex] = headerRow

    const rows = xlsx.utils.sheet_to_json(worksheet, {
      header: header,
      range: headerIndex + 1,
      raw: true,
    })

    const collectionManager = new CollectionManager()

    rows.forEach((row, rowNumber) => {
      // if (rowNumber >= 20) {
      //   return
      //   console.log(`Row ${rowNumber + 1}:`)
      //   console.table(Object.entries(row))
      // }

      if (!year) {
        console.log(`Year not found in file: ${name}`)
        return
      }

      const [record, error, level, warning, warningLevel] = _base_solver(
        rowNumber + 1, // 1-based index
        row,
        year,
        useDVE,
        useAward,
        name,
        type,
        id
      )

      if (record)
        if (!error) {
          parentCollectionManager.addSuccess(record)
          collectionManager.addSuccess(record)
        } else {
          console.log("failed", error)
          parentCollectionManager.addFailed(record)
          collectionManager.addFailed(record)
        }
      if (error) {
        parentCollectionManager.addError(error, level)
        collectionManager.addError(error, level)
      }
      if (warning) {
        parentCollectionManager.addWarning(warning, warningLevel)
        collectionManager.addWarning(warning, warningLevel)
      }
    })

    if (individualSummery) collectionManager.logSummary()
  } else {
    console.log(`No worksheets found in file: ${name}`)
  }
}

function generateReadableId() {
  const now = new Date()
  const year = now.getFullYear()
  const month = String(now.getMonth() + 1).padStart(2, "0")
  const day = String(now.getDate()).padStart(2, "0")
  const hours = String(now.getHours()).padStart(2, "0")
  const minutes = String(now.getMinutes()).padStart(2, "0")
  const seconds = String(now.getSeconds()).padStart(2, "0")

  return `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`
}

function _readAllFiles(dir, filePaths = [], fileNames = []) {
  const files = fs.readdirSync(dir)

  files
    .filter((file) => !file.startsWith("~$"))
    .forEach((file) => {
      const filePath = path.join(dir, file)
      const stat = fs.statSync(filePath)

      if (stat.isDirectory()) {
        _readAllFiles(filePath, filePaths, fileNames)
      } else {
        filePaths.push({ path: filePath, name: file })
        fileNames.push(file)
      }
    })

  return [filePaths, fileNames]
}

async function readAllFiles(id, collectionManager, inputDir) {
  for (const key of Object.keys(inputDir)) {
    const loc = inputDir[key]
    const name = path.basename(loc)

    if (loc && !fs.existsSync(loc)) {
      fs.mkdirSync(loc)
      continue
    }

    let [files, names] = _readAllFiles(loc)

    console.log(`Reading from folder: ${name}`)

    if (names.length) console.log(names)
    else console.log(orange("No file found"))

    console.log()

    for (const file of files) {
      await readExcel(file.path, name, file.name, collectionManager, id)
      console.log()
    }
  }
}

function mergeRecord(success, collection) {
  for (const [id, records] of success) {
    if (!id) continue
    for (const record of records) {
      let oldRecord = collection.get(id)
      if (!oldRecord) {
        collection.set(id, {
          ...record,
          __file: { ...record.__file },
        })
        continue
      }
      let newer = oldRecord.__year <= record.__year
      let oldType = oldRecord.__type.split(",")
      oldType = oldType[oldType.length - 1].trim()

      if (newer && oldType == record.__type) {
        writeIfExist(oldRecord, record)
        continue
      } else if (
        newer &&
        (record.type == reverseInputDirConfig.master ||
          record.type == reverseInputDirConfig.dve)
      ) {
        writeIfExist(oldRecord, record)
        continue
      }
      writeIfEmpty(oldRecord, record)
    }
  }
}

const debugYcSuccessGenerator = useExcelGenerator(exportSchema.scMasterSchema)
const debugSuccessGenerator = useExcelGenerator(exportSchema.debugSuccessSchema)
const debugNoDupeSchema = useExcelGenerator(exportSchema.debugNoDupeSchema)
const debugFailGenerator = useExcelGenerator(exportSchema.debugFailSchema)
const campusGenerator = useExcelGenerator(exportSchema.contactsc, "body")

function logMapKeys(map) {
  let keys = Array.from(map.keys())
  // map.forEach((value, key) => {
  // });
  console.log(keys)
  // console.log(keys.slice(-10))
}

function concatWorkbook(workbook, workbook2) {
  workbook2.eachSheet((worksheet, sheetId) => {
    const newWorksheet = workbook.addWorksheet(worksheet.name)

    worksheet.columns.forEach((col, index) => {
      const newCol = newWorksheet.getColumn(index + 1)
      newCol.width = col.width
    })

    worksheet.eachRow((row, rowNumber) => {
      const newRow = newWorksheet.getRow(rowNumber)
      newRow.alignment = row.alignment

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber)
        newCell.value = cell.value
        newCell.style = cell.style
        newCell.border = cell.border
        newCell.fill = cell.fill
        newCell.font = cell.font
        newCell.numFmt = cell.numFmt
        newCell.protection = cell.protection
      })

      newRow.height = row.height
    })
    if (worksheet.views && worksheet.views.length > 0) {
      newWorksheet.views = worksheet.views.map((view) => ({ ...view }))
    }
  })
}

const generateMaster = async (fromYc) => {
  const id = generateReadableId()
  const collectionManager = new CollectionManager()

  console.log("Reading files from input folders... \n")

  await readAllFiles(id, collectionManager, fromYc ? ycInputDir : inputDir)

  let workbookSuccess = null,
    workbookFailed = null,
    workbookUnique = null,
    workbookAward = null

  let noDupeCollection = collectionManager.noDupeCollection

  if (collectionManager.successCollection.size > 0) {
    const generator = fromYc ? debugYcSuccessGenerator : debugSuccessGenerator
    workbookSuccess = generator({
      [fromYc ? "總表" : "所有記錄"]: collectionManager.successCollection,
    })
    mergeRecord(collectionManager.successCollection, noDupeCollection)
    noDupeCollection.forEach(
      (rec) => (rec.__original = JSON.stringify(rec.__original, null, "    "))
    )

    // console.log('successCollection')
    // logMapKeys(collectionManager.successCollection)
    // console.log('noDupeCollection')
    // logMapKeys(noDupeCollection)
  } else {
    console.log(red("No successful records"))
    console.log()
  }

  if (collectionManager.successCollection.__hkId?.size > 0) {
    for (const [key, value] of collectionManager.successCollection.__hkId) {
      // console.log(key)
      if (!collectionManager.successCollection.__hkId_found.has(key))
        collectionManager.failedCollection.set(key, value)
    }
  }

  if (collectionManager.failedCollection.size > 0) {
    workbookFailed = debugFailGenerator({
      main: collectionManager.failedCollection,
    })
  } else {
    console.log(green("No failed records"))
    console.log()
  }

  let campusCollection = new Map()
  let deregCollection = new Map()
  let awardCollection = new Map()
  let nomineeCollection = new Map()
  // let masterDeregCollection = new Map()

  let campusFailedCollection = new Map()
  let campusSuccessCollection = new Map()

  let noCampusCollection = new Map()
  let noTimeCollection = new Map()

  if (noDupeCollection.size > 0 && !fromYc) {
    for (const [id, record] of noDupeCollection) {
      checker(record)
      if (record.awardYear) {
        awardCollection.set(id, record)
      } else {
        nomineeCollection.set(id, record)
        //if (record.__deregByMaster) masterDeregCollection.set(id, record)
      }
    }

    for (const [id, record] of nomineeCollection) {
      const currentCampus = record.campus
      const noDereg = !record.dereg
      const notAwarded = !record.awardYear
      const outOfRange = !record.__inRange

      if (currentCampus && noDereg && notAwarded && !outOfRange) {
        let campus = campusCollection.get(currentCampus)
        if (!campus) {
          campus = new Map()
          campusCollection.set(currentCampus, campus)
        }
        record.trade = null
        record.generic = null
        campus.set(id, record)
        campusSuccessCollection.set(id, record)
      } else if (notAwarded) {
        campusFailedCollection.set(id, record)
        if (!currentCampus)
          noCampusCollection.set(id, { ...record, ___forCampus: true })
        if (outOfRange) noTimeCollection.set(id, record)
        if (!noDereg) deregCollection.set(id, record)
      }
    }

    const granthamYear = parseInt(grantham_start_year) + 4
    const AY =
      `${granthamYear - 1}`.substring(2) + `${granthamYear}`.substring(2)

    workbookUnique = debugNoDupeSchema({
      ["所有未獲獎學生"]: nomineeCollection,
      //['新添加的 (總表 其他評語) dereg']: masterDeregCollection,
      ["不在YC Excel中"]: campusFailedCollection,
      ["Dereg"]: deregCollection,
      ["沒有Campus"]: noCampusCollection,
      [`超出Grantham選擇範圍(${grantham_start_year}/${grantham_start_month})`]:
        noTimeCollection,
      ["在YC Excel中"]: campusSuccessCollection,
      [`獲獎學生(AY${AY})`]: awardCollection,
    })
  }

  if (workbookSuccess) {
    if (workbookUnique) concatWorkbook(workbookUnique, workbookSuccess)

    collectionManager.logSummary()

    console.log("Generating excel...\n")

    const promises = [
      writeOutput(
        workbookUnique ?? workbookSuccess,
        "students",
        undefined,
        id,
        true
        // noDupeCollection.size
      ),
      // writeOutput(
      //   workbookAward,
      //   "award",
      //   awardDirName,
      //   id,
      //   true
      //   // awardCollection.size
      // ),
      writeOutput(
        workbookFailed,
        "_failed",
        undefined,
        id,
        true
        // countTotalRecords(collectionManager.failedCollection)
      ),
    ]

    if (campusCollection && campusCollection.size > 0)
      for (const [campus, collection] of campusCollection) {
        const workbookCampus = campusGenerator({
          [campus]: collection,
        })
        promises.push(
          writeOutput(
            workbookCampus,
            campus.replace("\\", "-").replace("/", "-"),
            campusDirName,
            id,
            false
            // collection.size
          )
        )
      }

    // promises.concat([
    //   // writeOutput(
    //   //   workbookSuccess,
    //   //   'records',
    //   //   undefined,
    //   //   id,
    //   //   true
    //   //   // countTotalRecords(collectionManager.successCollection)
    //   // ),
    // ])

    await Promise.all(promises)

    if (workbookSuccess || workbookUnique || workbookAward || workbookFailed)
      if (copyToBackup) moveFilesToBackup(excelDir, id)
  }
}

function copyDirectory(src, dest) {
  if (!fs.existsSync(dest)) {
    fs.mkdirSync(dest, { recursive: true })
  }

  const entries = fs.readdirSync(src, { withFileTypes: true })

  for (let entry of entries) {
    const srcPath = path.join(src, entry.name)
    const destPath = path.join(dest, entry.name)

    if (entry.isDirectory()) {
      copyDirectory(srcPath, destPath)
    } else {
      fs.copyFileSync(srcPath, destPath)
    }
  }
}

function deleteDirectory(src) {
  const entries = fs.readdirSync(src, { withFileTypes: true })

  for (let entry of entries) {
    const srcPath = path.join(src, entry.name)

    if (entry.isDirectory()) {
      deleteDirectory(srcPath)
    } else {
      fs.unlinkSync(srcPath)
    }
  }

  fs.rmdirSync(src)
}

function moveFilesToBackup(excelDir, id) {
  const backupDir = path.join(outputDir, id)
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir)
  }
  const backupSubDir = path.join(backupDir, backupDirName)
  if (!fs.existsSync(backupSubDir)) {
    fs.mkdirSync(backupSubDir)
  }

  const entries = fs.readdirSync(excelDir, { withFileTypes: true })

  for (let entry of entries) {
    const srcPath = path.join(excelDir, entry.name)
    const destPath = path.join(backupSubDir, entry.name)

    if (entry.isDirectory()) {
      const dirEntries = fs.readdirSync(srcPath)
      if (dirEntries.length > 0) {
        copyDirectory(srcPath, destPath)
        if (removeInput) {
          deleteDirectory(srcPath)
        }
      }
    } else {
      fs.copyFileSync(srcPath, destPath)
      if (removeInput) {
        fs.unlinkSync(srcPath)
      }
    }
  }

  console.log(
    ">",
    orange(`Moved inputs to backup folder:`),
    lightBlue(backupSubDir)
  )
}

async function writeOutput(data, name, dir, id, useId = true, size) {
  if (!data) return

  let outputSubDir = path.join(outputDir, id)
  if (!fs.existsSync(outputSubDir)) {
    fs.mkdirSync(outputSubDir)
  }
  if (dir) {
    outputSubDir = path.join(outputSubDir, dir)
    if (!fs.existsSync(outputSubDir)) {
      fs.mkdirSync(outputSubDir)
    }
  }

  let current_year = new Date().getFullYear()

  if (data) {
    const fileName = `${name}_${useId ? id : current_year}.xlsx`
    const filePath = path.join(outputSubDir, fileName)
    await data.xlsx.writeFile(filePath)
    console.log(`> ${green(name)} written to: ${lightBlue(filePath)}`)
  }
}

///=================================================================================================
//
// Manager
//
///=================================================================================================

function countTotalRecords(collection) {
  let totalRecords = 0
  for (const value of collection.values()) {
    if (Array.isArray(value)) {
      totalRecords += value.length // Count all records in the array
    } else {
      totalRecords += 1 // Count the single record
    }
  }
  return totalRecords
}

function isValueEmpty(value) {
  return value === undefined || value === null || value === ""
}

function writeRecord(old, record, fn) {
  if (!old) {
    return record
  }
  var oldType = old.__type
  var oldFile = old.__file
  var newFile = record.__file

  let oldOriginal

  if (Array.isArray(old.__original) && Array.isArray(record.__original)) {
    oldOriginal = [...old.__original, ...record.__original]
  }

  let oldYear = old.year,
    oldMonth = old.month

  fn(old, record)

  if (oldOriginal) old.__original = oldOriginal

  let isDve = record.__type.includes(reverseInputDirConfig.dve)
  let isWayout = record.__type.includes(reverseInputDirConfig.wayout)
  let hasDve = old.__type.includes(reverseInputDirConfig.dve)
  let hasWayout = old.__type.includes(reverseInputDirConfig.wayout)

  if ((hasDve && !isDve) || (hasWayout && !isWayout && !isDve)) {
    old.year = oldYear
    old.month = oldMonth
  } else {
    old.year = record.year
    old.month = record.month
  }

  Object.keys(newFile).forEach((key) => {
    if (!oldFile[key]) oldFile[key] = newFile[key]
    else oldFile[key].push(newFile[key])
  })

  old.__file = oldFile
  old.__type = oldType + `, ${record.__type}`

  old.__overwritten = true

  return old
}

function appendWithFile(files, oldAttr, attribute) {
  if (!attribute) return oldAttr || ""

  let filename = Object.keys(files).reduce((acc, key) => {
    var name = key
    var indexes = files[key]
    acc += `${name}(${indexes.join(", ")}), `
    return acc
  }, "")

  let newAttr = `[${filename.slice(0, -2)}]: ${attribute}`

  return oldAttr ? oldAttr + "; " + newAttr : newAttr
}

function writeIfExist(old, newRecord) {
  let oldClass = old.__programmeClass
  let oldRemark = old.remark
  let oldMasterRemark = old.__remark

  return writeRecord(old, newRecord, (old, record) => {
    Object.keys(record).forEach((key) => {
      if (!isValueEmpty(record[key])) {
        old[key] = record[key]
      }
    })

    old.remark = appendWithFile(record.__file, oldRemark, record.remark)
    old.__programmeClass = appendWithFile(
      record.__file,
      oldClass,
      record.__programmeClass
    )
    old.__remark = appendWithFile(
      record.__file,
      oldMasterRemark,
      record.__remark
    )
  })
}

function writeIfEmpty(old, newRecord) {
  return writeRecord(old, newRecord, (old, record) => {
    Object.keys(record).forEach((key) => {
      if (isValueEmpty(old[key])) {
        old[key] = record[key]
      }
    })
  })
}

class CollectionManager {
  constructor() {
    this.successCollection = new Map()
    this.failedCollection = new Map()
    this.noDupeCollection = new Map()
    this.errorCollection = {}
    this.warningCollection = {}
  }

  addToCollection(collection, record) {
    addToCollection(collection, record)
  }

  addSuccess(record) {
    this.addToCollection(this.successCollection, record)
    // this.addUnique(record)
  }

  addFailed(record) {
    this.addToCollection(this.failedCollection, record)
  }

  addError(error, level) {
    if (!this.errorCollection[level]) this.errorCollection[level] = {}
    if (!this.errorCollection[level][error])
      this.errorCollection[level][error] = 0
    this.errorCollection[level][error]++
  }

  addWarning(warning, level) {
    if (!this.warningCollection[level]) this.warningCollection[level] = {}
    if (!this.warningCollection[level][warning])
      this.warningCollection[level][warning] = 0
    this.warningCollection[level][warning]++
  }

  logWarning() {
    if (Object.keys(this.warningCollection).length > 0) {
      const warningArray = []
      for (const level in this.warningCollection) {
        for (const warning in this.warningCollection[level]) {
          warningArray.push({
            Level: level,
            Warning: warning,
            Count: this.warningCollection[level][warning],
          })
        }
      }
      console.log("Warnings:")
      console.table(warningArray)
      console.log()
    }
  }

  logFailedReasons() {
    if (Object.keys(this.errorCollection).length > 0) {
      const errorArray = []
      for (const level in this.errorCollection) {
        for (const error in this.errorCollection[level]) {
          errorArray.push({
            Level: level,
            Error: error,
            Count: this.errorCollection[level][error],
          })
        }
      }
      console.log("Failed reasons:")
      console.table(errorArray)
      console.log()
    }
  }

  logSummary() {
    const totalSuccessRecords = countTotalRecords(this.successCollection)
    const totalFailedRecords = countTotalRecords(this.failedCollection)

    const summary = {
      Records: totalSuccessRecords,
      Failed: totalFailedRecords,
      Students: this.noDupeCollection.size,
    }

    console.log("Summery:")
    console.table(summary)
    console.log()

    this.logWarning()
    this.logFailedReasons()
  }
}

const waitForUserExit = () => {
  console.log("Press any key to exit...")
  process.stdin.setRawMode(true)
  process.stdin.resume()
  process.stdin.on("data", process.exit.bind(process, 0))
}

const askForOption = () => {
  console.log("1) 為 YC 製作 Excel")
  console.log("2) 結合來自 YC 的 Excel")
  console.log("3) 為 DVE 教師製作 Excel")
  console.log()

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })

  return new Promise((resolve, reject) =>
    rl.question(`請選擇 1-3, 或輸入 0 退出: `, async (option) => {
      rl.close()
      console.log()

      if (option == "0") {
      } else if (option == "1") {
        await generateMaster()
      } else if (option == "2") {
        await generateMaster(true)
      } else {
        await askForOption()
      }

      resolve()
    })
  )
}

const main = async () => {
  try {
    await askForOption()
    // await generateMaster()
  } catch (error) {
    console.error("An error occurred:", error)
  } finally {
    if (isPkg) waitForUserExit()
  }
}

main()
