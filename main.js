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
import { generateTeacherWorkbook, useExcelGenerator } from "./assembler.js"
import { protectExcelFile } from "./xlsx-worker.js"

///=================================================================================================
//
// log utils
//
///=================================================================================================

export const customColor = (colorCode) => (text) => {
  return `\x1b[38;5;${colorCode}m${text}\x1b[0m`
}

const yellow = customColor(3)
const magenta = customColor(13)
const blue = customColor(26)
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

const outputsDir = path.join(baseDir, "_output")
if (!fs.existsSync(outputsDir)) {
  fs.mkdirSync(outputsDir)
}

const backupDirName = "__backup"
const campusDirName = "_campus"

const dveDateRowName = "面試批次"

// const badDirName = '_missing data'

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

const dveInputDir = generatePath({
  dve: {
    name: "for dve teachers",
  },
})
const ycInputDir = generatePath({
  yc: {
    name: "from YC",
  },
})
const inputDir = generatePath(inputDirConfig)

// Object.keys(inputDirConfig).reduce((acc, key) => {
//   const config = inputDirConfig[key]
//   acc[key] = path.join(excelDir, config.name)
//   return acc
// }, {})

//========//grantham config

const __config_file_name = "config.json"

const numberParser = (value) => (value ? parseInt(value) : 0)

const generateDate = (date = new Date()) => {
  const chiNumbers = ["日", "一", "二", "三", "四", "五", "六"]

  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, "0")
  const day = String(date.getDate()).padStart(2, "0")
  const weekday = chiNumbers[date.getDay()]

  return `${year}年${month}月${day}日（星期${weekday}）`
}

const configs = {
  startYear: {
    key: "Grantham Selection Range(Year)",
    default: new Date().getFullYear(),
    parser: numberParser,
  },
  startMonth: {
    key: "Grantham Selection Range(Month)",
    default: 6,
    parser: numberParser,
  },
  copyToBackup: {
    key: "Copy Input to Output",
    default: true,
  },
  removeInput: {
    key: "Remove Input On Each Run",
    default: false,
  },
  individualSummery: {
    key: "Make Console Summery For Each Excel",
    default: false,
  },
  teacherExcel: {
    key: "Params for generating Excels for Dve Teachers",
    content: {
      address: {
        key: "VDPO Address",
        default: "",
      },
      phone: {
        key: "VDPO Contact(phone)",
        default: "0000 0000",
      },
      fax: {
        key: "VDPO Contact(fax)",
        default: "0000 0000",
      },
      deadline: {
        key: "Deadline",
        default: generateDate(),
      },
      config: {
        key: "Config",
        content: {
          password: {
            key: "Excel Password",
            default: "12345678",
          },
          coveredByCampus: {
            key: "Store generated Excel by Campus",
            default: false,
          },
        },
      },
    },
  },
}

const reverseKeyedConfigs = ((_obj) => {
  const reverseKey = (obj) =>
    Object.keys(obj).reduce((acc, key) => {
      //so that content of the not reversed object doent get changed
      acc[obj[key].key] = { ...obj[key], key: key }

      if (acc[obj[key].key].content)
        acc[obj[key].key].content = reverseKey(acc[obj[key].key].content)

      return acc
    }, {})
  return reverseKey(_obj)
})(configs)

var _configFile,
  //for vs code suggestions lol
  _config = configs,
  usedDefault = false

const parseOptions = (configs, json) =>
  Object.keys(configs).reduce((acc, key) => {
    const option = configs[key]

    if (option.content) {
      acc[option.key] = parseOptions(
        option.content,
        typeof json[key] === "object" ? json[key] : {}
      )
    } else {
      if (json[key] == undefined) {
        acc[option.key] = option.default
        usedDefault = true
      } else
        acc[option.key] = option.parser ? option.parser(json[key]) : json[key]
    }

    return acc
  }, {})

const generateDefaultConfig = (configs) =>
  Object.keys(configs).reduce((acc, curr) => {
    const option = configs[curr]
    if (option.content) acc[curr] = generateDefaultConfig(option.content)
    else acc[curr] = option.default
    return acc
  }, {})

const logConfigs = (config, _config, key) => {
  const withContent = {}
  const flat = Object.keys(_config)
    .filter((k) => {
      if (!config[k]) return false
      if (typeof _config[k] === "object") {
        const oriKey = config[k].key
        // console.log("orikey", oriKey)
        withContent[key ? `${key}/${oriKey}` : `${oriKey}`] = {
          config: config[k].content ? config[k].content : config[k],
          _config: _config[k],
        }
        return false
      }
      return true
    })
    .reduce((acc, key) => {
      const oriKey = config[key].key
      acc[oriKey] = _config[key]
      return acc
    }, {})

  if (key) console.log(`[${green(key)}]:`)
  if (Object.keys(flat).length > 0) {
    console.table(flat)
    console.log()
  }
  for (const key in withContent) {
    const pair = withContent[key]
    logConfigs(pair.config, pair._config, key)
  }
}

try {
  var _configFile = fs.readFileSync(path.join(baseDir, __config_file_name))
  const json = JSON.parse(_configFile)

  _config = parseOptions(reverseKeyedConfigs, json)

  console.log(`${green("Config Found")}, loaded with these values:`)
} catch (e) {
  console.log(
    `${orange("Config Not Found")}, creating a new one with default values\n`
  )

  _config = generateDefaultConfig(configs)
  usedDefault = true
}

const syncKeyName = (config, _config) =>
  Object.keys(_config).reduce((acc, key) => {
    if (!config[key]) return acc

    const oriKey = config[key].key

    acc[oriKey] =
      typeof _config[key] === "object" && config[key].content
        ? syncKeyName(config[key].content, _config[key])
        : _config[key]

    return acc
  }, {})

if (usedDefault) {
  fs.writeFileSync(
    path.join(baseDir, __config_file_name),
    JSON.stringify(syncKeyName(configs, _config), null, "\t")
  )
}

logConfigs(configs, _config)

export const grantham_start_year = parseInt(_config.startYear)
export const grantham_start_month = parseInt(_config.startMonth)

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
  console.log(`Reading file: ${name}, year: ${yellow(extractYear(name))}`)
  const fileBuffer = fs.readFileSync(file)
  const workbook = xlsx.read(fileBuffer)

  if (workbook.SheetNames.length > 0) {
    console.log(
      `Worksheets:`,
      `[${workbook.SheetNames.map((name) => green(name)).join(", ")}]`
    )
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

    if (_config.individualSummery) collectionManager.logSummary()
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
    console.log(`Reading from folder: [${green(name)}]`)

    if (loc && !fs.existsSync(loc)) {
      fs.mkdirSync(loc)
      console.log(orange("Folder not found, created a new one\n"))
      continue
    }

    let [files, names] = _readAllFiles(loc)

    if (names.length)
      console.table(
        names.reduce((acc, name, i) => {
          acc[i + 1] = name
          return acc
        }, {})
      )
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
  collection.forEach(
    (rec) => (rec.__original = JSON.stringify(rec.__original, null, "    "))
  )
}

const debugYcSuccessGenerator = useExcelGenerator(exportSchema.scMasterSchema)
const debugSuccessGenerator = useExcelGenerator(
  exportSchema.debugSuccessSchema,
  undefined,
  3,
  true,
  15
)
const debugNoDupeSchema = useExcelGenerator(
  exportSchema.debugNoDupeSchema,
  undefined,
  3,
  true,
  15
)
const debugFailGenerator = useExcelGenerator(
  exportSchema.debugFailSchema,
  undefined,
  3,
  true,
  15
)
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

const generateTeacherExcel = async () => {
  const id = generateReadableId()
  const collectionManager = new CollectionManager()

  console.log("Reading files from input folders... \n")

  await readAllFiles(id, collectionManager, dveInputDir)

  let workbookFailed = null

  const teacherStudentCollection = new Map()
  const teacherCollection = new Map()

  let count = 0

  if (collectionManager.successCollection.size > 0) {
    mergeRecord(
      collectionManager.successCollection,
      collectionManager.noDupeCollection
    )

    console.log("Generating excel...\n")
    const passedSet = new Set()

    for (const [id, student] of collectionManager.noDupeCollection) {
      // if (student.programmeClass_name === "DEREG") console.log(student)
      if (student.dereg) continue
      if (student.generic || student.trade) {
        const setTeacher = (teacher, student) => {
          if (!teacher || !teacher.email) return

          if (
            !student.programmeClass_name ||
            !student.diploma ||
            !student.diploma.id
          ) {
            console.log(student)
            throw new Error("Student has teacher but no class or diploma")
          }

          if (!teacherCollection.has(teacher.email))
            teacherCollection.set(teacher.email, teacher)

          if (teacherStudentCollection.has(teacher.email))
            teacherStudentCollection.get(teacher.email).push(student)
          else teacherStudentCollection.set(teacher.email, [student])

          passedSet.add(id)
        }

        setTeacher(student.generic, student)
        setTeacher(student.trade, student)
      }
    }

    for (const [id, student] of collectionManager.noDupeCollection) {
      if (!passedSet.has(id))
        collectionManager.failedCollection.set(id, student)
    }

    collectionManager.logSummary()
    stateOperationResult(collectionManager)

    if (collectionManager.failedCollection.size > 0) {
      workbookFailed = debugNoDupeSchema({
        main: collectionManager.failedCollection,
      })
    }

    if (teacherCollection.size == 0)
      return console.log(red("No Teacher Found\n"))

    const promises = []

    if (workbookFailed)
      promises.push(
        writeOutput(workbookFailed, "_not_included", undefined, id, true)
      )

    console.log("It might take some time... \n")

    for (const [email, students] of teacherStudentCollection) {
      count++
      // if (count > 1) continue

      const campus = Array.from(new Set(students.map((s) => s.campus))).join(
        ", "
      )

      const teacher = teacherCollection.get(email)

      const goodBook = generateTeacherWorkbook(
        _config.teacherExcel,
        students.map((s) => ({
          0: s.cname || s.ename,
          1: s.diploma?.id,
          2: s.programmeClass_name,
          3: s.campus,
        }))
      )

      promises.push(
        writeOutput(
          goodBook,
          `電郵老師的excel 表格sample (${email})_${campus}_${teacher.name}`.replace(
            /[\\/\<\>\|*"\?]/g,
            "-"
          ),
          _config.teacherExcel.config.coveredByCampus
            ? campus.replace(/[\\/\<\>\|*"\?]/g, "-")
            : undefined,
          id,
          false,
          _config.teacherExcel.config.password
        )
      )
    }
    await Promise.all(promises)

    console.log()

    if (collectionManager.successCollection.size > 0 || workbookFailed)
      if (_config.copyToBackup)
        moveFilesToBackup(excelDir, id, new Set(Object.values(dveInputDir)))
  }
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

  stateOperationResult(collectionManager)

  if (collectionManager.successCollection.size > 0) {
    const generator = fromYc ? debugYcSuccessGenerator : debugSuccessGenerator
    workbookSuccess = generator({
      [fromYc ? "總表" : "所有記錄"]: collectionManager.successCollection,
    })
    mergeRecord(collectionManager.successCollection, noDupeCollection)

    // console.log('successCollection')
    // logMapKeys(collectionManager.successCollection)
    // console.log('noDupeCollection')
    // logMapKeys(noDupeCollection)
  }

  if (collectionManager.failedCollection.size > 0) {
    workbookFailed = debugFailGenerator({
      main: collectionManager.failedCollection,
    })
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
      if (_config.copyToBackup)
        moveFilesToBackup(
          excelDir,
          id,
          new Set(Object.values(fromYc ? ycInputDir : inputDir))
        )
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

function moveFilesToBackup(excelDir, id, allowedDirs) {
  const outputDir = path.join(outputsDir, id)
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir)
  }
  const backupDir = path.join(outputDir, backupDirName)
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir)
  }

  const entries = fs.readdirSync(excelDir, { withFileTypes: true })

  for (let entry of entries) {
    const srcPath = path.join(excelDir, entry.name)
    const destPath = path.join(backupDir, entry.name)

    if (entry.isDirectory()) {
      if (!allowedDirs?.has(srcPath)) continue

      const dirEntries = fs.readdirSync(srcPath)
      if (dirEntries.length > 0) {
        copyDirectory(srcPath, destPath)
        if (_config.removeInput) {
          deleteDirectory(srcPath)
        }
      }
    }
    // else I dont think I should care lol
    // else {
    //   fs.copyFileSync(srcPath, destPath)
    //   if (_config.removeInput) {
    //     fs.unlinkSync(srcPath)
    //   }
    // }
  }

  console.log(
    ">",
    orange(`Moved inputs to backup folder:`),
    lightBlue(backupDir),
    "\n"
  )
}

async function writeOutput(data, name, dir, id, useId = true, password = "") {
  if (!data) return

  let outputSubDir = path.join(outputsDir, id)
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

    if (password) {
      // let t0, t1

      // t0 = performance.now()
      await protectExcelFile(filePath, password)
      // t1 = performance.now()

      // console.log(`Writing took: ${t1 - t0} ms`)
    }

    console.log(`> ${green(name)} written to: ${lightBlue(filePath)}`)
  }
}

///=================================================================================================
//
// Manager
//
///=================================================================================================

function stateOperationResult(collectionManager) {
  if (collectionManager.successCollection.__hkId?.size > 0) {
    for (const [key, value] of collectionManager.successCollection.__hkId) {
      // console.log(key)
      if (!collectionManager.successCollection.__hkId_found.has(key))
        collectionManager.failedCollection.set(key, value)
    }
  }

  if (
    collectionManager.successCollection.size == 0 &&
    collectionManager.failedCollection.size == 0
  ) {
    console.log(red("No record\n"))
  } else {
    if (collectionManager.successCollection.size == 0)
      console.log(red("No successful record\n"))

    if (collectionManager.failedCollection.size == 0)
      console.log(green("No failed record\n"))
  }
}

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
    rl.question(
      `請選擇 ${yellow("1")}-${yellow("3")}, 或輸入 ${red("0")} 退出: `,
      async (option) => {
        rl.close()
        console.log()

        if (option == "0") {
        } else if (option == "1") {
          await generateMaster()
        } else if (option == "2") {
          await generateMaster(true)
        } else if (option == "3") {
          await generateTeacherExcel()
        } else {
          await askForOption()
        }

        resolve()
      }
    )
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
