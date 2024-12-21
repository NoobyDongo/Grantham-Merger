import * as xlsx from 'xlsx'
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import { checker, exportSchema, header, remarkRegex, solver } from './schema.js'
import { useExcelGenerator } from './assembler.js'

///=================================================================================================
//
// params
//
///=================================================================================================

//========//grantham config
const grantham_start_year = 2020
const grantham_start_month = 6

const copyToBackup = true
const removeInput = false
const individualSummery = false

//========//exe config
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const isPkg = typeof process.pkg !== 'undefined'
export const baseDir = isPkg ? path.dirname(process.execPath) : __dirname

//========//excel config
const excelDir = path.join(baseDir, '_put your excels here')
if (!fs.existsSync(excelDir)) {
  fs.mkdirSync(excelDir)
}

const outputDir = path.join(baseDir, '_output')
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir)
}

const backupDirName = '__backup'
const awardDirName = '_award'
const campusDirName = '_campus'

const dveDateRowName = '面試批次'

// const badDirName = '_missing data'

const inputDirConfig = {
  award: {
    name: 'award'
  },
  wayout: {
    name: 'wayout'
  },
  master: {
    name: 'master'
  },
  dve: {
    name: 'dve interview'
  }
}
const reverseInputDirConfig = Object.keys(inputDirConfig).reduce((acc, key) => {
  const config = inputDirConfig[key]
  acc[key] = config.name
  return acc
}, {})

const inputDir = Object.keys(inputDirConfig).reduce((acc, key) => {
  const config = inputDirConfig[key]
  acc[key] = path.join(excelDir, config.name)
  return acc
}, {})

//========//solvers
let _base_headers = header
let _base_solver = solver

///=================================================================================================
//
// utils
//
///=================================================================================================

export const customColor = colorCode => text => {
  return `\x1b[38;5;${colorCode}m${text}\x1b[0m`
}

const green = customColor(2)
const red = customColor(1)
const orange = customColor(208)
const lightBlue = customColor(12)

function addToCollection (collection, record) {
  const id = record.id
  const hkid = record.hkId

  if (!collection.__hkId) collection.__hkId = new Map()

  let arr

  if (id && collection.has(id)) {
    arr = collection.get(id)
  } else if (hkid && collection.__hkId.has(hkid)) {
    arr = collection.__hkId.get(hkid)
  } else {
    arr = []
  }

  arr.push(record)

  if (id) collection.set(id, arr)
  if (hkid) collection.__hkId.set(hkid, arr)
  if (!id && !hkid) collection.set('no id', arr)
}

function extractYear (fileName) {
  const match = fileName.match(/\d{2,4}/)
  if (match) {
    let year = match[0]
    if (year.length === 2) {
      year = '20' + year // Assuming the years are in the 2000s
    }
    return year
  }
  return null
}

const remakeHeaders = headers => {
  let email
  return headers.map(header => {
    if (email) {
      const temp = email
      email = null
      return temp
    }
    if (remarkRegex.test(header)) {
      return '__remark'
    } else if (/trade/i.test(header)) {
      email = 'Email (Trade)'
      return 'Class Tutor (Trade, Fullname)'
    } else if (/generic|genric/i.test(header)) {
      email = 'Email (Generic)'
      return 'Class Tutor (Generic, Fullname)'
    }
    return header
  })
}

const findDveYear = worksheet => {
  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, raw: true })
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]
    if (row.some(cell => `${cell}`.includes(dveDateRowName))) {
      const interviewDateIndex = row.findIndex(cell =>
        `${cell}`.includes(dveDateRowName)
      )
      for (let j = interviewDateIndex + 1; j < row.length; j++) {
        if (row[j] && `${row[j]}`.trim() !== '') {
          let split = row[j].split('-')
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

const findHeader = worksheet => {
  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, raw: true })
  for (let i = 0; i < rows.length; i++) {
    const row = remakeHeaders(rows[i])
    const matchingHeaders = row.filter(header => _base_headers.includes(header))
    if (matchingHeaders.length >= 4) {
      console.log(`Confirmed header: ${row.slice(0, 8).join('-')}, ...`)
      return [row, i]
    }
  }
  console.log('No valid header found')
  return null
}

///=================================================================================================
//
// main
//
///=================================================================================================

const readExcel = async (file, type, name, parentCollectionManager, id) => {
  console.log(`Reading file: ${type} ${name}, year: ${extractYear(name)}`)
  const fileBuffer = fs.readFileSync(file)
  const workbook = xlsx.read(fileBuffer)

  if (workbook.SheetNames.length > 0) {
    console.log(`Worksheets:`, workbook.SheetNames)
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]

    const useAward = type === reverseInputDirConfig.award
    const useDVE = type === reverseInputDirConfig.dve
    const year = useDVE ? findDveYear(worksheet) : extractYear(name)

    const headerRow = findHeader(worksheet)

    if (!headerRow) {
      return
    }

    const [header, headerIndex] = headerRow

    const rows = xlsx.utils.sheet_to_json(worksheet, {
      header: header,
      range: headerIndex + 1,
      raw: true
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

function generateReadableId () {
  const now = new Date()
  const year = now.getFullYear()
  const month = String(now.getMonth() + 1).padStart(2, '0')
  const day = String(now.getDate()).padStart(2, '0')
  const hours = String(now.getHours()).padStart(2, '0')
  const minutes = String(now.getMinutes()).padStart(2, '0')
  const seconds = String(now.getSeconds()).padStart(2, '0')

  return `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`
}

function _readAllFiles (dir, filePaths = [], fileNames = []) {
  const files = fs.readdirSync(dir)

  files.forEach(file => {
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

async function readAllFiles (id, collectionManager) {
  for (const key of Object.keys(inputDir)) {
    const loc = inputDir[key]
    const name = path.basename(loc)

    if (loc && !fs.existsSync(loc)) {
      fs.mkdirSync(loc)
      continue
    }

    let [files, names] = _readAllFiles(loc)

    console.log(name)
    console.log(names)

    for (const file of files) {
      console.log()
      await readExcel(file.path, name, file.name, collectionManager, id)
      console.log()
    }
  }
}

function mergeRecord (success, collection) {
  for (const [id, records] of success) {
    if (!id) continue
    for (const record of records) {
      let oldRecord = collection.get(id)
      if (!oldRecord) {
        collection.set(id, {
          ...record,
          __file: { ...record.__file }
        })
        continue
      }
      let newer = oldRecord.__year <= record.__year
      let oldType = oldRecord.__type.split(',')
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

const debugSuccessGenerator = useExcelGenerator(exportSchema.debugSuccessSchema)
const debugNoDupeSchema = useExcelGenerator(exportSchema.debugNoDupeSchema)
const debugAwardSchema = useExcelGenerator(exportSchema.debugAwardSchema)
const debugFailGenerator = useExcelGenerator(exportSchema.debugFailSchema)
const campusGenerator = useExcelGenerator(exportSchema.contactsc, 'body')

function logMapKeys (map) {
  let keys = Array.from(map.keys())
  // map.forEach((value, key) => {
  // });
  console.log(keys)
  // console.log(keys.slice(-10))
}

const test = async () => {
  const id = generateReadableId()
  const collectionManager = new CollectionManager()

  await readAllFiles(id, collectionManager)

  let workbookSuccess = null,
    workbookFailed = null,
    workbookUnique = null,
    workbookAward = null

  let noDupeCollection = collectionManager.noDupeCollection

  if (collectionManager.successCollection.size > 0) {
    workbookSuccess = debugSuccessGenerator({
      main: collectionManager.successCollection
    })
    mergeRecord(collectionManager.successCollection, noDupeCollection)

    // console.log('successCollection')
    // logMapKeys(collectionManager.successCollection)
    // console.log('noDupeCollection')
    // logMapKeys(noDupeCollection)
  } else {
    console.log(red('No successful records'))
    console.log()
  }

  if (collectionManager.failedCollection.size > 0) {
    workbookFailed = debugFailGenerator({
      main: collectionManager.failedCollection
    })
  } else {
    console.log(green('No failed records'))
    console.log()
  }

  let campusCollection = new Map()
  let deregCollection = new Map()
  let awardCollection = new Map()
  let nomineeCollection = new Map()
  let masterDeregCollection = new Map()

  let campusFailedCollection = new Map()
  let campusSuccessCollection = new Map()

  let noCampusCollection = new Map()
  let noTimeCollection = new Map()

  if (noDupeCollection.size > 0) {
    for (const [id, record] of noDupeCollection) {
      checker(record)
      if (record.awardYear) {
        awardCollection.set(id, record)
      } else {
        nomineeCollection.set(id, record)
        if (record.__deregByMaster) masterDeregCollection.set(id, record)
      }
    }

    if (awardCollection.size > 0)
      workbookAward = debugAwardSchema({
        main: awardCollection
      })

    for (const [id, record] of nomineeCollection) {
      const currentCampus = record.campus?.id
      const noDereg = !record.dereg
      const notAwarded = !record.awardYear
      const passedYear =
        record.year > grantham_start_year ||
        (record.year == grantham_start_year &&
          record.month >= grantham_start_month)

      if (currentCampus && noDereg && notAwarded && passedYear) {
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
        if (!currentCampus) noCampusCollection.set(id, record)
        if (!passedYear) noTimeCollection.set(id, record)
        if (!noDereg) deregCollection.set(id, record)
      }
    }

    workbookUnique = debugNoDupeSchema({
      ['所有未獲獎學生(系統將會處理更改)']: nomineeCollection,
      ['新添加的 (總表 其他評語) dereg']: masterDeregCollection,
      ['不在YC Excel中']: campusFailedCollection,
      ['Dereg']: deregCollection,
      ['沒有Campus']: noCampusCollection,
      [`超出Grantham選擇範圍(${grantham_start_year}/${grantham_start_month})`]:
        noTimeCollection,
      ['在YC Excel中']: campusSuccessCollection
    })
  }

  if (campusCollection && campusCollection.size > 0)
    for (const [campus, collection] of campusCollection) {
      const workbookCampus = campusGenerator({
        [campus]: collection
      })
      await writeOutput(
        workbookCampus,
        campus.replace('\\', '-').replace('/', '-'),
        campusDirName,
        id,
        false
        // collection.size
      )
    }

  await writeOutput(
    workbookSuccess,
    'records',
    undefined,
    id,
    true
    // countTotalRecords(collectionManager.successCollection)
  )
  await writeOutput(
    workbookUnique,
    'students',
    undefined,
    id,
    true
    // noDupeCollection.size
  )
  await writeOutput(
    workbookAward,
    'award',
    awardDirName,
    id,
    true
    // awardCollection.size
  )
  await writeOutput(
    workbookFailed,
    '_failed',
    undefined,
    id,
    true
    // countTotalRecords(collectionManager.failedCollection)
  )

  if (workbookSuccess || workbookUnique || workbookAward || workbookFailed)
    if (copyToBackup) moveFilesToBackup(excelDir, id)

  collectionManager.logSummary()
}

function copyDirectory (src, dest) {
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

function deleteDirectory (src) {
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

function moveFilesToBackup (excelDir, id) {
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
    '>',
    orange(`Moved inputs to backup folder:`),
    lightBlue(backupSubDir)
  )
}

async function writeOutput (data, name, dir, id, useId = true, size) {
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

function countTotalRecords (collection) {
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

function isValueEmpty (value) {
  return value === undefined || value === null || value === ''
}

function writeRecord (old, record, fn) {
  if (!old) {
    return record
  }
  var oldType = old.__type
  var oldFile = old.__file
  var newFile = record.__file

  let oldYear = old.year,
    oldMonth = old.month

  fn(old, record)

  let isDve = record.__type.includes(reverseInputDirConfig.dve)
  let isWayout = record.__type.includes(reverseInputDirConfig.wayout)
  let hasDve = old.__type.includes(reverseInputDirConfig.dve)
  let hasWayout = old.__type.includes(reverseInputDirConfig.wayout)

  // console.log(
  //   'old',
  //   oldYear,
  //   old.year,
  //   record.year,
  //   {
  //     isDve,
  //     isWayout,
  //     hasDve,
  //     hasWayout,
  //     ifExist: !!ifExist
  //   },
  //   {
  //     type: old.__type
  //   },
  //   {
  //     type: record.__type
  //   }
  // )

  if ((hasDve && !isDve) || (hasWayout && !isWayout && !isDve)) {
    old.year = oldYear
    old.month = oldMonth
    // console.log('changed to', old.year, old.month)
  } else {
    old.year = record.year
    old.month = record.month
  }

  Object.keys(newFile).forEach(key => {
    if (!oldFile[key]) oldFile[key] = newFile[key]
    else oldFile[key].push(newFile[key])
  })

  old.__file = oldFile
  old.__type = oldType + `, ${record.__type}`

  old.__overwritten = true

  return old
}

function appendWithFile (files, oldAttr, attribute) {
  if (!attribute) return oldAttr || ''

  let filename = Object.keys(files).reduce((acc, key) => {
    var name = key
    var indexes = files[key]
    acc += `${name}(${indexes.join(', ')}), `
    return acc
  }, '')

  let newAttr = `[${filename.slice(0, -2)}]: ${attribute}`

  return oldAttr ? oldAttr + '; ' + newAttr : newAttr
}

function writeIfExist (old, newRecord) {
  let oldClass = old.__programmeClass
  let oldRemark = old.remark
  let oldMasterRemark = old.__remark

  return writeRecord(old, newRecord, (old, record) => {
    Object.keys(record).forEach(key => {
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

function writeIfEmpty (old, newRecord) {
  return writeRecord(old, newRecord, (old, record) => {
    Object.keys(record).forEach(key => {
      if (isValueEmpty(old[key])) {
        old[key] = record[key]
      }
    })
  })
}

class CollectionManager {
  constructor () {
    this.successCollection = new Map()
    this.failedCollection = new Map()
    this.noDupeCollection = new Map()
    this.errorCollection = {}
    this.warningCollection = {}
  }

  addToCollection (collection, record) {
    addToCollection(collection, record)
  }

  // addUnique (record) {
  //   let id = record.id
  //   if (!id) return
  //   let oldRecord = this.noDupeCollection.get(id)

  //   if (!oldRecord) {
  //     this.noDupeCollection.set(id, { ...record, __file: { ...record.__file } })
  //     return
  //   }
  //   let newer = oldRecord.__year <= record.__year
  //   let oldType = oldRecord.__type.split(',')
  //   oldType = oldType[oldType.length - 1].trim()

  //   if (newer && oldType == record.__type) {
  //     writeIfExist(oldRecord, record)
  //     return
  //   } else if (
  //     newer &&
  //     (record.type == reverseInputDirConfig.master ||
  //       record.type == reverseInputDirConfig.dve)
  //   ) {
  //     writeIfExist(oldRecord, record, true)
  //     return
  //   }
  //   writeIfEmpty(oldRecord, record)
  // }

  addSuccess (record) {
    this.addToCollection(this.successCollection, record)
    // this.addUnique(record)
  }

  addFailed (record) {
    this.addToCollection(this.failedCollection, record)
  }

  addError (error, level) {
    if (!this.errorCollection[level]) this.errorCollection[level] = {}
    if (!this.errorCollection[level][error])
      this.errorCollection[level][error] = 0
    this.errorCollection[level][error]++
  }

  addWarning (warning, level) {
    if (!this.warningCollection[level]) this.warningCollection[level] = {}
    if (!this.warningCollection[level][warning])
      this.warningCollection[level][warning] = 0
    this.warningCollection[level][warning]++
  }

  // logRandomSuccess () {
  //   if (this.successCollection.size > 0) {
  //     console.log('Random record for review:')
  //     let arr = Array.from(this.successCollection.entries())
  //     console.log(
  //       // arr[Math.floor(Math.random() * this.successCollection.size)][1],
  //       // arr[Math.floor(Math.random() * this.successCollection.size)][1],
  //       arr[Math.floor(Math.random() * this.successCollection.size)][1]
  //     )
  //   }
  // }

  logWarning () {
    if (Object.keys(this.warningCollection).length > 0) {
      const warningArray = []
      for (const level in this.warningCollection) {
        for (const warning in this.warningCollection[level]) {
          warningArray.push({
            Level: level,
            Warning: warning,
            Count: this.warningCollection[level][warning]
          })
        }
      }
      console.log('Warnings:')
      console.table(warningArray)
    }
  }

  logFailedReasons () {
    if (Object.keys(this.errorCollection).length > 0) {
      const errorArray = []
      for (const level in this.errorCollection) {
        for (const error in this.errorCollection[level]) {
          errorArray.push({
            Level: level,
            Error: error,
            Count: this.errorCollection[level][error]
          })
        }
      }
      console.log('Failed reasons:')
      console.table(errorArray)
    }
  }

  logSummary () {
    const totalSuccessRecords = countTotalRecords(this.successCollection)
    const totalFailedRecords = countTotalRecords(this.failedCollection)

    const summary = {
      Success: totalSuccessRecords,
      Failed: totalFailedRecords,
      Students: this.noDupeCollection.size
    }

    console.log()
    console.log('Summary:')
    console.table(summary)

    this.logWarning()
    this.logFailedReasons()
    //this.logRandomSuccess()
  }
}

const waitForUserInput = () => {
  console.log('Press any key to exit...')
  process.stdin.setRawMode(true)
  process.stdin.resume()
  process.stdin.on('data', process.exit.bind(process, 0))
}

const main = async () => {
  try {
    await test()
  } catch (error) {
    console.error('An error occurred:', error)
  } finally {
    if (isPkg) waitForUserInput()
  }
}

main()
