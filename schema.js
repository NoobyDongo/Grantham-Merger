// #region Definitions

import { grantham_start_month, grantham_start_year } from "./main.js"

///=================================================================================================
//
// Definitions
//
///=================================================================================================

export const headerRemakeNames = {
  diplomaName: "Diploma Name",
  tsEntry: "Teens Entry",
  dveEntry: "Dve Entry 2",
}

const enrollment = "期數",
  enrollment_eng = "DVE Entry",
  ename = "英文姓名",
  cname = "中文姓名",
  wayout_cname = "學生姓名",
  hkId = "身份證號碼",
  className1 = "學期",
  className2 = "班別",
  vdpId = "VDP 編號",
  guide_name = "學生輔導主任",
  diploma_id = "課程編號",
  diploma_name = "課程名稱",
  programme_campus = "上課地點",
  wayout_programme = "詳情",
  master_remark = "評語",
  programmeClass_name2 = "(總表 DVE Class)",
  //header for class tutor info would already be parsed by remakeTutorHeaders in main.js
  programmeClass_generic = {
    name: "Class Tutor (Generic, Fullname)",
    email: "Email (Generic)",
  },
  programmeClass_trade = {
    name: "Class Tutor (Trade, Fullname)",
    email: "Email (Trade)",
  },
  awardYear = "(得獎年份)",
  dereg = "(Dereg/Grad)"

export const remarkRegex = new RegExp(master_remark, "g")

const weakClassname = "TC00N00AA0AA"

//==================================
// Regex
//==================================

const regex = {
  className1_1: new RegExp("^\\d{2}[A-Z]{1}\\d{2}$", "i"),
  className1_2: new RegExp("^[A-Z]{2,3}\\d{1}[A-Z]{2}$", "i"),

  className: new RegExp("^TC\\d{2,3}[A-Z]\\d{2}[A-Z]{2,3}\\d[A-Z]{2}$", "i"),

  //18N01 OATM? && 18N02 SFTO from master 2020
  classNameRare: new RegExp("^[A-Z]{4}$", "i"),

  //around 2015 from award
  className3_1: new RegExp("^\\d{2}[A-Z]{1}\\d{2}$", "i"),
  className3_2: new RegExp("^[A-Z]{2,3}\\d{1}[A-Z]{2,3}$", "i"),

  //around 2010 from award
  className2_1: new RegExp("^\\d{2,3}[A-Z]{2,3}$", "i"),
  className2_2: new RegExp("^[A-Z]{1,3}$", "i"),

  entry: new RegExp("^\\d{4}/\\d{2}$"),

  details: new RegExp(
    "[A-Z]{2}\\d{6}[A-Z]?\\s?-\\s?.+\\s?-\\s?(IVE\\([A-Z]{2,3}\\)|YC\\([A-Z]{2,3}\\)|CCI|HKDI|[A-Z]{2,10})", //i give up
    "i"
  ),

  // 818/16/pc, 404/22/PDFB
  class_complex: new RegExp(".{2,3}\\/.{2}\\/.{2,4}"),
  // AB123456, AB123456A, AB123456/12/AB
  class_withProgrammeCode: new RegExp(
    "^[A-Z]{2}\\d{6}[A-Z]?(-|\\s|/\\d{2}(/[A-Z]{2})?)?$",
    "i"
  ),
  // seeing new things everyday
  class_withProgrammeCodeAndClassWithNoSpace: new RegExp(
    "^[A-Z]{2}\\d{6}[A-Z]?.{2,5}$",
    "i"
  ),
  // AB123456/AB, AB123456/ABCD, AB123456/11AB, AB123456A/1D, AB123456A/A1D, AB123456/ddd2a
  class_programmeCodeWithClass: new RegExp(
    "^[A-Z]{2}\\d{6}[A-Z]?/.{2,5}$",
    "i"
  ),
  // A1A, V12c, a-1d, d1, A11A, 1A, 11A
  class_pure: new RegExp("^([A-Z]|[A-Z]-)?\\d{1,2}[A-Z][A-Z]?$", "i"),
  class_veryPure: new RegExp("^[A-Z]{1}$", "i"),

  programmeCode_pure: new RegExp("^[A-Z]{2}\\d{6}$", "i"),
  programmeCode_variant: new RegExp("^[A-Z]{2}\\d{6}[A-Z]{1,2}$", "i"),
  programmeName_variant: new RegExp("^.*[A-Z]{1,2}$", "i"),

  //                                  bro i swear i had seen everything
  no: /(dereg|grad|no|de-reg|defer|quit|withdraw|granduated|transferred out)/i,

  //so sad
  masterNo:
    /(dereg|de-reg|defer|quit|withdraw|employ|(?!未)退學|no record|no this student|no data|沒記錄|(?!未)畢業|no student|not found|unable to locate)/i,

  hkIdStrict: new RegExp("^[A-Z]{1,2}\\d{6}\\([0-9A]\\)$", "i"),
  hkId: new RegExp("^[A-Z]{1,2}\\d{6}\\(?[0-9A]\\)?$", "i"),

  empty: new RegExp("^n[/\\\\]?a$", "i"),
  space: new RegExp(
    "[\t\n\v\f\r \u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u2028\u2029\u3000]+"
  ),

  campus: new RegExp("^HTI|CCI|ICI|YC\\(.*\\)|IVE\\(.*\\)$", "i"),
}

//==================================
// Accessors + validators
//==================================

const xl = 20,
  lg = 15,
  md = 10,
  sm = 7.5,
  xs = 5

const getParsedValue = (obj, def, formats, checker) => {
  const names = def.name || def

  let pendingNames = Array.isArray(names) ? [...names] : [names],
    value,
    name

  while (!value && (name = pendingNames.shift())) {
    if (
      obj[name] !== null &&
      obj[name] !== undefined &&
      spaceChecker(checker ? checker(`${obj[name]}`) : obj[name])
    )
      value = `${obj[name]}`
  }

  if (
    value &&
    (formats
      ? Array.isArray(formats)
        ? formats.some((f) => f.test(value))
        : formats.test(value)
      : true)
  )
    return value
  else return null
}
const removeSpace = (str) => (str ? `${str}`.replace(regex.space, "") : null)
const formatSpace = (str) =>
  str ? `${str}`.replace(regex.space, " ").trim() : null

const spaceChecker = (value) => Boolean(removeSpace(value))

const standardHeader = {
  dveEntry: enrollment_eng,
  ename: "Eng Name",
  cname: "Chi Name",
  hkid: "ID",
  tsEntry: "Programme",
  tsClass: "Course",
  sc: "SC",
  campus: "Campus",
  diplomaId: "DVE Programme Offered",
  diplomaName: "Programme",
  dveClass: "Class",
  dveTTn: programmeClass_trade.name,
  dveTTe: programmeClass_trade.email,
  dveTGn: programmeClass_generic.name,
  dveTGe: programmeClass_generic.email,
  masterRemark: "其他評語",
  avgMark: "總分(100%)",
}

const _className = {
  classname1: {
    name: [className1, headerRemakeNames.tsEntry],
    set: (record) => record.className?.slice(2, 7),
    get: (record) =>
      removeSpace(
        getParsedValue(record, _className.classname1, [
          regex.className1_1,
          regex.className2_1,
          regex.className3_1,
        ])
      ),
  },
  classname2: {
    name: [className2, standardHeader.tsClass, "teens class"],
    set: (record) => record.className?.slice(7),
    get: (record) => {
      const c2 = removeSpace(
        getParsedValue(record, _className.classname2, [
          regex.className1_2,
          regex.className2_2,
          regex.className3_2,
        ])
      )

      if (c2 && regex.classNameRare.test(c2))
        return c2.slice(0, 2) + "1" + c2.slice(2)
      return c2
    },
  },
}

const _base = {
  entryDate: {
    name: [standardHeader.dveEntry, headerRemakeNames.dveEntry, enrollment, ""],
    get: (record) => {
      let entry = removeSpace(
        getParsedValue(record, _base.entryDate, regex.entry)
      )

      if (!entry) return null

      entry = `${entry}`.match(regex.entry)?.[0]?.split("/")

      return {
        year: entry[0],
        month: entry[1],
      }
    },
    set: (record) => {
      if (!(record.year && record.month)) {
        return null
      }
      let entry = record.year + "/" + record.month.toString().padStart(2, "0")

      // if (!record.__entry && !record.___forCampus) {
      //   return `[${entry}]`
      // }
      return entry
    },
  },
  ename: {
    name: [ename, "English name", standardHeader.ename, ""],
    get: (record) =>
      formatSpace(
        getParsedValue(record, _base.ename)?.replace(regex.space, " ")
      ),
    set: (record) => record.ename,
    width: md,
  },
  cname: {
    name: [cname, wayout_cname, "Chinese name", standardHeader.cname, ""],
    get: (record) =>
      formatSpace(
        getParsedValue(record, _base.cname)?.replace(regex.space, " ")
      ),
    set: (record) => record.cname,
  },
  hkId: {
    name: [hkId, standardHeader.hkid, ""],
    get: (record) => {
      let id = removeSpace(
        getParsedValue(record, _base.hkId, [
          regex.hkId,
          regex.hkIdStrict,
        ])?.replace(regex.space, "")
      )

      if (id && !regex.hkIdStrict.test(id)) {
        id = id.replace(/^([A-Z]{1,2}\d{6})([0-9A])$/i, "$1($2)")
      }

      return id
    },
    set: (record) => record.hkId,
  },
  className: {
    name: [
      ...Object.values(_className)
        .map((e) => e.name)
        .flat(),
      "",
    ],
    get: (record) => {
      let className

      const c1 = _className.classname1.get(record)
      const c2 = _className.classname2.get(record)

      if (c1 && c2) className = `TC${c1}${c2}`
      else
        className = removeSpace(
          getParsedValue(record, _base.className, regex.className)
        )

      if (className && regex.className.test(className)) return className
      return weakClassname
    },
    set: (record) => record.classname,
  },
}

const _sc = {
  name: [guide_name, standardHeader.sc],
  get: (record) => formatSpace(getParsedValue(record, _sc)),
  set: (record) => record.guide,
  width: sm,
}

const _vdpId = {
  name: [vdpId, "VDPID"],
  get: (record) => removeSpace(getParsedValue(record, _vdpId)),
  set: (record) => record.vdpId,
}

const _dereg = {
  name: dereg,
  get: (record, className, error, generic, trade) => {
    let isDereg = false

    if (record[dereg] != undefined) {
      return (isDereg = Boolean(record[dereg]))
    }

    if (className != undefined && className != null) {
      let badClassName = !(
        regex.class_pure.test(className) ||
        regex.class_veryPure.test(className) ||
        regex.class_complex.test(className) ||
        regex.class_withProgrammeCode.test(className) ||
        // regex.class_withProgrammeCodeAndClassWithNoSpace.test(className) ||
        regex.class_programmeCodeWithClass.test(className)
      )

      const hasCurseWord =
        regex.no.test(className) || regex.masterNo.test(className)

      if (hasCurseWord) {
        ReturnWarning(
          error,
          `Possibly Deregistered Student, please confirm`,
          errors.critical
        )
        isDereg = true
      } else if (badClassName) {
        if (!regex.no.test(className) && !regex.masterNo.test(className)) {
          ReturnWarning(error, `DVE Class Unexpected Format`, errors.important)
          isDereg = false
        }
      } else if (!hasCurseWord) {
        if (!generic && !trade) {
          ReturnWarning(error, `DVE Class has No tutor`, errors.minor)
        }
      }
    }

    return isDereg
  },
  set: (record) => (record.dereg ? "TRUE" : ""),
  style: "system",
}

//==================================
// Temp DB
//==================================

const _diplomaIdStore = new Map()
const _diplomaNameStore = new Map()

//==================================

function getFromStore(store, key) {
  let obj = store.get(key)
  return obj
}

function parseCampus(campus) {
  if (!campus) return null
  if (/HTI|CCI|ICI/.test(campus)) {
    campus = "HTI/CCI"
  }
  campus = campus.replace(regex.space, "").replace(/POKFULAM/g, "PF")

  if (!regex.campus.test(campus)) campus = `YC(${campus})`

  return campus
}

const _diploma = {
  id: {
    name: [diploma_id, standardHeader.diplomaId],
    get: (record) =>
      removeSpace(getParsedValue(record, _diploma.id)?.toUpperCase()),
    set: (record) => record.programme?.diploma?.id || record.diploma?.id,
  },
  name: {
    name: [diploma_name, headerRemakeNames.diplomaName],
    width: lg,
    get: (record) =>
      removeSpace(
        getParsedValue(record, _diploma.name)
          ?.replace("（", "(")
          ?.replace("）", ")")
      ),
    set: (record) => record.programme?.diploma?.name || record.diploma?.name,
  },

  get: (record) => {
    let diploma = null

    const id = _diploma.id.get(record)
    const name = _diploma.name.get(record)

    if (id) {
      diploma = _diplomaIdStore.get(id)
      if (!diploma && id && name) {
        diploma = { id, name }
        _diplomaIdStore.set(id, diploma)
      }
    }

    return diploma
  },
}

const _campus = {
  name: [programme_campus, standardHeader.campus],
  get: (record) =>
    parseCampus(removeSpace(getParsedValue(record, _campus)?.toUpperCase())),
  set: (record) => record.programme?.campus || record.campus || null,
}

const _wayout_remark = {
  name: ["(Wayout Remark)", wayout_programme],
  get: (record) => formatSpace(getParsedValue(record, _wayout_remark)),
  set: (record) => record.remark,
  width: xl,
  style: "system",
}

const _programme = {
  campus: _campus,
  diploma: _diploma,
  get: (record) => {
    let remark = _wayout_remark.get(record)

    let campus = null,
      diploma = {}

    if (remark && regex.details.test(remark)) {
      remark = `${remark}`
        .match(regex.details)[0]
        .split("-")
        .map((e) => e.replace(regex.space, ""))
      ;[diploma.id, diploma.name, campus] = [
        remark[0]?.toUpperCase(),
        remark
          .slice(1, remark.length - 1)
          .join("-")
          .replace("（", "(")
          .replace("）", ")"),
        remark[remark.length - 1],
      ]
      remark = remark.join("-")
    } else if ((diploma = _diploma.get(record)) != null) {
      campus = _campus.get(record)
    }

    campus = parseCampus(campus)
    const res =
      campus && diploma?.id
        ? { diploma, campus, remark }
        : remark
        ? { remark }
        : null

    return res
  },
}

//remained unchanged, as these headers are parsed seperatly
const [_trade, _generic] = [
  { key: "trade", value: programmeClass_trade },
  { key: "generic", value: programmeClass_generic },
].map(({ key, value }) => ({
  get: (record) => {
    if (checkKeys(record, [value.name, value.email])) {
      const name = record[value.name]
      const email = (record[value.email] + "")
        .replace(regex.space, "")
        .split("@")[0]
        ?.trim()
        ?.toLowerCase()

      if (email && !regex.empty.test(email)) {
        return {
          name: `${name}`.replace(regex.space, "_").replace(/_+/g, " ").trim(),
          email: email,
        }
      }
    }
    return null
  },
  name: {
    name: value.name,
    get: (record) => record[value.name],
    set: (record) => record[key]?.name || null,
    width: sm,
  },
  email: {
    name: value.email,
    get: (record) => record[value.email],
    set: (record) => record[key]?.email || null,
    width: sm,
  },
}))

const _programmeClass = {
  name: {
    name: [standardHeader.dveClass, programmeClass_name2],
    get: (record) =>
      getParsedValue(record, _programmeClass.name)
        ?.replace(regex.space, " ")
        .trim(),
    set: (record) => record.programmeClass_name || null,
  },
  generic: _generic,
  trade: _trade,
}

const _award_year = {
  name: [awardYear, "award year"],
  get: (record) => record[awardYear],
  set: (record) => record.awardYear,
  style: "system",
}

//#endregion

// #region Utils

///=================================================================================================
//
// Utils
//
///=================================================================================================

function checkKey(key, record) {
  if (key == "") return true
  let value = record[key]
  if (value === undefined) return false
  if (value === null) return false
  if (`${value}`.replace(regex.space, "") == "") return false
  return true
}

function checkKeys(record, names, strict = true) {
  return strict
    ? names.every((name) => checkKey(name, record))
    : names.some((name) => checkKey(name, record))
}

//#endregion

// #region Parsing

///=================================================================================================
//
// Excel Parsing
//
///=================================================================================================

//==================================
// Error / Warning
//==================================

function Return(record, error, level, key = "__error") {
  const levelKey = key + "Level"
  const additionalKey = "__additional" + key

  let overwrite = record[levelKey] ? level <= record[levelKey] : true
  if (overwrite) {
    if (record[key])
      record[additionalKey] = record[additionalKey]
        ? record[additionalKey] + record[key]
        : record[key]
    record[key] = error
    record[levelKey] = level
  } else {
    record[additionalKey] = record[additionalKey]
      ? record[additionalKey] + error
      : error
  }
  return record
}

function ReturnError(record, error, level) {
  return Return(record, error, level)
}

function ReturnWarning(record, error, level) {
  return Return(record, error, level, "__warning")
}

const errors = {
  critical: 1,
  important: 2,
  minor: 3,
}

const errorsFlipped = Object.keys(errors).reduce((acc, key) => {
  acc[errors[key]] = key
  return acc
}, {})

//==================================
// Parser
//==================================

const checkProgrammeClass = (error, record) => {
  let hasTeacher = !!record.generic || !!record.trade
  if (
    !record.awardYear &&
    record.__programmeClass &&
    !record.campus &&
    hasTeacher
  ) {
    ReturnWarning(
      error,
      `DVE's Class is present with tutor(s), but ${programme_campus} is missing`,
      errors.important
    )
  }
  // if (
  //   !record.awardYear &&
  //   record.dereg &&
  //   record.programmeClass &&
  //   hasTeacher
  // ) {
  //   ReturnWarning(
  //     error,
  //     `DVE's class doesnt not match the expected format`,
  //     errors.important
  //   )
  // } else if (
  //   !record.awardYear &&
  //   record.dereg &&
  //   record.programmeClass &&
  //   !hasTeacher
  // ) {
  //   ReturnWarning(
  //     error,
  //     `DVE's class has no class tutor, student will be deregistered`,
  //     errors.important
  //   )
  // }
}

const checkEntry = (error, record) => {
  record.__inRange = true

  if (!record.__entry) {
    ReturnWarning(error, `Using default DVE entry`, errors.minor)
    record.year = record.__year
    record.month = 9
  } else if (
    record.year < grantham_start_year ||
    (record.year == grantham_start_year && record.month < grantham_start_month)
  ) {
    record.__inRange = false
    ReturnWarning(
      error,
      `Out of DVE selection range ${grantham_start_year}/${grantham_start_month}`,
      errors.minor
    )
  }
}

const checkId = (error, record) => {
  if (!record.className && !record.cname) {
    ReturnError(
      error,
      `Unique identifier missing, it has no ${className1}/${className2}(TC...) and ${cname}`,
      errors.critical
    )
  }
  // no error so the record falls to the success pool for merging at later stage
  // if (!record.id) {
  //   ReturnError(
  //     error,
  //     `Cannot create ID for this student, either Chinese name or ${className1}/${className2}(TC...) is missing`,
  //     errors.critical
  //   )
  // }
}

const checkHkId = (error, record) => {
  if (record.hkId && !regex.hkId.test(record.hkId)) {
    ReturnWarning(error, `Unexpected HKID format`, errors.minor)
  }
}

const checkClassName = (error, className) => {
  // if (className && !regex.className.test(className)) {
  //   ReturnWarning(
  //     error,
  //     `Teen's class name is in an unexpected format, it should be in the format of TC...`,
  //     errors.minor
  //   )
  // }
}

const checkMasterDereg = (error, record) => {
  let __masterRemark = _master_remark.get(record)

  if (!__masterRemark) return false

  let masterRemark = __masterRemark.split(";")
  masterRemark = masterRemark[masterRemark.length - 1].split(":")
  masterRemark = masterRemark[masterRemark.length - 1]

  if (masterRemark && regex.masterNo.test(masterRemark)) {
    //record.__deregByMaster = true
    ReturnWarning(
      error,
      `Student deregistered by 總表 其他評語, please confirm`,
      errors.important
    )
    return true
  }
  //record.__deregByMaster = false
  return false
}

export const checker = (record) => {
  const error = {
    __error: null,
    __errorLevel: null,
    __additional__error: null,
    __warning: null,
    __warningLevel: null,
    __additional__warning: null,
  }

  let tempClass = record.__programmeClass
  if (record.__programmeClass) {
    let split = `${record.__programmeClass}`.split(";")
    let parsedFile = split[split.length - 1].split(":")
    let parsed = parsedFile[parsedFile.length - 1].trim()
    record.__programmeClass = parsed
  }

  record.dereg =
    record.campus && record.__programmeClass
      ? _dereg.get(
          record,
          record.__programmeClass,
          error,
          record.generic,
          record.trade
        ) || checkMasterDereg(error, record)
      : false

  if (record.dereg) {
    record.programmeClass_name = null
  }

  checkId(error, record)
  checkEntry(error, record)
  checkProgrammeClass(error, record)

  record.__programmeClass = tempClass

  Object.assign(record, error)
}

let count__ = 0

export const solver = (
  index,
  row,
  _year,
  dve,
  award,
  file,
  type,
  operation
) => {
  let awardYear = (award ? _year : null) || _award_year.get(row)
  if (dve) {
    row[enrollment_eng] = row[enrollment] = _year.trim()
  }

  // if (count__ < 1) console.log(row)

  let entry = _base.entryDate.get(row)

  let chi_Name = _base.cname.get(row)
  let eng_name = _base.ename.get(row)
  let hkId = _base.hkId.get(row)
  let vdpId = _vdpId.get(row)
  let guide = _sc.get(row)

  let error = {}

  let className = _base.className.get(row)
  checkClassName(error, className)

  let id

  if (chi_Name && className && className != weakClassname) {
    id = className + "-" + chi_Name
  }

  const masterRemark = _master_remark.get(row)
  const avgMark = _avg_mark.get(row)

  const programme = _programme.get(row)
  const diploma = _diploma.get(row)
  const campus = _campus.get(row)
  const programmeClass = _programmeClass.name.get(row)

  let record = {
    id,
    ...entry,
    __entry: entry,
    __file: {
      [file]: [index],
    },
    __avg_mark: avgMark,
    __remark: masterRemark,
    __operation: operation,
    __year: _year,
    __type: type,
    __programmeClass: programmeClass,
    programmeClass_name: programmeClass,
    remark: programme?.remark || row.remark, //...i cant remember why
    hkId,
    vdpId,
    guide,
    cname: chi_Name,
    ename: eng_name,
    //prevent weakClassname from overwriting the actual classname
    //of the dupe record within the same file
    className: className == weakClassname ? null : className,
    diploma: diploma || programme?.diploma,
    campus: campus || programme?.campus,
    awardYear,
  }

  // if (count__ < 1) console.log(record)

  let [generic, trade] = [_generic, _trade].map((e) => e.get(row))

  record.dereg =
    campus && programmeClass
      ? _dereg.get(row, programmeClass, error, generic, trade) ||
        checkMasterDereg(error, record)
      : false

  if (!record.dereg) {
    record.programmeClass_name = programmeClass
    if (record.programmeClass_name && diploma?.id) {
      const code = diploma.id
      const className = record.programmeClass_name
      const reg = new RegExp(`^${code}`, "i")

      if (reg.test(className) && !regex.class_complex.test(className)) {
        record.programmeClass_name = className
          .replace(reg, "")
          .replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, "")
          .trim()
      }
    }

    // I wish i got a list of all the programme codes
    // if (record.diploma?.id && record.programmeClass_name) {
    //   record.programmeClass_name = record.programmeClass_name
    //     .split(" ")
    //     .filter(
    //       (c) =>
    //         c &&
    //         !(
    //           c.includes(record.diploma.id) &&
    //           c.length > record.diploma.id.length * 0.7
    //         )
    //     )
    //     .join(" ")
    // } else {
    //   // console.log("no programme class")
    // }

    record.trade = trade
    record.generic = generic
  } else {
    record.programmeClass_name = null
  }

  checkId(error, record)
  checkProgrammeClass(error, record)

  Object.assign(record, error)

  record.__original = [{ [file]: row }]

  if (record.__error || record.__warning)
    return [
      record,
      record.__error,
      errorsFlipped[record.__errorLevel],
      record.__warning,
      errorsFlipped[record.__warningLevel],
      error,
    ]

  count__++

  return [record]
}

// #endregion

// #region Export

export const header = Array.from(
  new Set(
    [
      ...Object.values(_base),
      ...Object.values(_className),
      _vdpId,
      _dereg,
      _sc,
      ...Object.values(_diploma),
      _campus,
      _programmeClass.name,
      ...Object.values(_trade),
      ...Object.values(_generic),
      _award_year,
    ]
      .reduce((prev, curr) => {
        if (Array.isArray(curr.name))
          prev.push(
            ...curr.name.filter((e) => {
              return Boolean(e) && typeof e === "string"
            })
          )
        else if (
          curr.name &&
          typeof curr.name === "string" &&
          typeof curr !== "function"
        ) {
          prev.push(curr.name)
        }
        return prev
      }, [])
      .filter((h) => typeof h !== "function")
  )
)

// #endregion

///=================================================================================================
//
// Excel Export
//
///=================================================================================================

function createGetFn(accessorKey) {
  const keys = accessorKey.split(".")

  return (row) => {
    let value = row
    for (const key of keys) {
      if (value === undefined) {
        break
      }
      value = value[key]
    }
    return value
  }
}

const excelWidthScale = 2.2
const worksheetColumns = (headers, scale = excelWidthScale) => {
  let res = Object.entries(headers).map(([key, header]) => ({
    header: Array.isArray(header.name) ? header.name[0] : header.name,
    key: key,
    width: (header.width || xs) * scale,
    _style: header.style || undefined,
    get: header.set || header.get || createGetFn(key),
  }))
  return res
}

const _system = {
  operation: {
    name: "(Operation)",
    set: (record) => record.__operation,
    width: xs,
    style: "system",
  },
  type: {
    name: "(Type)",
    set: (record) => record.__type,
    width: xs,
    style: "system",
  },
  filename: {
    name: "(File)",
    set: (record) => {
      let files = Object.keys(record.__file).reduce((acc, key) => {
        var name = key
        var indexes = record.__file[key]
        acc += `${name}[${indexes.join(", ")}]; `
        return acc
      }, "")
      return files.slice(0, -2)
    },
    width: xs,
    style: "system",
  },
}

const _master_remark = {
  name: ["(總表 其他評語)", standardHeader.masterRemark],
  set: (record) => record.__remark,
  get: (record) => {
    if (record.__remark) return record.__remark

    return removeSpace(getParsedValue(record, _master_remark))
  },
  style: "system",
}

const _avg_mark = {
  name: ["(總分)", standardHeader.avgMark],
  width: xs,
  style: "system",
  get: (record) => removeSpace(getParsedValue(record, _avg_mark)),
  set: (record) => {
    if (!record?.__avg_mark) return null

    try {
      const mark = parseFloat(record.__avg_mark).toFixed(2)
      return mark
    } catch (e) {
      return null
    }
  },
}

const _original = {
  original: {
    name: "(Original)",
    set: (record) => record.__original,
    width: xs,
    style: "system",
  },
}

const _id = {
  id: { name: "(ID)", set: (record) => record.id, style: "system" },
}

const _oriProgrammeClass = {
  originalProgrammeClass: {
    name: `(總表 DVE Class)`,
    set: (record) => record.__programmeClass || null,
    style: "system",
  },
}

const _warning = {
  warning: {
    name: "(Warning)",
    set: (record) => record.__warning || null,
    width: xl,
    style: "warning",
  },
  warningLevel: {
    name: "(Warning Level)",
    set: (record) =>
      record.__warningLevel ? errorsFlipped[record.__warningLevel] : null,
    style: "warning",
  },
  additionalWarning: {
    name: "(Additional Warning)",
    set: (record) => record.__additional__warning || null,
    width: xl,
    style: "warning",
  },
}

const _error = {
  error: {
    name: "(Error)",
    set: (record) => record.__error || null,
    width: xl,
    style: "error",
  },
  errorLevel: {
    name: "(Error Level)",
    set: (record) =>
      record.__errorLevel ? errorsFlipped[record.__errorLevel] : null,
    style: "error",
  },
  additionalError: {
    name: "(Additional Error)",
    set: (record) => record.__additional__error || null,
    width: xl,
    style: "error",
  },
}

const baseExportSchema = {
  enrollment: _base.entryDate,
  ename: _base.ename,
  cname: _base.cname,
  hkId: _base.hkId,
  ..._className,
  guide: _sc,
  campus_id: _campus,
  diploma_id: _diploma.id,
  diploma_name: _diploma.name,
}

const namelessProgrammeClassSchema = {
  programmeClass_generic_name: _generic.name,
  programmeClass_generic_email: _generic.email,
  programmeClass_trade_name: _trade.name,
  programmeClass_trade_email: _trade.email,
}
const programmeClassSchema = {
  programmeClass_name: _programmeClass.name,
  ...namelessProgrammeClassSchema,
}

const allExportSchema = {
  enrollment: _base.entryDate,
  ename: _base.ename,
  cname: _base.cname,
  hkId: _base.hkId,
  ..._className,
  guide: _sc,
  campus_id: _campus,
  diploma_id: _diploma.id,
  diploma_name: _diploma.name,
  remark: _wayout_remark,
  programmeClass_name: _programmeClass.name,
  ..._oriProgrammeClass,
  dereg: _dereg,
  master_remark: _master_remark,
  avg_mark: _avg_mark,
  programmeClass_generic_name: _generic.name,
  programmeClass_generic_email: _generic.email,
  programmeClass_trade_name: _trade.name,
  programmeClass_trade_email: _trade.email,
  awardYear: _award_year,
  vdpId: _vdpId,
}

const debugSuccessSchema = {
  ...allExportSchema,
  ..._warning,
  ..._id,
  ..._original,
  ..._system,
}

const debugNoDupeSchema = {
  ...allExportSchema,
  ..._warning,
  ..._id,
  ..._original,
  ..._system,
}
delete debugNoDupeSchema.awardYear

const debugAwardSchema = {
  ...debugNoDupeSchema,
}
delete debugAwardSchema.dereg

const debugFailSchema = {
  // ..._fragileId,
  ...allExportSchema,
  ..._error,
  ..._id,
  ..._original,
  ..._system,
}
delete debugFailSchema.awardYear

const addName = (schema, blank) => {
  return Object.keys(schema).reduce((acc, curr) => {
    acc[curr] = { ...schema[curr] }
    acc[curr].name = [standardHeader[curr]]

    if (blank) {
      delete acc[curr].get
      acc[curr].set = () => ""
      delete acc[curr].style
    }

    return acc
  }, {})
}

//to counter width multiplier
//this is so sad
const normalizeWidth = (schema, scale) => {
  return Object.keys(schema).reduce((acc, curr) => {
    acc[curr] = { ...schema[curr], width: schema[curr].width / scale }
    return acc
  }, {})
}

const markingSchema = {
  politeness: {
    name: ["整體禮貌"],
    width: sm,
  },
  attitudes: {
    name: ["學習態度(面授及網教)"],
    width: sm,
  },
  responsiblity: {
    name: ["責任感"],
    width: sm,
  },
  friendliness: {
    name: ["待人接物"],
    width: sm,
  },
  attentiveness: {
    name: ["服務他人"],
    width: sm,
  },
  accScore: {
    name: ["學業成績"],
    width: sm,
  },
}

export const dveTeacherMarkingSchema = [
  {
    name: "spacer",
    width: 10,
  },
  {
    name: "學員姓名",
    width: 17,
    type: "blue",
  },
  {
    name: "課程編號",
    width: 17,
    type: "blue",
  },
  {
    name: "現時就讀班別",
    width: 17,
    type: "blue",
  },
  {
    name: "分校",
    width: 17,
    type: "blue",
  },
  {
    name: "填寫老師姓名",
    width: 24,
    type: "gray",
  },
  {
    name: "填寫老師聯絡電話",
    width: 24,
    type: "gray",
  },
  {
    name: "填寫老師電郵",
    width: 24,
    type: "gray",
  },
  {
    name: "不推薦 (請在以下填寫原因)",
    width: 56,
    type: "no",
  },
  {
    name: "學生整體出席率 % \n(請填實際數值)",
    width: 21.11,
    type: "yes",
  },
  ...Object.values(markingSchema).map((s) => ({
    name:
      s.name[0]
        //lol
        .replace("(", "\n(") + "\n (1-10分)",
    width: 14,
    type: "yes",
  })),
]

const scMasterSchema = {
  ...addName({
    dveEntry: _base.entryDate,
    ename: _base.ename,
    cname: _base.cname,
    hkid: _base.hkId,
    tsEntry: _className.classname1,
    tsClass: _className.classname2,
    sc: _sc,
    campus: _campus,
    diplomaId: _diploma.id,
    diplomaName: _diploma.name,
    dveClass: _programmeClass.name,
    dveTTn: _programmeClass.trade.name,
    dveTTe: _programmeClass.trade.email,
    dveTGn: _programmeClass.generic.name,
    dveTGe: _programmeClass.generic.email,
  }),

  ...normalizeWidth(
    {
      attendance: {
        name: ["出席率 (佔50%)"],
        width: sm,
      },
      ...markingSchema,
      avgSix: {
        name: ["6項之平均分"],
        width: sm,
      },
      avg: {
        name: ["平均分(佔50%)"],
        width: sm,
      },
    },
    1.8
  ),

  ...addName(
    {
      avgMark: { ..._avg_mark, width: sm },
      masterRemark: { ..._master_remark, width: xl },
    },
    true
  ),

  remark: {
    name: ["備註"],
    width: xl,
    set: () => "",
  },

  ..._oriProgrammeClass,
  dereg: _dereg,
  ..._warning,
  ..._id,
  ..._original,
  ..._system,
}

// console.log(scMasterSchema)

export const exportSchema = {
  all: worksheetColumns({
    vdpId: _vdpId,
    ...baseExportSchema,
    awardYear: _award_year,
    dereg: _dereg,
  }),
  scMasterSchema: worksheetColumns(scMasterSchema),
  debugAwardSchema: worksheetColumns(debugAwardSchema),
  debugNoDupeSchema: worksheetColumns(debugNoDupeSchema),
  debugSuccessSchema: worksheetColumns(debugSuccessSchema),
  debugFailSchema: worksheetColumns(debugFailSchema),
  contactsc: worksheetColumns({ ...baseExportSchema, ...programmeClassSchema }),
}
