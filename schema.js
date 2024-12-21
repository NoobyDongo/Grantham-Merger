import crypto from 'crypto'

// #region Definitions

///=================================================================================================
//
// Definitions
//
///=================================================================================================

const enrollment = '期數',
  enrollment_eng = 'DVE Entry',
  ename = '英文姓名',
  cname = '中文姓名',
  wayout_cname = '學生姓名',
  hkId = '身份證號碼',
  className1 = '學期',
  className2 = '班別',
  className = '班別',
  vdpId = 'VDP 編號',
  guide_name = '學生輔導主任',
  diploma_id = '課程編號',
  diploma_name = '課程名稱',
  programme_campus = '上課地點',
  wayout_programme = '詳情',
  master_remark = '評語',
  programmeClass_name = 'Class',
  programmeClass_name2 = '(總表 DVE Class)',
  programmeClass_generic = {
    name: 'Class Tutor (Generic, Fullname)',
    email: 'Email (Generic)'
  },
  programmeClass_trade = {
    name: 'Class Tutor (Trade, Fullname)',
    email: 'Email (Trade)'
  },
  awardYear = '(得獎年份)',
  dereg = '(Dereg/Grad)'

export const remarkRegex = new RegExp(master_remark, 'g')

const weakClassname = 'TC00N00AA0AA'

//==================================
// Regex
//==================================

const regex = {
  className1_1: new RegExp('^\\d{2}[A-Z]{1}\\d{2}$', 'i'),
  className1_2: new RegExp('^[A-Z]{2,3}\\d{1}[A-Z]{2}$', 'i'),

  className: new RegExp('^TC\\d{2,3}[A-Z]\\d{2}[A-Z]{2,3}\\d[A-Z]{2}$', 'i'),

  //18N01 OATM? && 18N02 SFTO from master 2020
  classNameRare: new RegExp('^[A-Z]{4}$', 'i'),

  //around 2015 from award
  className3_1: new RegExp('^\\d{2}[A-Z]{1}\\d{2}$', 'i'),
  className3_2: new RegExp('^[A-Z]{2,3}\\d{1}[A-Z]{2,3}$', 'i'),

  //around 2010 from award
  className2_1: new RegExp('^\\d{2,3}[A-Z]{2,3}$', 'i'),
  className2_2: new RegExp('^[A-Z]{1,3}$', 'i'),

  entry: new RegExp('^\\d{4}/\\d{2}$'),

  details: new RegExp(
    '[A-Z]{2}\\d{6}[A-Z]?\\s?-\\s?.+\\s?-\\s?(IVE\\([A-Z]{2,3}\\)|YC\\([A-Z]{2,3}\\)|CCI|HKDI)',
    'i'
  ),

  // 818/16/pc, 404/22/PDFB
  class_complex: new RegExp('.{2,3}\\/.{2}\\/.{2,4}'),
  // AB123456, AB123456A, AB123456/12/AB
  class_withProgrammeCode: new RegExp(
    '^[A-Z]{2}\\d{6}[A-Z]?(-|\\s|/\\d{2}(/[A-Z]{2})?)?$',
    'i'
  ),
  // AB123456/AB, AB123456/ABCD, AB123456/11AB, AB123456A/1D, AB123456A/A1D, AB123456/ddd2a
  class_programmeCodeWithClass: new RegExp(
    '^[A-Z]{2}\\d{6}[A-Z]?/.{2,5}$',
    'i'
  ),
  // A1A, V12c, a-1d, d1, A11A, 1A, 11A
  class_pure: new RegExp('^([A-Z]|[A-Z]-)?\\d{1,2}[A-Z][A-Z]?$', 'i'),
  class_veryPure: new RegExp('^[A-Z]{1}$', 'i'),

  programmeCode_pure: new RegExp('^[A-Z]{2}\\d{6}$', 'i'),
  programmeCode_variant: new RegExp('^[A-Z]{2}\\d{6}[A-Z]{1,2}$', 'i'),
  programmeName_variant: new RegExp('^.*[A-Z]{1,2}$', 'i'),

  no: /(dereg|grad|no|de-reg|defer|quit|withdraw)/i,

  //so sad
  masterNo:
    /(dereg|de-reg|defer|quit|withdraw|employ|退學(?!未)|no record|no this student|no data|no student|not found|unable to locate)/i,

  hkIdStrict: new RegExp('^[A-Z]{1,2}\\d{6}\\([0-9A]\\)$', 'i'),
  hkId: new RegExp('^[A-Z]{1,2}\\d{6}\\(?[0-9A]\\)?$', 'i'),

  empty: new RegExp('n[/\\\\]?a', 'i'),
  space: new RegExp(
    '[\t\n\v\f\r \u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u2028\u2029\u3000]+'
  ),

  campus: new RegExp('^HTI|CCI|ICI|YC\\(.*\\)|IVE\\(.*\\)$', 'i')
}

//==================================
// Accessors + validators
//==================================

const xl = 20,
  lg = 15,
  md = 10,
  sm = 7.5,
  xs = 5

const _base = {
  entryDate: {
    name: [enrollment_eng, enrollment, ''],
    get: record => {
      let entry
      if (record[enrollment] && regex.entry.test(record[enrollment])) {
        entry = record[enrollment]
      } else if (
        record[enrollment_eng] &&
        regex.entry.test(record[enrollment_eng])
      ) {
        entry = record[enrollment_eng]
      }
      if (!entry) return null
      entry = entry.split('/')
      return {
        year: entry[0],
        month: entry[1]
      }
    },
    set: record =>
      record.year && record.month
        ? record.year + '/' + record.month.toString().padStart(2, '0')
        : null
  },
  ename: {
    name: [ename, ''],
    get: record =>
      record[ename]
        ? `${record[ename]}`.trim().replace(regex.space, ' ')
        : null,
    set: record => record.ename,
    width: md
  },
  cname: {
    name: [cname, wayout_cname, ''],
    get: record =>
      record[cname]
        ? record[cname]
          ? `${record[cname]}`.trim().replace(regex.space, '')
          : null
        : record[wayout_cname]
        ? `${record[wayout_cname]}`?.trim().replace(regex.space, '')
        : null,
    set: record => record.cname
  },
  hkId: {
    name: [hkId, ''],
    get: record => {
      let id = record[hkId]
        ? `${record[hkId]}`.toUpperCase().replace(regex.space, '')
        : null
      if (regex.hkId.test(id)) {
        if (!regex.hkIdStrict.test(id)) {
          id = id.replace(/^([A-Z]{1,2}\d{6})([0-9A])$/i, '$1($2)')
        }

        return id
      }
      return null
    },
    set: record => record.hkId
  },
  className: {
    name: [className1, className2, ''],
    get: record => {
      let className
      if (record[className1]) {
        className = `${record[className1]}`?.replace(regex.space, '')
      } else if (record[className2]) {
        className = `${record[className2]}`?.replace(regex.space, '')
      } else return weakClassname

      if (regex.className.test(className)) return className
      else if (checkKeys(record, [className1, className2])) {
        const c1 = `${record[className1]}`?.replace(regex.space, '')
        const _c2 = `${record[className2]}`?.replace(regex.space, '')
        const c2 = regex.classNameRare.test(_c2)
          ? _c2.slice(0, 2) + '1' + _c2.slice(2)
          : _c2

        className = 'TC' + c1 + c2
        if (
          (regex.className3_1.test(c1) && regex.className3_2.test(c2)) ||
          (regex.className2_1.test(c1) && regex.className2_2.test(c2)) ||
          (regex.className1_1.test(c1) && regex.className1_2.test(c2))
        )
          return className
      }
      return weakClassname
    },
    set: record => record.classname
  }
}

const _guide = {
  name: guide_name,
  get: record => {
    let name = record[guide_name]
    if (name) {
      name = `${name}`.replace(regex.space, ' ')
      return { name }
    }
    return null
  },
  set: record => record.guide?.name,
  width: sm
}

const _vdpId = {
  name: vdpId,
  get: record => record[vdpId],
  set: record => record.vdpId
}

const _dereg = {
  name: dereg,
  get: (className, error, generic, trade) => {
    let isDereg = true

    if (className != undefined && className != null) {
      let badClassName = !(
        regex.class_pure.test(className) ||
        regex.class_veryPure.test(className) ||
        regex.class_complex.test(className) ||
        regex.class_withProgrammeCode.test(className) ||
        regex.class_programmeCodeWithClass.test(className)
      )

      if (badClassName) {
        if (!regex.no.test(className)) {
          ReturnWarning(error, `DVE Class Unexpected Format`, errors.important)
          isDereg = false
        } else {
          ReturnWarning(
            error,
            `Possibly Deregistered Student, please confirm`,
            errors.critical
          )
          isDereg = true
        }
      } else {
        if (!generic && !trade) {
          ReturnWarning(error, `DVE Class has No tutor`, errors.minor)
        }
        isDereg = false
      }
    }

    return isDereg
  },
  set: record => (record.dereg ? 'TRUE' : ''),
  style: 'system'
}

//==================================
// Temp DB
//==================================

const _diplomaIdStore = new Map()
const _diplomaNameStore = new Map()

//==================================

function getFromStore (store, key) {
  let obj = store.get(key)
  return obj
}

function parseCampus (campus) {
  if (!campus) return null
  if (/HTI|CCI|ICI/.test(campus)) {
    campus = 'HTI/CCI'
  }
  campus = campus.replace(regex.space, '').replace(/POKFULAM/g, 'PF')

  if (!regex.campus.test(campus)) campus = `YC(${campus})`

  return campus
}

const _diploma = {
  get: record => {
    if (checkKeys(record, [diploma_id, diploma_name], false)) {
      let diploma
      let id = record[diploma_id] ? `${record[diploma_id]}`.toUpperCase() : null
      let name = record[diploma_name]
        ? `${record[diploma_name]}`.replace('（', '(').replace('）', ')')
        : null
      let shortenIndexFlag = { flag: false }

      if (id) {
        diploma = getFromStore(_diplomaIdStore, id, shortenIndexFlag)
        if (!diploma) {
          diploma = name
            ? getFromStore(_diplomaNameStore, name, shortenIndexFlag)
            : null
          if (!diploma) {
            diploma = { id, name }
          } else {
            if (!diploma.id) diploma.id = id
            if (!diploma.name && name) diploma.name = name
          }
          _diplomaIdStore.set(id, diploma)
          if (name) _diplomaNameStore.set(name, diploma)
        } else {
          if (!diploma.name && name) diploma.name = name
          if (name) _diplomaNameStore.set(name, diploma)
        }
      } else if (name) {
        diploma = getFromStore(_diplomaNameStore, name)
        if (!diploma) {
          diploma = { id: null, name }
          _diplomaNameStore.set(name, diploma)
        }
      }

      return diploma
    }
    return null
  },
  id: {
    name: diploma_id,
    get: record => record[diploma_id],
    set: record => record.programme?.diploma?.id || record.diploma?.id
  },
  name: {
    name: diploma_name,
    width: lg,
    get: record => record[diploma_name],
    set: record => record.programme?.diploma?.name || record.diploma?.name
  }
}

const _campus = {
  name: programme_campus,
  get: record => {
    if (!record[programme_campus]) return null

    let campus = `${record[programme_campus]}`
      ?.toUpperCase()
      .replace(regex.space, '')

    campus = parseCampus(campus)

    return {
      id: campus
    }
  },
  set: record => record.programme?.campus?.id || record.campus?.id
}

const _programme = {
  get: record => {
    let details = record[wayout_programme],
      campus = {},
      diploma = {}

    if (details) {
      let passed = regex.details.test(details)
      if (passed) {
        details = `${details}`
          .match(regex.details)[0]
          .split('-')
          .map(e => e.replace(regex.space, ''))
        ;[diploma.id, diploma.name, campus.id] = [
          details[0]?.toUpperCase(),
          details
            .slice(1, details.length - 1)
            .join('-')
            .replace('（', '(')
            .replace('）', ')'),
          details[details.length - 1]
        ]
      }
      record.remark = record[wayout_programme]
    } else if ((diploma = _diploma.get(record)) != null) {
      campus = _campus.get(record)
    }
    campus.id = parseCampus(campus.id)
    // if(!(campus?.id && diploma?.id) && (campus?.id || diploma?.id)) {
    //     console.log(record, campus, diploma)
    //     throw new Error('Campus or Diploma ID missing')
    // }
    return campus?.id && diploma?.id ? { diploma, campus } : null
  },
  campus: _campus,
  diploma: _diploma
}

const [_trade, _generic] = [
  { key: 'trade', value: programmeClass_trade },
  { key: 'generic', value: programmeClass_generic }
].map(({ key, value }) => ({
  get: record => {
    if (checkKeys(record, [value.name, value.email])) {
      const name = record[value.name]
      const email = (record[value.email] + '')
        .split('@')[0]
        ?.trim()
        ?.toLowerCase()
      if (email && !(regex.empty.test(email) || regex.space.test(email))) {
        return {
          name: `${name}`.replace(regex.space, '_').replace(/_+/g, ' ').trim(),
          email: email
        }
      }
    }
    return null
  },
  name: {
    name: value.name,
    get: record => record[value.name],
    set: record => record[key]?.name || null,
    width: sm
  },
  email: {
    name: value.email,
    get: record => record[value.email],
    set: record => record[key]?.email || null,
    width: sm
  }
}))

const _programmeClass = {
  name: {
    name: [programmeClass_name, programmeClass_name2],
    get: record => {
      let name = record[programmeClass_name] || record[programmeClass_name2]
      if (name) {
        name = `${name}`.replace(regex.space, '')
        return name
      }
      return null
    },
    set: record =>
      record.programmeClass_name ||
      // || record.programmeClass
      null
  },
  generic: _generic,
  trade: _trade
}

//#endregion

// #region Utils

///=================================================================================================
//
// Utils
//
///=================================================================================================

function hashStringSync (message) {
  return crypto.createHash('sha256').update(message).digest('hex')
}

function checkKey (key, record) {
  if (key == '') return true
  let value = record[key]
  if (value === undefined) return false
  if (value === null) return false
  if (`${value}`.replace(regex.space, '') == '') return false
  return true
}

function parsefieldNames (fields) {
  return fields
    .filter(name => !!name)
    .map(name => `{${name}}`)
    .join('/')
}

function checkKeys (record, names, strict = true) {
  return strict
    ? names.every(name => checkKey(name, record))
    : names.some(name => checkKey(name, record))
}

function checkEachKey (record, checkers, strict = true) {
  let errorFields = new Set()
  for (const names of checkers) {
    let valid = strict
      ? names.every(name => checkKey(name, record))
      : names.some(name => checkKey(name, record))

    if (!valid) {
      errorFields.add(parsefieldNames(names))
    }
  }
  return errorFields
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

function Return (record, error, level, key = '__error') {
  const levelKey = key + 'Level'
  const additionalKey = '__additional' + key

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

function ReturnError (record, error, level) {
  return Return(record, error, level)
}

function ReturnWarning (record, error, level) {
  return Return(record, error, level, '__warning')
}

const errors = {
  critical: 1,
  important: 2,
  minor: 3
}

const errorsFlipped = Object.keys(errors).reduce((acc, key) => {
  acc[errors[key]] = key
  return acc
}, {})

//==================================
// Parser
//==================================

const baseChecker = Object.keys(_base).map(key => _base[key].name)

const checkProgrammeClass = (error, record) => {
  let hasTeacher = !!record.generic || !!record.trade
  if (
    !record.awardYear &&
    record.__programmeClass &&
    !record.campus?.id &&
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
  // if (!record.awardYear && (!record.year || !record.month) && !record.year === 0) {
  //   ReturnError(
  //     error,
  //     `${enrollment_eng}/${enrollment} is Missing/Incorrect (YYYY/MM, eg:2024/09), student will not get nominated`,
  //     errors.minor
  //   )
  // }
}

const checkId = (error, record) => {
  if (!record.className && !record.cname) {
    ReturnError(
      error,
      `Unique identifier missing, it has no ${className1}/${className2}(TC...) and ${cname}`,
      errors.critical
    )
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

  if(!__masterRemark) return false
  
  let masterRemark = __masterRemark.split(";")
  masterRemark = masterRemark[masterRemark.length - 1].split(":")
  masterRemark = masterRemark[masterRemark.length - 1]

  if (masterRemark && regex.masterNo.test(masterRemark)) {
    record.__deregByMaster = true
    ReturnWarning(
      error,
      `Student deregistered by 總表 其他評語, please confirm`,
      errors.important
    )
    return true
  }
  record.__deregByMaster = false
  return false
}

export const checker = record => {
  const error = {
    __error: null,
    __errorLevel: null,
    __additional__error: null,
    __warning: null,
    __warningLevel: null,
    __additional__warning: null
  }

  let tempClass = record.__programmeClass
  if (record.__programmeClass) {
    let split = `${record.__programmeClass}`.split(';')
    let parsedFile = split[split.length - 1].split(':')
    let parsed = parsedFile[parsedFile.length - 1].trim()
    record.__programmeClass = parsed
  }

  record.dereg =
    record.campus?.id && record.__programmeClass
      ? _dereg.get(
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
    //console.log(_year)
  }
  let entry = _base.entryDate.get(row) || { year: _year, month: 9 }

  //let entry = _base.entryDate.get(row) || { year, month: 9 } //assume september if not provided
  let chi_Name = _base.cname.get(row)
  let eng_name = _base.ename.get(row)
  let hkId = _base.hkId.get(row)
  let vdpId = _vdpId.get(row)
  let guide = _guide.get(row)

  let error = {}

  let baseErrors = checkEachKey(row, baseChecker, false)
  if (baseErrors.size > 0) {
    ReturnError(
      error,
      'Missing required fields: ' + Array.from(baseErrors).join(', '),
      errors.important
    )
  }

  let className = _base.className.get(row)
  checkClassName(error, className)

  let id,
    name = chi_Name,
    fragileId

  if (name && className && className != weakClassname) {
    id = className + '-' + name
  }
  // else if (hkId) {
  //   id = hkId
  //   fragileId = true
  //   ReturnError(
  //     error,
  //     'Unique identifier missing, it has no a combination of class name(TC...) and name',
  //     //`Teen's class name(${className2}, eg: TC...)(${className1}/${className2}, eg: 22N03ACG1DD) or name(${cname}/${ename}) are missing, it is now using HKID as the unique identifier`,
  //     errors.critical
  //   )
  // }

  const masterRemark = _master_remark.get(row)
  const programme = _programme.get(row)
  const diploma = _diploma.get(row)
  const campus = _campus.get(row)
  const programmeClass = _programmeClass.name.get(row)

  let record = {
    id,
    ...entry,
    __file: {
      [file]: [index]
    },
    __remark: masterRemark,
    __operation: operation,
    __year: _year,
    __type: type,
    __programmeClass: programmeClass,
    programmeClass_name: programmeClass,
    remark: row.remark,
    hkId,
    vdpId,
    guide,
    cname: chi_Name,
    ename: eng_name,
    className: className,
    diploma: diploma || programme?.diploma,
    campus: campus || programme?.campus,
    awardYear
  }

  if (programme?.remark) {
    record.remark = programme.remark
  }

  let [generic, trade] = [_generic, _trade].map(e => e.get(row))

  record.dereg =
    campus?.id && programmeClass
      ? _dereg.get(programmeClass, error, generic, trade) ||
        checkMasterDereg(error, record)
      : false

  if (!record.dereg) {
    record.programmeClass_name = programmeClass
    record.trade = trade
    record.generic = generic
  } else {
    record.programmeClass_name = null
  }

  checkEntry(error, record)
  checkId(error, record)
  checkProgrammeClass(error, record)

  Object.assign(record, error)

  record.__bad_id = !!fragileId
  record.__original = row

  if (record.__error || record.__warning)
    return [
      record,
      record.__error,
      errorsFlipped[record.__errorLevel],
      record.__warning,
      errorsFlipped[record.__warningLevel],
      error
    ]

  return [record]
}

// #endregion

// #region Export

export const header = [
  enrollment,
  ename,
  cname,
  wayout_cname,
  hkId,
  className1,
  className2,
  className,
  vdpId,
  dereg,
  guide_name,
  diploma_id,
  diploma_name,
  programme_campus,
  wayout_programme,
  programmeClass_name,
  Object.keys(programmeClass_generic).map(key => programmeClass_generic[key]),
  Object.keys(programmeClass_trade).map(key => programmeClass_trade[key]),
  awardYear
]
// #endregion

///=================================================================================================
//
// Excel Export
//
///=================================================================================================

function createGetFn (accessorKey) {
  const keys = accessorKey.split('.')

  return row => {
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
const worksheetColumns = (headers, block = []) => {
  let res = Object.entries(headers).map(([key, header]) => ({
    header: Array.isArray(header.name) ? header.name[0] : header.name,
    key: key,
    width: (header.width || xs) * excelWidthScale,
    _style: header.style || undefined,
    get: header.set || header.get || createGetFn(key)
  }))
  return res
}

const _system = {
  operation: {
    name: '(Operation)',
    set: record => record.__operation,
    width: xs,
    style: 'system'
  },
  type: {
    name: '(Type)',
    set: record => record.__type,
    width: md,
    style: 'system'
  },
  filename: {
    name: '(File)',
    set: record => {
      let files = Object.keys(record.__file).reduce((acc, key) => {
        var name = key
        var indexes = record.__file[key]
        acc += `${name}[${indexes.join(', ')}]; `
        return acc
      }, '')
      return files.slice(0, -2)
    },
    width: xl,
    style: 'system'
  }
}

const _remark = {
  name: '(Wayout Remark)',
  set: record => record.remark,
  style: 'system'
}

const _master_remark = {
  name: '(總表 其他評語)',
  set: record => record.__remark,
  get: record => record.__remark,
  style: 'system'
}

const _award_year = {
  name: awardYear,
  get: record => record[awardYear],
  set: record => record.awardYear,
  style: 'system'
}

const _className = {
  classname1: {
    name: className1,
    set: record => record.className?.slice(2, 7)
  },
  classname2: {
    name: className2,
    set: record => record.className?.slice(7)
  }
}

const _original = {
  original: {
    name: '(Original)',
    set: record => record.__original,
    width: md,
    style: 'system'
  }
}

const _id = {
  id: { name: '(ID)', set: record => record.id, style: 'system' }
}

const _oriProgrammeClass = {
  originalProgrammeClass: {
    name: `(總表 DVE Class)`,
    set: record => record.__programmeClass || null,
    style: 'system'
  }
}

const _warning = {
  warning: {
    name: '(Warning)',
    set: record => record.__warning || null,
    width: xl,
    style: 'warning'
  },
  warningLevel: {
    name: '(Warning Level)',
    set: record =>
      record.__warningLevel ? errorsFlipped[record.__warningLevel] : null,
    style: 'warning'
  },
  additionalWarning: {
    name: '(Additional Warning)',
    set: record => record.__additional__warning || null,
    width: xl,
    style: 'warning'
  }
}

const _error = {
  error: {
    name: '(Error)',
    set: record => record.__error || null,
    width: xl,
    style: 'error'
  },
  errorLevel: {
    name: '(Error Level)',
    set: record =>
      record.__errorLevel ? errorsFlipped[record.__errorLevel] : null,
    style: 'error'
  },
  additionalError: {
    name: '(Additional Error)',
    set: record => record.__additional__error || null,
    width: xl,
    style: 'error'
  }
}

const baseExportSchema = {
  enrollment: _base.entryDate,
  ename: _base.ename,
  cname: _base.cname,
  hkId: _base.hkId,
  ..._className,
  guide: _guide,
  campus_id: _campus,
  diploma_id: _diploma.id,
  diploma_name: _diploma.name
}

const namelessProgrammeClassSchema = {
  programmeClass_generic_name: _generic.name,
  programmeClass_generic_email: _generic.email,
  programmeClass_trade_name: _trade.name,
  programmeClass_trade_email: _trade.email
}
const programmeClassSchema = {
  programmeClass_name: _programmeClass.name,
  ...namelessProgrammeClassSchema
}

const allExportSchema = {
  ..._id,
  vdpId: _vdpId,
  ...baseExportSchema,
  ...programmeClassSchema,
  awardYear: _award_year,
  remark: _remark,
  dereg: _dereg,
  master_remark: _master_remark
}

const debugSuccessSchema = {
  ..._system,
  ...allExportSchema,
  ..._oriProgrammeClass,
  ..._warning,
  ..._original
}

const debugNoDupeSchema = {
  ..._system,
  ...allExportSchema,
  ..._oriProgrammeClass,
  ..._warning,
  ..._original
}
delete debugNoDupeSchema.awardYear

const debugAwardSchema = {
  ...debugNoDupeSchema
}
delete debugAwardSchema.dereg

const debugFailSchema = {
  ..._system,
  // ..._fragileId,
  ...allExportSchema,
  ..._oriProgrammeClass,
  ..._error,
  ..._original
}
delete debugFailSchema.awardYear

export const exportSchema = {
  all: worksheetColumns({
    vdpId: _vdpId,
    ...baseExportSchema,
    awardYear: _award_year,
    dereg: _dereg
  }),
  debugAwardSchema: worksheetColumns(debugAwardSchema),
  debugNoDupeSchema: worksheetColumns(debugNoDupeSchema),
  debugSuccessSchema: worksheetColumns(debugSuccessSchema),
  debugFailSchema: worksheetColumns(debugFailSchema),
  contactsc: worksheetColumns({ ...baseExportSchema, ...programmeClassSchema })
}
