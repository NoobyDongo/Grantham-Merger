import fs from "fs"
import path from "path"
import { orange, green } from "./colors.js"
import { baseDir } from "./base-dir.js"

const __config_file_name = "config.json"

const numberParser = (value) => (value ? parseInt(value) : 0)

const configs = {
  grantham: {
    __key: "篩選標準",
    startYear: {
      key: "葛量洪 DVE Enrty 底線（年）",
      default: new Date().getFullYear() - 4,
      parser: numberParser,
    },
    startMonth: {
      key: "葛量洪 DVE Enrty 底線（月）",
      default: 6,
      parser: numberParser,
    },
  },
  logging: {
    __key: "日誌記錄",
    individualSummery: {
      key: "對每個輸入的 Excel 檔案提供總結 (true/false)",
      default: false,
    },
  },
  io: {
    __key: "I/O",
    copyToBackup: {
      key: "將輸入資料複製到輸出資料夾 (true/false)",
      default: true,
    },
    removeInput: {
      key: "每次完成執行時刪除輸入數據 (true/false)",
      default: false,
    },
    passwordIn: {
      key: "輸入的 Excel 的密碼（如有）",
      default: "",
    },
    passwordOut: {
      key: "輸出的 Excel 的密碼（如有）",
      default: "",
    },
    removeInputPassword: {
      key: "刪除輸入的 Excel 的密碼 (true/false)",
      default: true,
    },
  },
  outputDisplay: {
    __key: "輸出顯示",
    displayOutputFolder: {
      key: "每次完成執行時打開輸出資料夾 (true/false)",
      default: true,
    },
    displayOutputExcel: {
      key: "每次完成執行時打開輸出 Excel (true/false)",
      default: false,
    },
  },
  teacherExcel: {
    __key: "Dve Teacher Excel",
    address: {
      key: "VDPO 中文地址",
      default: "",
    },
    phone: {
      key: "VDPO 聯絡方式（電話）",
      default: "0000 0000",
    },
    fax: {
      key: "VDPO 聯絡方式（傳真）",
      default: "0000 0000",
    },
    deadline: {
      key: "Excel 截止收集日期 (yyyy/mm/dd)",
      default: new Date().toLocaleDateString("en-CA").replace(/-/g, "/"),
      parser: (d) => new Date(d),
    },
    coveredByCampus: {
      key: "將生成的 Excel 用所屬的園校分類 (true/false)",
      default: true,
    },
  },
  format: {
    __key: "排版",
    rowHeight: {
      __key: "行高",
      ycCampus: {
        key: "YC Campus Excel",
        default: 15,
        parser: numberParser,
      },
      others: {
        key: "除 YC Campus Excel 外的 Excel 行高",
        default: 60,
        parser: numberParser,
      },
      header: {
        key: "除 YC Campus Excel 外的 Excel Header 行高",
        default: 60,
        parser: numberParser,
      },
    },
    centered: {
      __key: "Excel 內容置中",
      horizontal: {
        key: "水平 (true/false)",
        default: true,
      },
      vertical: {
        key: "垂直 (true/false)",
        default: true,
      },
    },
  },
}

const loadConfig = () => {
  const reverseKeyedConfigs = ((_obj) => {
    const reverseKey = (obj) =>
      Object.keys(obj).reduce((acc, key) => {
        const chiKey = obj[key].key ?? obj[key].__key

        if (!chiKey) acc.__key = obj[key]
        else if (obj[key].__key)
          acc[chiKey] = reverseKey({ ...obj[key], __key: key })
        else acc[chiKey] = { ...obj[key], key: key }

        return acc
      }, {})
    return reverseKey(_obj)
  })(configs)

  var _configFile,
    //for vs code suggestions lol
    _config = { ...configs },
    //so sad
    parseAfterwards = [],
    usedDefault = false

  const parseOptions = (configs, json) =>
    Object.keys(configs).reduce((acc, key) => {
      const option = configs[key]

      if (option.__key) {
        acc[option.__key] = parseOptions(
          option,
          typeof json[key] === "object" ? json[key] : {}
        )
      } else if (option.key) {
        if (json[key] == undefined) {
          acc[option.key] = option.default

          usedDefault = true
        } else acc[option.key] = json[key]

        if (option.parser)
          parseAfterwards.push(() => {
            acc[option.key] = option.parser(acc[option.key])
          })
      }

      return acc
    }, {})

  const generateDefaultConfig = (configs) =>
    Object.keys(configs).reduce((acc, curr) => {
      const option = configs[curr]
      if (option.__key) acc[curr] = generateDefaultConfig(option)
      if (!option.key) return acc
      else acc[curr] = option.default

      if (option.parser)
        parseAfterwards.push(() => {
          acc[option.key] = option.parser(acc[option.key])
        })

      return acc
    }, {})

  const logConfigs = (config, _config, baseKey) => {
    const withContent = {}
    const flat = Object.keys(_config)
      .filter((k) => {
        if (!config[k]) return false
        if (typeof _config[k] === "object") {
          if (_config[k] instanceof Date) {
            return true
          }
          const key = config[k].key || config[k].__key
          withContent[baseKey ? `${baseKey}/${key}` : `${key}`] = {
            config: config[k],
            _config: _config[k],
          }
          return false
        }
        return true
      })
      .reduce((acc, key) => {
        const oriKey = config[key].key
        acc[oriKey] = config[key].censored
          ? _config[key].replace(/./g, "*")
          : _config[key]

        return acc
      }, {})

    if (baseKey) console.log(`[${green(baseKey)}]:`)
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

    console.group(
      `${green(`成功加載設定`)}[${green(
        __config_file_name
      )}], 已套用以下選項:\n`
    )
  } catch (e) {
    console.group(
      `${orange("無法加載設定")}[${orange(
        __config_file_name
      )}], 已套用預設值並建立新的設定文件\n`
    )
    _config = generateDefaultConfig(configs)
    usedDefault = true
  }

  const syncKeyName = (config, _config) =>
    Object.keys(_config).reduce((acc, key) => {
      if (!config[key]) return acc

      const oriKey = config[key].key ?? config[key].__key

      acc[oriKey] = config[key].__key
        ? syncKeyName(config[key], _config[key])
        : _config[key]

      return acc
    }, {})

  logConfigs(configs, _config)

  if (usedDefault) {
    fs.writeFileSync(
      path.join(baseDir, __config_file_name),
      JSON.stringify(syncKeyName(configs, _config), null, "\t")
    )
  }

  for (const f of parseAfterwards) {
    f()
  }

  console.groupEnd()

  return _config
}

export const config = loadConfig()

export const reloadConfig = () => Object.assign(config, loadConfig())
