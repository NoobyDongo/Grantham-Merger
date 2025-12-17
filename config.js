import fs from "fs"
import path from "path"
import { orange, green } from "./colors.js"
import { baseDir } from "./base-dir.js"

const __config_file_name = "config.json"

const numberParser = (value) => (value ? parseInt(value) : 0)

const configs = {
  grantham: {
    key: "篩選標準",
    content: {
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
  },
  logging: {
    key: "日誌記錄",
    content: {
      individualSummery: {
        key: "對每個輸入的 Excel 檔案提供總結 (true/false)",
        default: false,
      },
    },
  },
  output: {
    key: "I/O",
    content: {
      copyToBackup: {
        key: "將輸入資料複製到輸出資料夾 (true/false)",
        default: true,
      },
      removeInput: {
        key: "每次完成執行時刪除輸入數據 (true/false)",
        default: false,
      },
    },
  },
  outputDisplay: {
    key: "輸出顯示",
    content: {
      displayOutputFolder: {
        key: "每次完成執行時打開輸出資料夾 (true/false)",
        default: true,
      },
      displayOutputExcel: {
        key: "每次完成執行時打開輸出 Excel (true/false)",
        default: false,
      },
    },
  },
  teacherExcel: {
    key: "給 Dve 教師 Excel 的選項",
    content: {
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
      config: {
        key: "系統配置",
        content: {
          password: {
            key: "Excel 密碼",
            default: "12345678",
          },
          coveredByCampus: {
            key: "將生成的 Excel 用所屬的園校分類 (true/false)",
            default: true,
          },
        },
      },
    },
  },
  // _mailTest: {
  //   key: "Params for Email Service",
  //   content: {
  //     config: {
  //       key: "Config",
  //       content: {
  //         service: {
  //           key: "Mailing service Host",
  //           default: "smtp.office365.com",
  //         },
  //         port: {
  //           key: "Port",
  //           default: 587,
  //           parser: numberParser,
  //         },
  //       },
  //     },
  //     password: {
  //       key: "App password",
  //       default: "12345678",
  //       censored: true,
  //     },
  //     from: {
  //       key: "User",
  //       default: "xxx@xxx.com",
  //     },
  //     to: {
  //       key: "Mail Target",
  //       default: "xxx@xxx.com",
  //     },
  //     contactPhone: {
  //       key: "Contact(phone)",
  //       default: "1234-5678",
  //     },
  //     contactName: {
  //       key: "Name for contact(eg: 與XXX聯絡)",
  //       default: "X姑娘",
  //     },
  //     sender: {
  //       key: "Sender",
  //       default: "Andy Chan",
  //     },
  //     post: {
  //       key: "Post(title)",
  //       default: "SC",
  //     },
  //   },
  // },
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
  //so sad
  parseAfterwards = [],
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
    if (option.content) acc[curr] = generateDefaultConfig(option.content)
    else acc[curr] = option.default

    if (option.parser)
      parseAfterwards.push(() => {
        acc[option.key] = option.parser(acc[option.key])
      })

    return acc
  }, {})

const logConfigs = (config, _config, key) => {
  const withContent = {}
  const flat = Object.keys(_config)
    .filter((k) => {
      if (!config[k]) return false
      if (typeof _config[k] === "object") {
        if (_config[k] instanceof Date) {
          return true
        }
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
      acc[oriKey] = config[key].censored
        ? _config[key].replace(/./g, "*")
        : _config[key]

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

  console.log(
    `${green(`成功加載設定`)}[${green(__config_file_name)}], 已套用以下選項:\n`
  )
} catch (e) {
  console.log(
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

    const oriKey = config[key].key

    acc[oriKey] =
      typeof _config[key] === "object" && config[key].content
        ? syncKeyName(config[key].content, _config[key])
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

export const config = _config
