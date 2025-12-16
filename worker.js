import path from "node:path"
import {
  Worker,
  isMainThread,
  parentPort,
  workerData,
} from "node:worker_threads"
import xlsxPopulate from "xlsx-populate"
import { fileURLToPath } from "node:url"

const workerPath =
  typeof __dirname === "undefined" && typeof process.pkg == "undefined"
    ? fileURLToPath(import.meta.url)
    : path.join(__dirname, path.basename(fileURLToPath(import.meta.url)))

async function runWorker() {
  try {
    const { file, password } = workerData

    const wb = await xlsxPopulate.fromDataAsync(file)
    const buffer = await wb.outputAsync({ password })

    parentPort.postMessage(buffer)
  } catch (err) {
    parentPort.postMessage({ error: err.message })
  }
}

export function protectExcelFile(file, password) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(workerPath, {
      workerData: { file, password },
    })

    worker.on("message", (msg) => {
      if (msg && msg.error) {
        reject(new Error(msg.error))
      } else {
        resolve(msg)
      }
    })

    worker.on("error", reject)

    worker.on("exit", (code) => {
      if (code !== 0) {
        reject(new Error(`Worker stopped with exit code ${code}`))
      }
    })
  })
}

if (!isMainThread) {
  runWorker()
}
