import {
  Worker,
  isMainThread,
  parentPort,
  workerData,
} from "node:worker_threads"
import { fileURLToPath } from "node:url"
import xlsxPopulate from "xlsx-populate"

const __filename = fileURLToPath(import.meta.url)

async function runWorker() {
  try {
    const { filePath, password } = workerData

    const wb = await xlsxPopulate.fromFileAsync(filePath)
    await wb.toFileAsync(filePath, { password })

    parentPort.postMessage("done")
  } catch (err) {
    parentPort.postMessage({ error: err.message })
  }
}

export function protectExcelFile(filePath, password) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(__filename, {
      workerData: { filePath, password },
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
