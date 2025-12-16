import path from "path"
import { fileURLToPath } from "url"

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export const isPkg = typeof process.pkg !== "undefined"
export const baseDir = isPkg ? path.dirname(process.execPath) : __dirname
