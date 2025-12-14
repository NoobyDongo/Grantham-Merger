const now = new Date()
const year = now.getFullYear()
const month = String(now.getMonth() + 1).padStart(2, "0")
const day = String(now.getDate()).padStart(2, "0")
const hours = String(now.getHours()).padStart(2, "0")
const minutes = String(now.getMinutes()).padStart(2, "0")
const seconds = String(now.getSeconds()).padStart(2, "0")
const dateString = `${year}${month}${day}_${hours}${minutes}${seconds}`
const packageName = "temp"
const outputFilename = `${packageName}-${dateString}.exe`

import { execSync } from "child_process"

try {
  console.log("Bundling... \n")
  execSync(`npx webpack --config webpack.config.js --progress`, {
    stdio: "inherit",
    encoding: "utf8",
  })
  console.log("\nBuilding... \n")
  execSync(
    `pkg ./dist/bundle/bundle.js --target latest-win-x64 --output dist/${outputFilename}`,
    {
      stdio: "inherit",
      encoding: "utf8",
    }
  )
  console.log("\nDone. \n")
} catch (e) {
  console.error(e)
}
