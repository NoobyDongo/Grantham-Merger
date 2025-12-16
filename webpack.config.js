import path from "path"
import { fileURLToPath } from "url"
// import CopyPlugin from "copy-webpack-plugin"

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export default {
  entry: { main: "./main.js", worker: "./worker.js" },
  target: "node",
  mode: "production",
  output: {
    path: path.resolve(__dirname, "dist", "bundle"),
    filename: "[name].js",
    publicPath: "./",
    chunkFormat: "commonjs",
    library: {
      type: "commonjs",
    },
  },
  node: {
    __dirname: false,
    __filename: false,
  },
  // plugins: [
  //   new CopyPlugin({
  //     patterns: [{ from: "resources", to: "resources" }],
  //   }),
  // ],
}
