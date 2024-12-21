import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

export default {
  entry: './main.js',
  target: 'node',
  mode: 'production',
  output: {
    path: path.resolve(__dirname, 'dist', 'bundle'),
    filename: 'bundle.js'
  },
}
