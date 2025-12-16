export const customColor = (colorCode) => (text) => {
  return `\x1b[38;5;${colorCode}m${text}\x1b[0m`
}

export const yellow = customColor(3)
export const green = customColor(2)
export const red = customColor(1)
export const orange = customColor(208)
export const lightBlue = customColor(12)
