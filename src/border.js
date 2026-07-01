import tinycolor from 'tinycolor2'
import { getSchemeColorFromTheme } from './schemeColor'
import { getTextByPathList } from './utils'
import { getGradientFill, dealGradientFill, getSolidFill } from './fill'

export function getBorder(node, elType, warpObj) {
  let lineNode = getTextByPathList(node, ['p:spPr', 'a:ln']) || getTextByPathList(node, ['a:ln'])
  if (!lineNode) {
    const lnRefNode = getTextByPathList(node, ['p:style', 'a:lnRef'])
    if (lnRefNode) {
      const lnIdx = getTextByPathList(lnRefNode, ['attrs', 'idx'])
      lineNode = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:lnStyleLst']['a:ln'][Number(lnIdx) - 1]
    }
  }
  if (!lineNode) lineNode = node

  const isNoFill = getTextByPathList(lineNode, ['a:noFill'])

  let borderWidth = isNoFill ? 0 : (parseInt(getTextByPathList(lineNode, ['attrs', 'w'])) / 12700)
  if (isNaN(borderWidth)) {
    if (getTextByPathList(lineNode, ['a:solidFill'])) borderWidth = 2
    else if (lineNode) borderWidth = 0
    else if (elType !== 'obj') borderWidth = 0
    else borderWidth = 1
  }
  if (elType === 'text' && node['p:txBody']) {
    const textNode = node['p:txBody']
    const textNodeARList = getTextByPathList(textNode, ['a:p', 'a:r'])
    const samplingItem = Array.isArray(textNodeARList) && textNodeARList.length > 0 ? textNodeARList[0] : textNodeARList
    const flag = getTextByPathList(samplingItem, ['a:rPr', 'a:ln', 'attrs', 'w']) === '0'
    if (flag) {
      borderWidth = 0
    }
  }
  let borderColorObj
  if (getTextByPathList(lineNode, ['a:gradFill'])) {
    const gradFillObj = lineNode['a:gradFill'] 
    const gradientFillInfo = getGradientFill(gradFillObj, warpObj)
    borderColorObj = {
      type: 'gradient',
      value: gradientFillInfo,
    }
    dealGradientFill(borderColorObj)
  }


  let borderColor = getTextByPathList(lineNode, ['a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  if (!borderColor) {
    const schemeClrNode = getTextByPathList(lineNode, ['a:solidFill', 'a:schemeClr'])
    const schemeClr = 'a:' + getTextByPathList(schemeClrNode, ['attrs', 'val'])
    borderColor = getSchemeColorFromTheme(schemeClr, warpObj)
  }
  // 如果有 a:solidFill 是用getSolidFill获得正确的包含透明度、对比度等信息的颜色
  const solidFillNode = getTextByPathList(lineNode, ['a:solidFill'])
  if (solidFillNode) {
    const rightColor = getSolidFill(solidFillNode, warpObj)
    if (rightColor) borderColor = rightColor.replace('#', '')
  }

  if (!borderColor) {
    const schemeClrNode = getTextByPathList(node, ['p:style', 'a:lnRef', 'a:schemeClr'])
    const schemeClr = 'a:' + getTextByPathList(schemeClrNode, ['attrs', 'val'])
    borderColor = getSchemeColorFromTheme(schemeClr, warpObj)

    if (borderColor) {
      let shade = getTextByPathList(schemeClrNode, ['a:shade', 'attrs', 'val'])

      if (shade) {
        shade = parseInt(shade) / 100000
        
        const color = tinycolor('#' + borderColor).toHsl()
        borderColor = tinycolor({ h: color.h, s: color.s, l: color.l * shade, a: color.a }).toHex()
      }
    }
  }

  if (!borderColor) borderColor = '#000000'
  else borderColor = `#${borderColor}`

  const type = getTextByPathList(lineNode, ['a:prstDash', 'attrs', 'val'])
  let borderType = 'solid'
  let strokeDasharray = '0'
  switch (type) {
    case 'solid':
      borderType = 'solid'
      strokeDasharray = '0'
      break
    case 'dash':
      borderType = 'dashed'
      strokeDasharray = '5'
      break
    case 'dashDot':
      borderType = 'dashed'
      strokeDasharray = '5, 5, 1, 5'
      break
    case 'dot':
      borderType = 'dotted'
      strokeDasharray = '1, 5'
      break
    case 'lgDash':
      borderType = 'dashed'
      strokeDasharray = '10, 5'
      break
    case 'lgDashDotDot':
      borderType = 'dotted'
      strokeDasharray = '10, 5, 1, 5, 1, 5'
      break
    case 'sysDash':
      borderType = 'dashed'
      strokeDasharray = '5, 2'
      break
    case 'sysDashDot':
      borderType = 'dotted'
      strokeDasharray = '5, 2, 1, 5'
      break
    case 'sysDashDotDot':
      borderType = 'dotted'
      strokeDasharray = '5, 2, 1, 5, 1, 5'
      break
    case 'sysDot':
      borderType = 'dotted'
      strokeDasharray = '2, 5'
      break
    default:
  }

  return {
    borderColor,
    borderColorObj,
    borderWidth,
    borderType,
    strokeDasharray,
  }
}

