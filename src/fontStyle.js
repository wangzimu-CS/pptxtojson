import { getTextByPathList } from './utils'
import { getShadow } from './shadow'
import { getFillType, getGradientFill, getSolidFill } from './fill'
import { RATIO_EMUs_Points } from './constants'

export function getFontType(node, type, warpObj, slideLayoutSpNode, slideMasterSpNode, slideMasterTextStyles) {
  const extractFont = (targetNode, isDirectRun = false) => {
    if (!targetNode) return null
    
    let rPr
    if (isDirectRun) rPr = getTextByPathList(targetNode, ['a:rPr']) 
    else {
      rPr = getTextByPathList(targetNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr'])
      if (!rPr) rPr = getTextByPathList(targetNode, ['p:txBody', 'a:p', 'a:pPr', 'a:defRPr'])
    }

    if (!rPr) return null

    return getTextByPathList(rPr, ['a:latin', 'attrs', 'typeface']) || getTextByPathList(rPr, ['a:ea', 'attrs', 'typeface'])
  }

  let typeface = extractFont(node, true)

  if (!typeface) typeface = extractFont(slideLayoutSpNode)
  if (!typeface) typeface = extractFont(slideMasterSpNode)

  if (!typeface) {
    let stylePath = []
    if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
      stylePath = ['p:titleStyle', 'a:lvl1pPr', 'a:defRPr']
    } 
    else if (type === 'body') {
      stylePath = ['p:bodyStyle', 'a:lvl1pPr', 'a:defRPr']
    } 
    else {
      stylePath = ['p:otherStyle', 'a:lvl1pPr', 'a:defRPr']
    }
    const masterGlobalRPr = getTextByPathList(slideMasterTextStyles, stylePath)
    if (masterGlobalRPr) {
      typeface = getTextByPathList(masterGlobalRPr, ['a:latin', 'attrs', 'typeface']) || getTextByPathList(masterGlobalRPr, ['a:ea', 'attrs', 'typeface'])
    }
  }

  if (!typeface || typeface.startsWith('+')) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

    if (fontSchemeNode) {
      if (typeface && typeface.startsWith('+')) {
        switch (typeface) {
          case '+mj-lt': 
            return getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
          case '+mn-lt': 
            const targetText = getTextByPathList(node, ['a:t'])
            if (typeof targetText === 'string' && targetText.length === 1) {
              const charScript = checkStringScripts(targetText)
              const targetTextScript = charScript && charScript.length >= 1 && charScript[0] ? charScript[0].script || '' : ''
              const targetTextTypefaceList = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:font'])

              const item = targetTextTypefaceList && targetTextTypefaceList.length
                ? targetTextTypefaceList.find(obj => obj && obj.attrs && obj.attrs.script === targetTextScript)
                : null

              const targetTextTypeface = item && item.attrs && item.attrs.typeface
                ? item.attrs.typeface
                : getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface']) || ''

              return targetTextTypeface
            }
            return getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
          case '+mj-ea': 
            return getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:ea', 'attrs', 'typeface'])
          case '+mn-ea': 
            return getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:ea', 'attrs', 'typeface'])
          default: 
            return typeface.replace(/^\+/, '')
        }
      }
    }

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      typeface = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface']) || getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:ea', 'attrs', 'typeface'])
    }
    else {
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    }
  }

  return typeface || ''
}

/**
 * 判断单个字符属于哪种文字脚本（Script）
 * @param {string} char - 单个字符
 * @returns {string|null} 脚本代码（如 Hans/Hant/Jpan...），无法识别返回 null
 */
function getCharScript(char) {
  if (!char || char.length !== 1) return null
  // ==================== 中文标点（必须单独判断）====================
  const isChinesePunctuation = /[？！。，；：‘’“”【】（）——…]/u.test(char)
  if (isChinesePunctuation) return 'Hans'

  const regexMap = {
    Hans: /\p{Script=Han}/u,
    Jpan: /\p{Script=Hiragana}|\p{Script=Katakana}/u,
    Hang: /\p{Script=Hangul}/u,
    Arab: /\p{Script=Arabic}/u,
    Hebr: /\p{Script=Hebrew}/u,
    Thai: /\p{Script=Thai}/u,
    Ethi: /\p{Script=Ethiopic}/u,
    Beng: /\p{Script=Bengali}/u,
    Gujr: /\p{Script=Gujarati}/u,
    Khmr: /\p{Script=Khmer}/u,
    Knda: /\p{Script=Kannada}/u,
    Guru: /\p{Script=Gurmukhi}/u,
    Cans: /\p{Script=Canadian_Aboriginal}/u,
    Cher: /\p{Script=Cherokee}/u,
    Yiii: /\p{Script=Yi}/u,
    Tibt: /\p{Script=Tibetan}/u,
    Thaa: /\p{Script=Thaana}/u,
    Deva: /\p{Script=Devanagari}/u,
    Telu: /\p{Script=Telugu}/u,
    Taml: /\p{Script=Tamil}/u,
    Syrc: /\p{Script=Syriac}/u,
    Orya: /\p{Script=Oriya}/u,
    Mlym: /\p{Script=Malayalam}/u,
    Laoo: /\p{Script=Lao}/u,
    Sinh: /\p{Script=Sinhala}/u,
    Mong: /\p{Script=Mongolian}/u,
    Viet: /[àáạảãâầấậẩẫêềếệểễòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/iu,
    Uigh: /\p{Script=Arabic}/u,
  }

  // 越南文
  if (regexMap.Viet.test(char)) return 'Viet'
  // 韩文
  if (regexMap.Hang.test(char)) return 'Hang'
  // 日文
  if (regexMap.Jpan.test(char)) return 'Jpan'
  // 汉字
  if (regexMap.Hans.test(char)) {
    try {
      const seg = new Intl.Segmenter('zh-TW', { granularity: 'grapheme' }).segment(char).next().value
      return seg.isTraditional ? 'Hant' : 'Hans'
    }
    catch (e) {
      return 'Hans'
    }
  }
  // 维吾尔文
  if (regexMap.Uigh.test(char)) return 'Uigh'
  // 其他语言
  for (const [script, regex] of Object.entries(regexMap)) {
    if (['Hans', 'Hant', 'Jpan', 'Hang', 'Viet', 'Uigh'].includes(script)) continue
    if (regex.test(char)) return script
  }

  return null
}

/**
 * 批量判断字符串中每个字符的脚本类型
 * @param {string} str - 输入字符串
 * @returns {Array<{char: string, script: string|null}>}
 */
function checkStringScripts(str) {
  return Array.from(str).map(char => ({
    char,
    script: getCharScript(char)
  }))
}

export function getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj) {
  const rPrNode = getTextByPathList(node, ['a:rPr'])
  let filTyp, color
  if (rPrNode) {
    filTyp = getFillType(rPrNode)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = rPrNode['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
    if (filTyp === 'GRADIENT_FILL') {
      const gradientFillNode = rPrNode['a:gradFill']
      const gradient = getGradientFill(gradientFillNode, warpObj)
      return gradient
    }
  }
  if (!color && getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])) {
    const lstStyledefRPr = getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])
    filTyp = getFillType(lstStyledefRPr)
    if (filTyp === 'SOLID_FILL') {
      const solidFillNode = lstStyledefRPr['a:solidFill']
      color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    }
  }
  if (!color) {
    const sPstyle = getTextByPathList(pNode, ['p:style', 'a:fontRef'])
    if (sPstyle) color = getSolidFill(sPstyle, undefined, undefined, warpObj)
    if (!color && pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
  }
  return color || ''
}
export function getBackGroundColor(node, warpObj) {
  const rPrNode = getTextByPathList(node, ['a:rPr'])
  // let filTyp, color
  let color
  if (rPrNode) {
    const highlight = rPrNode['a:highlight']
    // console.log('(00)-pptxtojson-getBackGround:highlight:', highlight)
    if (highlight) {
      color = getSolidFill(highlight, undefined, undefined, warpObj)
    }
    // const gradFill = rPrNode['a:gradFill']
    // console.log('(00)-pptxtojson-getBackGround:gradFill:', gradFill)
    // if (gradFill) {
    //   // color = getSolidFill(highlight, undefined, undefined, warpObj)
    //   color = 'linear-gradient(to right, #ff6a6a, #6a6aff)'
    // }
  }
  // filTyp = getFillType(rPrNode)
  // if (filTyp === 'SOLID_FILL') {
  //   const solidFillNode = rPrNode['a:solidFill']
  //   color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
  // }
  //   if (filTyp === 'GRADIENT_FILL') {
  //     const gradientFillNode = rPrNode['a:gradFill']
  //     const gradient = getGradientFill(gradientFillNode, warpObj)
  //     return gradient
  //   }
  // }
  // if (!color && getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])) {
  //   const lstStyledefRPr = getTextByPathList(lstStyle, ['a:lvl' + lvl + 'pPr', 'a:defRPr'])
  //   filTyp = getFillType(lstStyledefRPr)
  //   if (filTyp === 'SOLID_FILL') {
  //     const solidFillNode = lstStyledefRPr['a:solidFill']
  //     color = getSolidFill(solidFillNode, undefined, undefined, warpObj)
  //   }
  // }
  // if (!color) {
  //   const sPstyle = getTextByPathList(pNode, ['p:style', 'a:fontRef'])
  //   if (sPstyle) color = getSolidFill(sPstyle, undefined, undefined, warpObj)
  //   if (!color && pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
  // }
  return color || ''
}

export function getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode) {
  let fontSize

  if (getTextByPathList(node, ['a:rPr', 'attrs', 'sz'])) fontSize = getTextByPathList(node, ['a:rPr', 'attrs', 'sz']) / 100

  if ((isNaN(fontSize) || !fontSize) && pNode) {
    if (getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz'])) {
      fontSize = getTextByPathList(pNode, ['a:endParaRPr', 'attrs', 'sz']) / 100
    }
  }

  if ((isNaN(fontSize) || !fontSize) && textBodyNode) {
    const lstStyle = getTextByPathList(textBodyNode, ['a:lstStyle'])
    if (lstStyle) {
      let lvl = 1
      if (pNode) {
        const lvlNode = getTextByPathList(pNode, ['a:pPr', 'attrs', 'lvl'])
        if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1
      }

      const sz = getTextByPathList(lstStyle, [`a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
      if (sz) fontSize = parseInt(sz) / 100
    }
  }

  if ((isNaN(fontSize) || !fontSize)) {
    const sz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    if (sz) fontSize = parseInt(sz) / 100
  }

  if ((isNaN(fontSize) || !fontSize) && slideLayoutSpNode) {
    let lvl = 1
    if (pNode) {
      const lvlNode = getTextByPathList(pNode, ['a:pPr', 'attrs', 'lvl'])
      if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1
    }
    const layoutSz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', `a:lvl${lvl}pPr`, 'a:defRPr', 'attrs', 'sz'])
    if (layoutSz) fontSize = parseInt(layoutSz) / 100
  }

  if ((isNaN(fontSize) || !fontSize) && pNode) {
    const paraSz = getTextByPathList(pNode, ['a:pPr', 'a:defRPr', 'attrs', 'sz'])
    if (paraSz) fontSize = parseInt(paraSz) / 100
  }

  if (isNaN(fontSize) || !fontSize) {
    let sz
    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'body') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'dt' || type === 'sldNum') {
      sz = '1200'
    } 
    else if (!type) {
      sz = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    }
    if (sz) fontSize = parseInt(sz) / 100
  }

  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (baseline && !isNaN(fontSize)) {
    // fontSize -= 10
  }

  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return fontSize + 'pt'
}

export function getFontBold(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'b']) === '1' ? 'bold' : ''
}

export function getFontItalic(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'i']) === '1' ? 'italic' : ''
}

export function getFontDecoration(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'u']) === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'strike']) === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node) {
  const spc = getTextByPathList(node, ['a:rPr', 'attrs', 'spc'])
  return spc ? (parseInt(spc) / 100 + 'pt') : ''
}

export function getFontSubscript(node) {
  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (!baseline) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, warpObj) {
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  if (txtShadow) {
    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) {
      const { h, v, blur, color } = shadow
      if (!isNaN(v) && !isNaN(h)) {
        return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
      }
    }
  }
  return ''
}

export function getFontOutLine(node, warpObj) {
  // const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:ln', 'a:solidFill'])
  const w = getTextByPathList(node, ['a:rPr', 'a:ln', 'attrs', 'w']) || (1 / RATIO_EMUs_Points)
  if (txtShadow) {
    // const shadow = getShadow(txtShadow, warpObj)
    if (w) {
      const color = getSolidFill(txtShadow, undefined, undefined, warpObj)
      // const { h, v, blur, color } = shadow
      const h = w * RATIO_EMUs_Points
      const v = w * RATIO_EMUs_Points
      const blur = w * RATIO_EMUs_Points
      if (!isNaN(v) && !isNaN(h)) {
        return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
      }
    }
  }
  return ''
}

// PPT 7 种列表标识符 → 自动生成对应标号
export function getPPMark(index, type) {
  // 序号从 1 开始
  const num = parseInt(index) || 1

  switch (type) {
    // 1. 阿拉伯数字 1. 2. 3.
    case 'arabicPeriod':
      return `${num}.`

    // 2. 带圈数字 ① ② ③
    case 'circleNumDbPlain':
      return `①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳`[num - 1] || `${num}圈`

    // 3. 大写罗马数字 I. II. III.
    case 'romanUcPeriod':
      return toRoman(num) + '.'

    // 4. 大写字母 A. B. C.
    case 'alphaUcPeriod':
      return toLetter(num, true) + '.'

    // 5. 小写字母带右括号 a) b) c)
    case 'alphaLcParenR':
      return toLetter(num, false) + ')'

    // 6. 小写字母 a. b. c.
    case 'alphaLcPeriod':
      return toLetter(num, false) + '.'

    // 7. 中文数字 一、二、三、
    case 'ea1JpnChsDbPeriod':
      return toChineseNum(num) + '、'

    default:
      return `${num}.`
  }
}

// 辅助：转大写/小写字母
function toLetter(num, upper = false) {
  const code = (num - 1) % 26
  const char = String.fromCharCode(97 + code)
  return upper ? char.toUpperCase() : char
}

// 辅助：转罗马数字
function toRoman(num) {
  const roman = [
    ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX'],
    ['', 'X', 'XX', 'XXX', 'XL', 'L', 'LX', 'LXX', 'LXXX', 'XC'],
    ['', 'C', 'CC', 'CCC', 'CD', 'D', 'DC', 'DCC', 'DCCC', 'CM'],
  ]
  const digits = num.toString().padStart(3, '0').split('')
  return (
    roman[2][digits[0]] + roman[1][digits[1]] + roman[0][digits[2]]
  )
}

// 辅助：转中文数字
function toChineseNum(num) {
  const ch = '零一二三四五六七八九十'
  if (num <= 10) return ch[num]
  if (num < 20) return '十' + ch[num - 10]
  return num
}