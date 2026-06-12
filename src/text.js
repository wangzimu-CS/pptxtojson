import { getHorizontalAlign, getParagraphSpacing } from './align'
import { getTextByPathList } from './utils'
import { getSolidFill, } from './fill'

import {
  getFontType,
  getFontColor,
  getBackGroundColor,
  getFontSize,
  getFontBold,
  getFontItalic,
  getFontDecoration,
  getFontDecorationLine,
  getFontSpace,
  getFontSubscript,
  getFontShadow,
  getFontOutLine, getPPMark
} from './fontStyle'

function parsePPTTextToLinesNew(shapeData, width) {
  // 基础校验
  if (!shapeData || typeof shapeData !== 'object') return []

  // ======================
  // 1. 读取文本框宽度（完全来自传入数据，不写死）
  // ======================
  const spPr = shapeData['p:spPr']
  if (!spPr) return []

  const xfrm = spPr['a:xfrm']
  if (!xfrm) return []

  const ext = xfrm['a:ext']
  if (!ext || !ext.attrs || ext.attrs.cx === null) return []

  const cx = Number(ext.attrs.cx)
  if (isNaN(cx) || cx <= 0) return []
  // const textBoxWidthPx = cx / 12700 // PPT官方公式，无写死
  const textBoxWidthPx = width // PPT官方公式，无写死


  // ======================
  // 2. 读取段落
  // ======================
  const txBody = shapeData['p:txBody']
  if (!txBody) return []

  let paragraphs = txBody['a:p']
  if (!paragraphs) return []
  paragraphs = Array.isArray(paragraphs) ? paragraphs : [paragraphs]

  const result = []

  for (const p of paragraphs) {
    if (!p) continue

    let runs = p['a:r']
    if (!runs) continue
    runs = Array.isArray(runs) ? runs : [runs]

    // 拼接完整文本
    let fullText = ''
    let realFontSize = 36 // 兜底，但优先从XML读取

    // ======================
    // 3. 从 XML 里读取真实字号 sz → 计算真实字符宽度（无写死！）
    // ======================
    for (const r of runs) {
      if (r && typeof r['a:t'] === 'string') {
        fullText += r['a:t']
      }
      // 读取真实字号 sz 3600 = 36pt
      if (r && r['a:rPr'] && r['a:rPr'].attrs && r['a:rPr'].attrs.sz) {
        const sz = Number(r['a:rPr'].attrs.sz)
        if (!isNaN(sz) && sz > 0) {
          realFontSize = sz / 100 // 3600 → 36pt
        }
      }
    }

    // 真正计算字符宽度：pt → px（标准公式，无写死）
    // const charWidth = realFontSize * 1.3333
    const charWidth = realFontSize
    const abs = false
    if (abs) {
      textBoxWidthPx, charWidth
    }

    // ======================
    // 4. 自动换行（完全动态计算）
    // ======================
    // const lines = wrapTextReal(fullText, textBoxWidthPx, charWidth)
    // const lines = wrapTextReal(fullText, newTextBoxWidthPx, newCharWidth)
    const lines = wrapTextProfessional(fullText, textBoxWidthPx - 20, charWidth)

    result.push({
      text: fullText,
      lines: lines,
    })
  }

  return result
}

const cacheBuAutoNumType = []
// 生成字符级样式数组
function getCharStyleList(fullText, runList) {
  const textLength = fullText.length
  const charStyleList = new Array(textLength)

  // 遍历每个文本片段区间，把样式填充到对应字符位置
  // for (const run of runList) {
  //   const { start, end } = run

  //   // 安全边界判断
  //   if (start < 0 || end > textLength || start >= end) continue

  //   // 从 start ~ end-1 每个字符都赋值为该段样式
  //   for (let i = start; i < end; i++) {
  //     charStyleList[i] = { ...run }
  //     // console.log('(00)-genTextBody-:new-切行结果-charStyleList:【charStyleList[i], run】:', charStyleList[i], run)
  //   }
  // }
  let curGetIndex = 0
  for (let i = 0;i < fullText.length;i++) {
    const curChar = fullText[i]
    curChar
    const curStyle = runList[curGetIndex]
    
    if (curStyle.start <= i && curStyle.end > i) {
      // charStyleList.push(runList[curGetIndex])
    }
    else if (curStyle.end === i) {
      curGetIndex++
    }
    else {
      //
    }
    charStyleList[i] = { char: `【${curChar}】`, ...runList[curGetIndex]}
  }

  return charStyleList
}

function parsePPTTextToLines(spNode, getStyleUseInfo, paragraphInfo, styleObj, paramsObj) {
  console.log('(00)-tableEl-text-[切行]-parsePPTTextToLines:', 'paramsObj')
  if (!spNode || typeof spNode !== 'object') return []
  const spPr = spNode['p:spPr'] || spNode['a:tcPr']
  // console.log('(00)-tableEl-text-[spPr]:', spPr)
  if (!spPr) return []

  if (!paramsObj || !paramsObj.height || !paramsObj.width) return []

  const {lineHeight, isInGroup} = paramsObj
  const lineHeightImport = lineHeight
  // const xfrm = getTextByPathList(spPr, ['a:xfrm'])
  // if (xfrm || (!width && !height)) return []

  // const xfrm = spPr['a:xfrm']
  // if (!xfrm) return []

  // const ext = xfrm['a:ext']
  // if (!ext || !ext.attrs || ext.attrs.cx === null) return []

  // const cx = Number(ext.attrs.cx)
  // if (isNaN(cx) || cx <= 0) return []
  const textBoxWidthPx = (paramsObj && paramsObj.isVertical ? paramsObj.height : paramsObj.width)

  const txBody = spNode['p:txBody'] || spNode['a:txBody']
  if (!txBody) return []

  if (!getStyleUseInfo) return []


  const {
    pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj
  } = getStyleUseInfo

  let paragraphs = txBody['a:p']
  if (!paragraphs) return []
  paragraphs = Array.isArray(paragraphs) ? paragraphs : [paragraphs]

  const result = []
  const isOnlyOneP = paragraphs.length === 1
  let buAutoNumTypeCache
  let buAutoNumTypeCnt = 0
  let listSpanEl
  for (const p of paragraphs) {
    const listSpanElObj = {}
    listSpanElObj.styleText = ''
    let itemColor
    const listType = getListType(p)
    // const listLevel = getListLevel(p)
    if (listType && listType === 'ul') {
      listSpanElObj.type = 'ul'
      const buFont = getTextByPathList(p, ['a:pPr', 'a:buFont'])
      let buFontType
      if (buFont) buFontType = getTextByPathList(p, ['a:pPr', 'a:buFont', 'attrs', 'typeface'])
      const slideMasterTextStyles = warpObj['slideMasterTextStyles']
      const fontType = getFontType(p, type, warpObj, slideLayoutSpNode, slideMasterSpNode, slideMasterTextStyles)
      fontType
      // const buchar = getTextByPathList(p, ['a:pPr', 'a:buChar'])
      const buClr = getTextByPathList(p, ['a:pPr', 'a:buClr'])
      let buClrValue
      if (buClr) buClrValue = getSolidFill(buClr)
      const buChar = getTextByPathList(p, ['a:pPr', 'a:buChar'])
      let buCharType = ''
      if (buChar) buCharType = getTextByPathList(p, ['a:pPr', 'a:buChar', 'attrs', 'char'])
      // const buAutoNum = getTextByPathList(p, ['a:pPr', 'a:buAutoNum'])
      listSpanEl = `<span style="color: ${buClrValue}; font-family: ${buFontType};">${buCharType}</span>`
      if (buClrValue) {
        listSpanElObj.color = buClrValue
        listSpanElObj.styleText += `color: ${buClrValue};`
      }
      if (buFontType) {
        listSpanElObj.fontFamily = buFontType
        listSpanElObj.styleText += `font-family: ${buFontType};`
      }
      if (buCharType) {
        listSpanElObj.text = buCharType
      }
      // listSpanEl = `<span style="${listSpanElObj.styleText}">${listSpanElObj.text}</span>`
    }
    if (listType && listType === 'ol') {
      listSpanElObj.type = 'ol'
      // const buchar = getTextByPathList(p, ['a:pPr', 'a:buChar'])
      // const buClr = getTextByPathList(p, ['a:pPr', 'a:buClr'])
      const buAutoNum = getTextByPathList(p, ['a:pPr', 'a:buAutoNum'])
      const buFont = getTextByPathList(p, ['a:pPr', 'a:buFont'])
      let buFontType
      if (buFont) buFontType = getTextByPathList(p, ['a:pPr', 'a:buFont', 'attrs', 'typeface'])

      const slideMasterTextStyles = warpObj['slideMasterTextStyles']
      const fontType = getFontType(p, type, warpObj, slideLayoutSpNode, slideMasterSpNode, slideMasterTextStyles)

      let buAutoNumType = ''
      if (buAutoNum) buAutoNumType = getTextByPathList(p, ['a:pPr', 'a:buAutoNum', 'attrs', 'type'])
      if (buAutoNumTypeCache === buAutoNumType) {
        buAutoNumTypeCnt++
      }
      else {
        buAutoNumTypeCnt = 0
        cacheBuAutoNumType.push(buAutoNumType)
      }
      const markText = getPPMark(buAutoNumTypeCnt, buAutoNumType)
      // listSpanEl = `<span style="color: ${buClrValue}; font-family: ${buFontType};">${buCharType}</span>`
      // listSpanEl = `<span style="font-family: ${fontType || buFontType};">${markText}</span>`
      if (fontType || buFontType) {
        listSpanElObj.fontFamily = fontType || buFontType
        listSpanElObj.styleText += `font-family: ${fontType || buFontType};`
      }
      if (markText) {
        listSpanElObj.text = markText
      }
      listSpanEl = `<span style="${listSpanElObj.styleText}">${listSpanElObj.text}</span>`
      buAutoNumTypeCache = buAutoNumType
    }
    if (!p) continue
    let runs = p['a:r']
    const endParaRPr = p['a:endParaRPr']
    if (!runs && endParaRPr) {
      let lineHeight = 1.2
      const sz = Number(endParaRPr.attrs.sz)
      if (!isNaN(sz) && sz > 0) lineHeight = lineHeight * (sz / 100)
      const useImportLineHeight = lineHeightImport + (lineHeightImport - 1) * 0.5
      const paragraphHtml = `<p style="line-height: ${lineHeight * useImportLineHeight}px;">` + '<br></p>'
      result.push({
        text: '',
        lines: [],
        lineSpans: '',
        paragraphHtml: paragraphHtml,
        paraStyle: ''
      })
    }
    if (!runs) continue
    runs = Array.isArray(runs) ? runs : [runs]

    let fullText = ''
    let realFontSize = 36
    const runList = []
    console.log('(00)-tableEl-text-[切行]-runs:', runs)
    for (const r of runs) {
      if (r && typeof r['a:t'] === 'string') {
        const text = r['a:t']
        const start = fullText.length
        fullText += text
        const end = fullText.length
        const rNodeItem = r
        const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
        const style = getRunStyle(r, realFontSize)
        let styleText
        let styleObj
        if (styleInfo) {
          styleText = styleInfo.styleText
          styleObj = styleInfo.styleObj
        }
        runList.push({ start, end, text, ...style, styleText, styleObj })
      }

      if (r && r['a:rPr'] && r['a:rPr'].attrs && r['a:rPr'].attrs.sz) {
        const sz = Number(r['a:rPr'].attrs.sz)
        if (!isNaN(sz) && sz > 0) realFontSize = sz / 100
      }
    }

    const charStyleList = getCharStyleList(fullText, runList)
    const charWidth = realFontSize
    let lines
    if (fullText.replaceAll(' ', '').length > 5 && fullText.includes(' ') && !isInGroup) {
      lines = wrapTextProfessional(fullText, textBoxWidthPx - 13, charWidth, charStyleList)
    }
    else if (isInGroup) {
      console.log('(00)-tableEl-text-[切行]-lines:', '111111111111111111111111111111')
      lines = wrapTextProfessional(fullText, textBoxWidthPx + 30, charWidth, charStyleList)
    }
    else if (paramsObj.isVertical) {
      lines = wrapTextProfessional(fullText, textBoxWidthPx - 10, charWidth, charStyleList)
    }
    else {
      console.log('(00)-tableEl-text-[切行]-lines:', '2222222222222222222222')
      // lines = wrapTextProfessional(fullText, textBoxWidthPx + 15, charWidth, charStyleList)
      lines = wrapTextProfessional(fullText, textBoxWidthPx - 10, charWidth, charStyleList)
    }
    // console.log('(00)-tableEl-text-[切行]-parsePPTTextToLines:-[textBoxWidthPx]:', 'textBoxWidthPx')
    // const linesNew = wrapTextProfessionalNew(fullText, textBoxWidthPx, charWidth, charStyleList)
    // console.log('(00)-tableEl-text-[切行]-[linesNew,textBoxWidthPx]:', linesNew, textBoxWidthPx)
    console.log('(00)-tableEl-text-[切行]-lines:', lines)


    // ===================== 正确生成 HTML =====================
    const lineSpans = []
    let currentPos = 0
    // const lineHightRadio = lines.length>=5
    let lineHightRadio = 1.2
    if (isOnlyOneP || lines.length > 1) {
      lineHightRadio = 1.3333
    }

    // 行高规则-默认是字号的1.2倍
    let lineHight = realFontSize * 1.2

    for (const line of lines) {
      let curStyleText = charStyleList[0].styleText
      let curSpan = `<span style="${curStyleText}">` 
      let styleLeftSymbol 
      for (let i = 0; i < line.length; i++) {

        const lineChar = line[i]
        lineChar
        let thisStyle = charStyleList[currentPos]
        if (thisStyle.styleObj.color) {
          itemColor = thisStyle.styleObj.color
        } 
        currentPos ++
        if (curStyleText === thisStyle.styleText) {

          curSpan += line[i]
        }
        else {
          // 右符号要跟左符号样式统一
          if (line[i] === '）') {
            thisStyle = styleLeftSymbol

          }
          curStyleText = thisStyle.styleText
          curSpan += `</span><span style="${curStyleText}">` + line[i]

        }
        if (line[i].includes('（') && !line[i].includes('）')) {
          styleLeftSymbol = thisStyle
        }
        // charIndex += 1
      }
      curSpan += '</span>'

      // 行 span 无任何样式！
      lineHight = realFontSize * lineHightRadio
      // lineSpans.push('<span style=" white-space: pre; line-height:36px ">' + curSpan + '</span><br>')
      // lineSpans.push(`<span style=" white-space: pre; line-height:${lineHight}px ">` + curSpan + '</span><br>')
      // lineSpans.push(`<span style=" white-space: pre; line-height:1.5; ">` + curSpan + '</span><br>')
      lineSpans.push(`<span style=" white-space: pre;">` + curSpan + '</span><br>')

    }

    // for (const line of lines) {
    //   // // console.log('(00)-genTextBody-:new-切行结果-line_in_lines:', line)
    //   const lineLen = line.length
    //   const lineStart = currentPos
    //   const lineEnd = currentPos + lineLen
    //   currentPos = lineEnd

    //   const parts = []
    //   for (const run of runList) {
    //     if (run.end <= lineStart || run.start >= lineEnd) continue
    //     const s = Math.max(run.start, lineStart) - run.start
    //     const e = Math.min(run.end, lineEnd) - run.start
    //     const txt = run.text.slice(s, e)

    //     let styleStr = ''
    //     for (const key in run.style) {
    //       styleStr += key + ':' + run.style[key] + ';'
    //     }
    //     if (run.styleText) {
    //       parts.push('<span style="' + run.styleText + '">' + txt + '</span>')
    //     }
    //     else {
    //       parts.push('<span style="' + styleStr + '">' + txt + '</span>')
    //     }
    //   }

    //   // 行 span 无任何样式！
    //   lineSpans.push('<span style=" white-space: pre ">' + parts.join('') + '</span>')
    // }

    const paraStyle = {}
    // const paragraphHtml = '<p style="">' + lineSpans.join('') + '</p>'
    // const linesString = ` <span style=" white-space: pre-wrap; line-height:${lineHight}px  ">` + lineSpans.join('') + '</span>'
    lineHight
    const useImportLineHeight = lineHeightImport + (lineHeightImport - 1) * 0.5
    const linesString = ` <span style=" white-space: pre-wrap ">` + lineSpans.join('') + '</span>'
    let paragraphHtml = '<p style="">' + linesString + '</p>'
    const useStyleText = styleObj.styleText || ''
    if (paragraphInfo && paragraphInfo.start && paragraphInfo.end) {
      paragraphHtml = paragraphInfo.start + linesString + paragraphInfo.end
      // paragraphHtml = `<p style="line-height:${lineHight}px">` + linesString + '</p>'
      // paragraphHtml = `<p style="${useStyleText} line-height: ${lineHight * useImportLineHeight}px;">` + linesString + '</p>'
      const lineHeightValue = isInGroup ? realFontSize * lineHeightImport : lineHight * useImportLineHeight
      // paragraphHtml = `<p style="${useStyleText} line-height: ${lineHight * useImportLineHeight}px;">` + linesString + '</p>'
      // paragraphHtml = `<p style="${useStyleText} line-height: ${realFontSize * lineHeightImport}px;">` + linesString + '</p>'
      let listSpanElText = listSpanEl || ''
      if (itemColor && listSpanElObj.text && listSpanElObj.styleText) {
        listSpanElObj.styleText += `color: ${itemColor}`
        listSpanElText = `<span style="${listSpanElObj.styleText}">${listSpanElObj.text}</span>`
      }
      paragraphHtml = `<p style="${useStyleText} line-height: ${lineHeightValue}px;">` + listSpanElText + linesString + '</p>'
      
    }
    result.push({
      text: fullText,
      lines: lines,
      lineSpans: lineSpans,
      paragraphHtml: paragraphHtml,
      paraStyle: paraStyle
    })
  }
  return result
}

// ===================== 独立抽取：获取单个 a:r 样式 =====================
function getRunStyle(r, defaultFontSize) {
  const style = {}
  let fontSize = defaultFontSize
  let bold = false
  let underline = false
  let fontFamily = ''

  if (r['a:rPr']) {
    const rPr = r['a:rPr']
    if (rPr.attrs && rPr.attrs.sz) {
      const sz = Number(rPr.attrs.sz)
      if (!isNaN(sz)) fontSize = sz / 100
    }
    if (rPr.attrs && rPr.attrs.b === '1') bold = true
    if (rPr.attrs && rPr.attrs.u === 'sng') underline = true
    if (rPr['a:ea'] && rPr['a:ea'].attrs && rPr['a:ea'].attrs.typeface) {
      fontFamily = rPr['a:ea'].attrs.typeface
    }
  }

  if (fontSize) style['font-size'] = fontSize + 'px'
  if (fontFamily) style['font-family'] = fontFamily
  if (bold) style['font-weight'] = 'bold'
  if (underline) style['text-decoration'] = 'underline'

  return { style }
}


/**
 * 获取单个字符在指定样式下的实际宽度（px）
 * @param {string} char - 单个字符（a/i/m/1/,...）
 * @param {string} fontSize - 字号，如 '16px'
 * @param {string} fontFamily - 字体，如 'Arial'
 * @param {string} fontWeight - 字重 默认 'normal'
 * @returns {number} 宽度 px
 */
function getCharWidth(
  char,
  fontSize = '16px',
  fontFamily = 'Arial',
  fontWeight = 'normal'
) {
  const span = document.createElement('span')
  // 必须白边、不换行、不影响布局
  span.style.visibility = 'hidden'
  span.style.position = 'absolute'
  span.style.whiteSpace = 'nowrap'
  span.style.fontSize = fontSize
  span.style.fontFamily = fontFamily
  span.style.fontWeight = fontWeight
  span.textContent = char

  document.body.appendChild(span)
  const width = span.offsetWidth
  document.body.removeChild(span)

  return width
}
// 高精度测量：用 getBoundingClientRect() 而不是 offsetWidth（支持小数）
function getPreciseWidth(char, fontSize = '30px', fontFamily = 'Arial', fontWeight = 'normal') {
  // const styleObj = {
  //   fontSize, 
  //   fontFamily,
  //   fontWeight,
  // }
  const span = document.createElement('span')
  span.style.visibility = 'hidden'
  span.style.position = 'absolute'
  span.style.whiteSpace = 'nowrap'
  span.style.fontSize = fontSize
  span.style.fontFamily = fontFamily
  span.style.fontWeight = fontWeight
  span.style.letterSpacing = 'normal'
  span.textContent = char
  document.body.appendChild(span)
  const rect = span.getBoundingClientRect()
  document.body.removeChild(span)
  return rect.width // 小数宽度！
  // return rect.height // 小数宽度！
}

/**
 * 判断单个字符是否为全角字符
 * @param {string} char - 单个字符
 * @returns {boolean} true=全角，false=半角
 */
function isFullWidthChar(char) {
  // 匹配：中文汉字 + 全角符号(全角英文/数字/标点)
  const fullWidthReg = /[\u4E00-\u9FFF\uFF00-\uFFEF]/
  // 必须是单个字符才判断
  return char.length === 1 && fullWidthReg.test(char)
}

/**
 * 真正正确的 PPT 中文自动换行
 * 汉字 = 全宽
 * 空格/标点 = 半宽
 * 严格按像素计算
 * 标点不出现在行首
 */

// 行首禁止出现的标点
const NO_LINE_START = new Set([
  '，', '。', '、', '；', '：', '）', '”', '！', '？', '…'
])

const spicalChar = new Set(['“', '”'])
// 空格、英文标点使用半宽
const HALF_WIDTH_CHARS = new Set([
  ' ', ' ', ' ', ' ', '\t',
  ',', '.', ';', ':', '!', '?', '"', ')', ']'
])

function wrapTextProfessional(text, maxLineWidthPx, fullCharWidth, charStyleList) {
  if (typeof text !== 'string' || text === '') return []
  if (maxLineWidthPx <= 0 || fullCharWidth <= 0) return []

  const halfCharWidth = fullCharWidth * 0.5
  const lines = []
  let currentLine = ''
  let currentWidth = 0

  let index = 0
  const useNew = false
  if (!useNew) {
    const currentWidthList = []
    for (const char of text) {
    // ✅ 关键修复：空格 != 汉字宽度
    // const charW = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      const charW1 = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      charW1
      // const charW = isFullWidthChar(char) && (!HALF_WIDTH_CHARS.has(char)) ? fullCharWidth : halfCharWidth
      const charIsFullWidth = isFullWidthChar(char)
      const charW = charIsFullWidth || char === '。' ? fullCharWidth : halfCharWidth

      let useCharW = charW
      if (charStyleList && charStyleList.length >= index) {

        // const curStyle = charStyleList[index].style
        // const fontSize = curStyle['font-size']
        // const fontFamily = curStyle['font-family']
        // const fontWeight = curStyle['font-weight']
        const curStyle = charStyleList[index].styleObj
        const {fontSize, fontFamily, fontWeight} = curStyle
      

        const realCharW = getCharWidth(char, fontSize, fontFamily, fontWeight) || charW
        const realCharW1 = getPreciseWidth(char, fontSize, fontFamily, fontWeight) || charW

        realCharW
        useCharW = spicalChar.has(char) ? fullCharWidth : realCharW1
      }

      // isFullWidthChar(char) && (!HALF_WIDTH_CHARS.has(char)) ? fullCharWidth : halfCharWidth
      HALF_WIDTH_CHARS.has(char)

      const useMaxLineWidthPx = maxLineWidthPx
      // 超宽判断
      const diffW = currentWidth + useCharW - useMaxLineWidthPx
      // const half = diffW / useCharW > 0.15 
      diffW
      const half = true
      if (currentWidth + useCharW > useMaxLineWidthPx && half) {
        // 标点不能放行首
        if (NO_LINE_START.has(char)) {
        // if (noLineEnter) {
          currentLine += char
          currentWidth += useCharW
          currentWidthList.push(useCharW)
          // const endWithLeftSymbol = checkLastIsLeftSymbol(currentLine)
          const endWithLeftSymbol1 = checkLastLeftSymbol(currentLine)
          let newStr = ''
          let newStrWidth = 0
          if (endWithLeftSymbol1.match) {
            removeRightChar('sadfas', 1)
            const rmRCharsResult = removeRightChars(currentLine, endWithLeftSymbol1.rightPosition)
            currentLine = rmRCharsResult.newStr
            newStr = rmRCharsResult.removedChars
            newStrWidth = sumRightNums(currentWidthList, rmRCharsResult.removedChars.length)
          }
          lines.push(currentLine)
          currentLine = newStr
          currentWidth = newStrWidth
        }
        else {
          // const endWithLeftSymbol = checkLastIsLeftSymbol(currentLine)
          const endWithLeftSymbol1 = checkLastLeftSymbol(currentLine)
          let newStr = ''
          let newStrWidth = 0
          if (endWithLeftSymbol1.match) {
            removeRightChar('sadfas', 1)
            const rmRCharsResult = removeRightChars(currentLine, endWithLeftSymbol1.rightPosition)
            currentLine = rmRCharsResult.newStr
            newStr += rmRCharsResult.removedChars
            newStrWidth = sumRightNums(currentWidthList, rmRCharsResult.removedChars.length - 1)
          }
          lines.push(currentLine)
          currentLine = newStr + char
          currentWidth = newStrWidth + useCharW
          currentWidthList.push(useCharW)
          // lines.push(currentLine)
          // currentLine = char
          // currentWidth = useCharW
        }
      }
      else {
        currentLine += char
        currentWidth += useCharW
        currentWidthList.push(useCharW)
      }

      index++
    }
  }
  else {
    for (let i = 0;i < text.length;i++) {
      const char = text[i]
      // for (const char of text) {
      // ✅ 关键修复：空格 != 汉字宽度
      const charW1 = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      charW1
      const charIsFullWidth = isFullWidthChar(char)
      const charW = charIsFullWidth || char === '。' ? fullCharWidth : halfCharWidth

      let useCharW = charW
      if (charStyleList && charStyleList.length >= index) {

        const curStyle = charStyleList[index].styleObj
        const {fontSize, fontFamily, fontWeight} = curStyle
      

        const realCharW = getCharWidth(char, fontSize, fontFamily, fontWeight) || charW
        const realCharW1 = getPreciseWidth(char, fontSize, fontFamily, fontWeight) || charW

        realCharW
        useCharW = spicalChar.has(char) ? fullCharWidth : realCharW1
      }


      HALF_WIDTH_CHARS.has(char)

      const useMaxLineWidthPx = maxLineWidthPx
      // 超宽判断
      if (currentWidth + useCharW > useMaxLineWidthPx) {
      // const noLineEnter = isRemainingALLEmpty(index, text)
        const remainInfo = isRemainingALLEmpty(i, text)

        // 标点不能放行首
        // if (NO_LINE_START.has(char)) {
        // if (noLineEnter) {
        //   currentLine += char
        //   currentWidth += useCharW
        //   lines.push(currentLine)
        //   currentLine = ''
        //   currentWidth = 0
        // }
        if (remainInfo.allInOneline) {
          currentLine += remainInfo.remainChars
          currentWidth += useCharW
          lines.push(currentLine)
          currentLine = ''
          currentWidth = 0
        }
        else {
          lines.push(currentLine)
          currentLine = char
          currentWidth = useCharW
        }
      }
      else {
        currentLine += char
        currentWidth += useCharW
      }

      index++
    }
  }

  let isEmptyLine = true
  for (const lineChar of currentLine) {
    if (lineChar !== ' ') {
      isEmptyLine = false
      break
    }
  }
  isEmptyLine
  if (currentLine !== '') {
    if (lines.length > 0 && isEmptyLine) {
      lines[lines.length - 1] += currentLine
    }
    else {
      lines.push(currentLine)
    }
    // lines.push(currentLine)
  }

  return lines
}
function wrapTextProfessionalNew(text, maxLineWidthPx, fullCharWidth, charStyleList) {
  if (typeof text !== 'string' || text === '') return []
  if (maxLineWidthPx <= 0 || fullCharWidth <= 0) return []

  const halfCharWidth = fullCharWidth * 0.5
  const lines = []
  let currentLine = ''
  let currentWidth = 0

  let index = 0
  const useNew = false
  if (!useNew) {
    const currentWidthList = []
    for (const char of text) {
    // ✅ 关键修复：空格 != 汉字宽度
    // const charW = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      const charW1 = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      charW1
      // const charW = isFullWidthChar(char) && (!HALF_WIDTH_CHARS.has(char)) ? fullCharWidth : halfCharWidth
      const charIsFullWidth = isFullWidthChar(char)
      const charW = charIsFullWidth || char === '。' ? fullCharWidth : halfCharWidth

      let useCharW = charW
      if (charStyleList && charStyleList.length >= index) {

        // const curStyle = charStyleList[index].style
        // const fontSize = curStyle['font-size']
        // const fontFamily = curStyle['font-family']
        // const fontWeight = curStyle['font-weight']
        const curStyle = charStyleList[index].styleObj
        const {fontSize, fontFamily, fontWeight} = curStyle
      

        const realCharW = getCharWidth(char, fontSize, fontFamily, fontWeight) || charW
        const realCharW1 = getPreciseWidth(char, fontSize, fontFamily, fontWeight) || charW

        realCharW
        useCharW = spicalChar.has(char) ? fullCharWidth : realCharW1
      }

      // isFullWidthChar(char) && (!HALF_WIDTH_CHARS.has(char)) ? fullCharWidth : halfCharWidth
      HALF_WIDTH_CHARS.has(char)

      const useMaxLineWidthPx = maxLineWidthPx
      // 超宽判断
      const diffW = currentWidth + useCharW - useMaxLineWidthPx
      // const half = diffW / useCharW > 0.15 
      diffW
      const half = true
      if (currentWidth + useCharW > useMaxLineWidthPx && half) {
        // 标点不能放行首
        if (NO_LINE_START.has(char)) {
        // if (noLineEnter) {
          currentLine += char
          currentWidth += useCharW
          currentWidthList.push(useCharW)
          // const endWithLeftSymbol = checkLastIsLeftSymbol(currentLine)
          const endWithLeftSymbol1 = checkLastLeftSymbol(currentLine)
          let newStr = ''
          let newStrWidth = 0
          if (endWithLeftSymbol1.match) {
            removeRightChar('sadfas', 1)
            const rmRCharsResult = removeRightChars(currentLine, endWithLeftSymbol1.rightPosition)
            currentLine = rmRCharsResult.newStr
            newStr = rmRCharsResult.removedChars
            newStrWidth = sumRightNums(currentWidthList, rmRCharsResult.removedChars.length)
          }
          lines.push(currentLine)
          currentLine = newStr
          currentWidth = newStrWidth
        }
        else {
          // const endWithLeftSymbol = checkLastIsLeftSymbol(currentLine)
          const endWithLeftSymbol1 = checkLastLeftSymbol(currentLine)
          let newStr = ''
          let newStrWidth = 0
          if (endWithLeftSymbol1.match) {
            removeRightChar('sadfas', 1)
            const rmRCharsResult = removeRightChars(currentLine, endWithLeftSymbol1.rightPosition)
            currentLine = rmRCharsResult.newStr
            newStr += rmRCharsResult.removedChars
            newStrWidth = sumRightNums(currentWidthList, rmRCharsResult.removedChars.length - 1)
          }
          lines.push(currentLine)
          currentLine = newStr + char
          currentWidth = newStrWidth + useCharW
          currentWidthList.push(useCharW)
          // lines.push(currentLine)
          // currentLine = char
          // currentWidth = useCharW
        }
      }
      else {
        currentLine += char
        currentWidth += useCharW
        currentWidthList.push(useCharW)
      }

      index++
    }
  }
  else {
    for (let i = 0;i < text.length;i++) {
      const char = text[i]
      // for (const char of text) {
      // ✅ 关键修复：空格 != 汉字宽度
      const charW1 = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth
      charW1
      const charIsFullWidth = isFullWidthChar(char)
      const charW = charIsFullWidth || char === '。' ? fullCharWidth : halfCharWidth

      let useCharW = charW
      if (charStyleList && charStyleList.length >= index) {

        const curStyle = charStyleList[index].styleObj
        const {fontSize, fontFamily, fontWeight} = curStyle
      

        const realCharW = getCharWidth(char, fontSize, fontFamily, fontWeight) || charW
        const realCharW1 = getPreciseWidth(char, fontSize, fontFamily, fontWeight) || charW

        realCharW
        useCharW = spicalChar.has(char) ? fullCharWidth : realCharW1
      }


      HALF_WIDTH_CHARS.has(char)

      const useMaxLineWidthPx = maxLineWidthPx
      // 超宽判断
      if (currentWidth + useCharW > useMaxLineWidthPx) {
      // const noLineEnter = isRemainingALLEmpty(index, text)
        const remainInfo = isRemainingALLEmpty(i, text)

        // 标点不能放行首
        // if (NO_LINE_START.has(char)) {
        // if (noLineEnter) {
        //   currentLine += char
        //   currentWidth += useCharW
        //   lines.push(currentLine)
        //   currentLine = ''
        //   currentWidth = 0
        // }
        if (remainInfo.allInOneline) {
          currentLine += remainInfo.remainChars
          currentWidth += useCharW
          lines.push({currentLine, currentWidth})
          currentLine = ''
          currentWidth = 0
        }
        else {
          lines.push({currentLine, currentWidth})
          currentLine = char
          currentWidth = useCharW
        }
      }
      else {
        currentLine += char
        currentWidth += useCharW
      }

      index++
    }
  }

  let isEmptyLine = true
  for (const lineChar of currentLine) {
    if (lineChar !== ' ') {
      isEmptyLine = false
      break
    }
  }
  isEmptyLine
  if (currentLine !== '') {
    if (lines.length > 0 && isEmptyLine) {
      lines[lines.length - 1] += currentLine
    }
    else {
      // lines.push(currentLine)
      lines.push({currentLine, currentWidth})
    }
    // lines.push(currentLine)
  }

  return lines
}

function sumRightNums(arr, n) {
  if (!Array.isArray(arr) || n <= 0) return 0
  const start = Math.max(0, arr.length - n)
  const rightItems = arr.slice(start)
  return rightItems.reduce((sum, num) => sum + Number(num) || 0, 0)
}

function removeRightChar(str, n) {
  if (n <= 0 || n > str.length) return str
  return str.slice(0, -n)
}
function removeRightChars(str, n) {
  const len = str.length
  
  if (n <= 0 || n > len) {
    return {
      newStr: str,
      removedChars: ''
    }
  }

  const removedChars = str.slice(-n)
  const newStr = str.slice(0, -n)
  return {
    newStr,
    removedChars
  }
}
/**
 * 判断字符串【最后一个非空格字符】是不是【各种左符号】
 * 包含：(  [  {  《  “  ‘
 */
// function checkLastIsLeftSymbol(str) {
//   // 1. 去掉末尾所有空格
//   const trimmed = str.trimEnd()

//   // 2. 空字符串直接返回 false
//   if (trimmed.length === 0) return false

//   // 3. 取最后一个字符
//   const last = trimmed.slice(-1)

//   // 4. 判断是不是【左符号】
//   const leftSymbols = ['(', '（', '[', '{', '《', '“', '‘', '《', '“', '‘', '〈', '『', '【']
//   return leftSymbols.includes(last)
// }

function checkLastLeftSymbol(str) {
  const len = str.length
  let pos = 0
  let char = ''

  for (let i = len - 1; i >= 0; i--) {
    pos++
    if (str[i] !== ' ') {
      char = str[i]
      break
    }
  }

  const symbols = ['(', '[', '{', '《', '〈', '『', '〖', '“', '‘', '【', '(', '（', '[', '{', '《', '“', '‘', '《', '“', '‘', '〈', '『', '【']
  const match = symbols.includes(char)
  
  return {
    match: match,
    rightPosition: pos,
    char: char
  }
}


function isRemainingALLEmpty(curIndex, charList ) {
  const remainInfo = {
    allInOneline: false,
    remainChars: '',
  }
  if (curIndex === charList.length - 1) {
    const char = charList[curIndex]
    remainInfo.allInOneline = NO_LINE_START.has(char) || char === ''
    remainInfo.remainChars = char
    return remainInfo
  }
  for (let i = curIndex;i < charList.length;i++) {
    const char = charList[i]
    if (char !== ' ') {
      remainInfo.allInOneline = true
      return remainInfo
    } 
    remainInfo.remainChars += char
    curIndex++
  }
  remainInfo.allInOneline = true
  return true
}

export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj, paramsObj) {
  if (!textBodyNode) return ''
  let text = ''

  const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])
  
  const pNode = textBodyNode['a:p']
  // const {width, height, useNewDeal} = paramsObj
  const pNodes = pNode.constructor === Array ? pNode : [pNode]
  
  const listTypes = []
  for (const pNode of pNodes) {
    let rNode = pNode['a:r']
    let fldNode = pNode['a:fld']
    let brNode = pNode['a:br']
    if (rNode) {
      rNode = (rNode.constructor === Array) ? rNode : [rNode]

      if (fldNode) {
        fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode]
        rNode = rNode.concat(fldNode)
      }
      if (brNode) {
        brNode = (brNode.constructor === Array) ? brNode : [brNode]
        brNode.forEach(item => item.type = 'br')
  
        if (brNode.length > 1) brNode.shift()
        rNode = rNode.concat(brNode)
        rNode.sort((a, b) => {
          if (!a.attrs || !b.attrs) return true
          return a.attrs.order - b.attrs.order
        })
      }
    }

    // 增加整个textbody对齐参信息的获取-解决文本框对齐问题
    const lstStyle = textBodyNode['a:lstStyle']
    const lstStyle_align = getTextByPathList(lstStyle, ['a:lvl1pPr', 'attrs', 'algn'])
    const align = getHorizontalAlign(pNode, spNode, type, warpObj, lstStyle_align)
    const spacing = getParagraphSpacing(pNode)
    let styleText = `text-align: ${align};`
    let lineHeight = 1
    const styleObj = {
      textAlign: align,
      styleText: styleText,
    }
    if (align === 'justify') {
      styleText += `text-align-last: ${align};`
      // styleText += `margin: 0;`
      // styleText += `padding: 0;`
      styleObj.textAlignLast = align
      // styleObj.margin = 0
      // styleObj.padding = 0
    }

    // const spPr = spNode['p:spPr'] || spNode['a:tcPr']
    // const attrsNode = getTextByPathList(spPr, ['attrs'])
    // // t、b、ctr
    // // const anchorCheckList = ['t', 'b', 'ctr']
    // const anchorDict = {'t': 'flex-start', 'b': 'flex-end', 'ctr': 'center'}
    // console.log('(00)-tableEl-text-[spPr]:', spPr, attrsNode, Object.keys(anchorDict))
    // let alignItems = anchorDict['t']
    // if (spPr && attrsNode) {
    //   const anchorInfo = getTextByPathList(attrsNode, ['anchor'])
    //   if (anchorInfo && Object.keys(anchorDict).includes(anchorInfo)) {
    //     alignItems = anchorDict[`${anchorInfo}`]
    //     styleText += `align-items: ${alignItems};`
    //     styleText += `display: flex;`
    //     styleObj.alignItems = alignItems
    //     styleObj.display = 'flex'
    //     styleText += `border: 1px solid red;`
    //     styleObj.border = ' 1px solid red'

    //     // styleText += `width: 100%;`
    //     // styleObj.width = ' 1px solid red'
    //     styleText += `height: 100%;`
    //     styleObj.height = '100%'
    //   }
    // }
    // console.log('(00)-tableEl-text-[spPr]:', spPr, attrsNode, alignItems, styleText, styleObj)


    if (spacing) {
      if (spacing.lineSpacing) {
        styleText += `line-height: ${spacing.lineSpacing};`
        lineHeight = spacing.lineSpacing
        styleObj.lineHeight = lineHeight
      }
      if (spacing.spaceBefore) {
        styleText += `margin-top: ${spacing.spaceBefore};`
        styleObj.marginTop = spacing.spaceBefore
        styleObj.styleText += `margin-top: ${spacing.spaceBefore};`
      }
      if (spacing.spaceAfter) {
        styleText += `margin-bottom: ${spacing.spaceAfter};`
        styleObj.marginBottom = spacing.spaceAfter
        styleObj.styleText += `margin-bottom: ${spacing.spaceAfter};`
      }
    }
    else {
      // styleText += `line-height: 1.2;`
      styleText += `line-height: 1;`
    }

    styleText += `justify-content: center;`
    styleObj.justifyContent = 'center'

    const listType = getListType(pNode)
    const listLevel = getListLevel(pNode)
    const paragraphInfo = {start: '', end: ''}
    if (listType) {
      while (listTypes.length > listLevel + 1) {
        const closedListType = listTypes.pop()
        text += `</${closedListType}>`
      }

      if (listTypes[listLevel] === undefined) {
        text += `<${listType}>`
        listTypes[listLevel] = listType
      }
      else if (listTypes[listLevel] !== listType) {
        text += `</${listTypes[listLevel]}>`
        text += `<${listType}>`
        listTypes[listLevel] = listType
      }
      text += `<li style="${styleText}">`
      paragraphInfo.start = `<li style="${styleText}">`
      paragraphInfo.end = `</li>`
    }
    else {
      while (listTypes.length > 0) {
        const closedListType = listTypes.pop()
        text += `</${closedListType}>`
      }
      text += `<p style="${styleText}">`
      paragraphInfo.start = `<p style="${styleText}">`
      paragraphInfo.end = `</p>`
    }
    
    if (!rNode) {
      text += genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
    } 
    else {
      let prevStyleInfo = null
      let accumulatedText = ''
      const isUseNewDeal = paramsObj && paramsObj.isUseNewDeal
      console.log('(00)-tableEl-text-[切行]-paramsObj:', paramsObj) 
      if (isUseNewDeal) {
        let newDeal = false
        if (!newDeal) {
          text = ''
          const getStyleInfo = {
            pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj
          }
          const cutLineResultNew = parsePPTTextToLinesNew(spNode, paramsObj.width)
          cutLineResultNew

          const isInGroup = warpObj.isInGroup
          const cutLineParamsObj = {
            ...paramsObj,
            isInGroup,
            lineHeight,
          }
          const cutLineResult = parsePPTTextToLines(spNode, getStyleInfo, paragraphInfo, styleObj, cutLineParamsObj)
          for (let i = 0; i < cutLineResult.length;i++ ) {
            text += cutLineResult[i].paragraphHtml
            newDeal = true
          }
        }

        if (newDeal) {
        // return text
          text = replaceNbspByLimit(text, 3)
          if (text.includes('&nbsp;')) {
            text = text.replace('&nbsp;', ' ')
          }
          return text
        }
      }
      else {
        for (const rNodeItem of rNode) {
          const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)

          if (!prevStyleInfo || prevStyleInfo.styleText !== styleInfo.styleText || prevStyleInfo.hasLink !== styleInfo.hasLink || styleInfo.hasLink) {
            if (accumulatedText) {
            // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
            // const processedText = accumulatedText.replace(/\s/g, '&nbsp;')
              const processedText = accumulatedText
              text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
              accumulatedText = ''
            }

            if (styleInfo.hasLink) {
            // const processedText = styleInfo.text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
            // const processedText = styleInfo.text.replace(/\s/g, '&nbsp;')
              const processedText = styleInfo.text
              text += `<span style="${styleInfo.styleText}"><a href="${styleInfo.linkURL}" target="_blank">${processedText}</a></span>`
              prevStyleInfo = null
            } 
            else {
              prevStyleInfo = styleInfo
              accumulatedText = styleInfo.text
            }
          } 
          else accumulatedText += styleInfo.text

        }
        if (accumulatedText && prevStyleInfo) {
        // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
          const processedText = accumulatedText
          text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
        }
      }
    }

    if (listType) text += '</li>'
    else text += '</p>'
  }
  while (listTypes.length > 0) {
    const closedListType = listTypes.pop()
    text += `</${closedListType}>`
  }

  text = replaceNbspByLimit(text, 3)
  text = addStyleToTag(text, 'p', ' margin: 0; padding: 0;')
  const result = checkSpanLastCharIsTonePinyin(text)
  if (result && result.length === 1 && result[0].isTonePinyin) {
    text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; line-break: strict; word-break: keep-all; overflow-wrap: break-word; white-space: nowrap')
  }
  else {
    text = addStyleToTag(text, 'span', 'line-break: strict; overflow-wrap: break-word; white-space: pre-wrap')
  }
  // text = addStyleToTag(text, 'span', 'line-height: 1.15')
  // text = addStyleToTag(text, 'span', 'white-space: pre-wrap;')
  if (text.includes('&nbsp;')) {
    text = text.replace('&nbsp;', ' ')
  }

  text = replaceSpaceInUnderlineSpan(text)

  return text
}

/**
 * 处理HTML字符串：给带下划线的span替换空格为透明占位符
 * @param {string} htmlStr - 传入的元素HTML字符串
 * @returns {string} 处理后的HTML字符串
 */
function replaceSpaceInUnderlineSpan(htmlStr) {
  // 1. 创建临时DOM容器解析HTML字符串（安全解析，不渲染到页面）
  const tempContainer = document.createElement('div')
  tempContainer.innerHTML = htmlStr

  // 2. 获取所有span元素
  const spanList = tempContainer.querySelectorAll('span')

  // 3. 遍历每个span进行处理
  spanList.forEach(span => {
    // 获取元素的行内样式（处理style属性）
    const style = span.style
    
    // 判断：是否包含 text-decoration: underline / underline 相关样式
    const hasUnderline = style.textDecoration.includes('underline') || 
                        style.textDecorationLine === 'underline'

    if (hasUnderline) {
      // 4. 有下划线：替换所有空格为透明占位标签
      // 替换规则：空格 → 透明"占"字标签
      // const replaceStr = '<span style="color:transparent "> </span>'
      const replaceStr = '<span> </span>'
      // span.innerHTML = span.innerHTML.replace(/ /g, replaceStr)
      span.innerHTML = span.innerHTML.replace(/\s/g, replaceStr)
    }
  })

  // 5. 返回处理后的HTML字符串
  return tempContainer.innerHTML
}

/**
 * 判断 HTML 字符串中所有 <span> 的最后一个字符是否是【带声调的拼音字母】
 * @param {string} htmlString - 传入的 HTML 字符串（如 <p><span>piē</span></p>）
 * @returns {Array<{ spanText: string, lastChar: string, isTonePinyin: boolean }>} 每个 span 的判断结果
 */
function checkSpanLastCharIsTonePinyin(htmlString) {
  // 1. 创建临时 DOM 容器解析 HTML
  const tempDiv = document.createElement('div')
  tempDiv.innerHTML = htmlString

  // 2. 取出所有 span 元素
  const spans = tempDiv.querySelectorAll('span')

  if (spans.length === 0) {
    return []
  }

  // 3. 正则：匹配所有带声调的汉语拼音字母（a o e i u ü 四声全覆盖）
  const tonePinyinRegex = /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜ]/i

  // 4. 遍历每个 span，判断最后一个字符
  const result = Array.from(spans).map(span => {
    const text = span.textContent.trim() // 去除空格
    const last3Char = text ? text.charAt(text.length - 3) : '' // 最后一个字符
    const last2Char = text ? text.charAt(text.length - 2) : '' // 最后一个字符
    const lastChar = text ? text.charAt(text.length - 1) : '' // 最后一个字符
    const isTonePinyin = lastChar ? tonePinyinRegex.test(lastChar) : false
    const isTonePinyin1 = last2Char ? tonePinyinRegex.test(last2Char) : false
    const isTonePinyin2 = last3Char ? tonePinyinRegex.test(last3Char) : false

    return {
      spanText: text, // span 内的文本
      lastChar: lastChar, // 最后一个字符
      isTonePinyin: isTonePinyin || isTonePinyin1 || isTonePinyin2 // 是否是带声调拼音字母
    }
  })

  return result
}

/**
 * 给指定 HTML 标签的内联样式追加样式
 * @param {string} htmlStr - 原始 HTML 字符串
 * @param {string} tag - 要处理的标签名，例如：span、p、div、h1
 * @param {string} addStyle - 要追加的样式语句
 * @returns {string} 处理后的 HTML 字符串
 */
function addStyleToTag(htmlStr, tag, addStyle) {
  // 确保样式最后有分号，避免样式出错
  const style = addStyle.trim().endsWith(';') ? addStyle : addStyle + ';'

  // 第一步：给【已有 style 属性】的标签追加样式
  const regex1 = new RegExp(`<${tag}([^>]*)style="([^"]*)"`, 'gi')
  let result = htmlStr.replace(regex1, (match, attr, oldStyles) => {
    return `<${tag}${attr}style="${oldStyles} ${style}"`
  })

  // 第二步：给【没有 style 属性】的标签自动加上 style
  const regex2 = new RegExp(`<${tag}(?!.*style=)`, 'gi')
  result = result.replace(regex2, `<${tag} style="${style}"`)

  return result
}

// function addStyleToSpans(htmlStr, addStyle) {
//   const style = addStyle.trim().endsWith(';') ? addStyle : addStyle + ';'
//   return htmlStr.replace(/<span([^>]*)style="([^"]*)"/gi, (match, attr, old) => {
//     return `<span${attr}style="${old} ${style}"`
//   }).replace(/<span(?!.*style=)/gi, `<span style="${style}"`)
// }

// function replaceMultiNbspBlocks(str) {
//   // 匹配 任意位置 连续 2个及以上的 &nbsp; 全局替换
//   return str.replace(/(&nbsp;){2,}/g, (match) => {
//     // 计算当前这个区块有多少个 &nbsp;
//     const nbspCount = (match.match(/&nbsp;/g) || []).length

//     // 规则：首尾变成普通空格，中间保留 &nbsp;
//     if (nbspCount === 1) {
//       return match
//     }
//     if (nbspCount === 2) {
//       return '  ' // 2个 → 两个普通空格
//     }
//     // 3个及以上：首尾普通空格 + 中间剩下的 &nbsp;
//     return ' ' + '&nbsp;'.repeat(nbspCount - 2) + ' '
//   })
// }

function replaceNbspByLimit(str, minCount) {
  // 匹配全局连续的 &nbsp; 区块
  return str.replace(/(&nbsp;)+/g, (match) => {
    // 计算当前区块有多少个 &nbsp;
    const total = (match.match(/&nbsp;/g) || []).length

    // 小于设定数量 → 不替换，直接返回
    if (total < minCount) return match

    // 达到数量 → 首尾变普通空格，中间保留 &nbsp;
    if (total === 1) return match
    if (total === 2) return '  '

    // 核心：空格 + 中间 &nbsp; + 空格
    return ' ' + '&nbsp;'.repeat(total - 2) + ' '
  })
}

export function getListType(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return ''

  if (pPrNode['a:buChar']) return 'ul'
  if (pPrNode['a:buAutoNum']) return 'ol'
  
  return ''
}
export function getListLevel(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return -1

  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) return parseInt(lvlNode)

  return 0
}

export function genSpanElement(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
  const { styleText, text, hasLink, linkURL } = getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
  // const processedText = text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
  const processedText = text

  if (hasLink) {
    return `<span style="${styleText}"><a href="${linkURL}" target="_blank">${processedText}</a></span>`
  }
  return `<span style="${styleText}">${processedText}</span>`
}

export function getSpanStyleInfo(node, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
  // const lstStyle = textBodyNode['a:lstStyle']
  const lstStyle = getTextByPathList(textBodyNode, ['a:lstStyle'])
  const slideMasterTextStyles = warpObj['slideMasterTextStyles']
  let lvl = 1
  const pPrNode = pNode['a:pPr']
  const pPrAttrsNode = getTextByPathList(pPrNode, ['attrs', 'algn'])
  const pPrAlgnNode = getTextByPathList(pPrNode, ['attrs', 'algn'])
  console.log('(00)-tableEl-text-[pPrAttrsNode]:', pPrAttrsNode)
  console.log('(00)-tableEl-text-[pPrAttrsNode]-[pPrAlgnNode]:', pPrAlgnNode)
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  // let text = node['a:t']
  let text = getTextByPathList(node, ['a:t'])
  if (typeof text !== 'string') text = getTextByPathList(node, ['a:fld', 'a:t'])
  // if (typeof text !== 'string') text = '&nbsp;'
  if (typeof text !== 'string') text = ' '

  let styleText = ''
  const fontColor = getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj)
  const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode)
  const fontType = getFontType(node, type, warpObj, slideLayoutSpNode, slideMasterSpNode, slideMasterTextStyles)
  const fontBold = getFontBold(node)
  const fontItalic = getFontItalic(node)
  const fontDecoration = getFontDecoration(node)
  const fontDecorationLine = getFontDecorationLine(node)
  const fontSpace = getFontSpace(node)
  let shadow = getFontShadow(node, warpObj)
  const subscript = getFontSubscript(node)
  let background = getBackGroundColor(node, warpObj)

  if (!shadow) {
    shadow = getFontOutLine(node, warpObj)
  }
  // if (node['a:rPr']) {
  //   console.log('(00)-pptxtojson-text====node属性a:ln:', text, node['a:rPr']['a:ln'])
  //   const lnInfo = node['a:rPr']['a:ln']
  //   if (lnInfo && lnInfo['a:solidFill']) {
  //     let w = 0
  //     if (lnInfo['attrs'] && lnInfo['attrs']['w']) {
  //       w = lnInfo['attrs']['w'] * RATIO_EMUs_Points
  //     }
  //     const outLineClr = getSolidFill(lnInfo['a:solidFill'], undefined, undefined, warpObj)
  //     // styleText += ``
  //   }
  // }
  if (fontColor) {
    if (typeof fontColor === 'string') styleText += `color: ${fontColor};`
    else if (fontColor.colors) {
      const { colors, path, rot } = fontColor
      // shape  //rect //circle // line
      const getStartPosi = (obj) => {
        let h = 'center'
        let v = 'center'
        const MaxValue = 100000
        if (obj.l && Math.abs(obj.l) === MaxValue) {
          h = obj.l / MaxValue < 0 ? 'left' : 'right'
        }
        else if (obj.r && Math.abs(obj.r) === MaxValue) {
          h = obj.r / MaxValue > 0 ? 'left' : 'right'
        }

        if (obj.b && Math.abs(obj.b) === MaxValue) {
          v = obj.b / MaxValue > 0 ? 'top' : 'bottom'
        }
        else if (obj.t && Math.abs(obj.t) === MaxValue) {
          v = obj.t / MaxValue < 0 ? 'top' : 'bottom'
        }

        if (h === v && h === 'center') {
          return 'center'
        }

        return `${h} ${v}`
      }
      let startPosi
      if (fontColor.pathCenter) {
        startPosi = getStartPosi(fontColor.pathCenter)
      }

      if (path === 'circle') {
        // 只用到你有的颜色和位置
        const colorStops = fontColor.colors
          .map(item => `${item.color} ${item.pos}`)
          .join(', ')
        background = `radial-gradient(circle at ${startPosi}, ${colorStops})`
        styleText += `background: ${background}; background-clip: text; color: transparent;`
      }
      else if (path === 'rect' || path === 'shape') {
        const stops = colors.map(item => `${item.color} ${item.pos}`).join(', ')
        const gradientStyle = `radial-gradient(ellipse at ${startPosi}, ${stops})`
        styleText += `background: ${gradientStyle}; background-clip: text; color: transparent;`
        background = gradientStyle        
      }
      else {
        const stops = colors.map(item => `${item.color} ${item.pos}`).join(', ')
        const gradientStyle = `linear-gradient(${rot + 90}deg, ${stops})`
        styleText += `background: ${gradientStyle}; background-clip: text; color: transparent;`
        background = gradientStyle
      }      
    }
  }
  if (fontSize) styleText += `font-size: ${fontSize};`
  if (fontType) styleText += `font-family: ${fontType};`
  if (fontBold) styleText += `font-weight: ${fontBold};`
  if (fontItalic) styleText += `font-style: ${fontItalic};`
  if (fontDecoration) styleText += `text-decoration: ${fontDecoration};`
  if (fontDecorationLine) styleText += `text-decoration-line: ${fontDecorationLine};`
  if (fontSpace) styleText += `letter-spacing: ${fontSpace};`
  if (subscript) styleText += `vertical-align: ${subscript};`
  if (shadow) styleText += `text-shadow: ${shadow};`
  if (background && !fontColor.colors) styleText += `background: ${background};`

  const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
  const hasLink = linkID && warpObj['slideResObj'][linkID]

  const styleObj = {
    fontSize: fontSize.replace('pt', 'px'),
    color: !fontColor.colors ? fontColor : '',
    fontType,
    fontBold,
    fontItalic,
    fontDecoration,
    fontDecorationLine,
    fontSpace,
    subscript,
    shadow,
    background
  }

  return {
    styleText,
    styleObj,
    text,
    hasLink,
    linkURL: hasLink ? warpObj['slideResObj'][linkID]['target'] : null,
  }
}