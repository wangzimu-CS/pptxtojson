import { getHorizontalAlign, getParagraphSpacing } from './align'
import { getTextByPathList } from './utils'

import {
  getFontType,
  getFontColor,
  getFontSize,
  getFontBold,
  getFontItalic,
  getFontDecoration,
  getFontDecorationLine,
  getFontSpace,
  getFontSubscript,
  getFontShadow,
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
  if (width === cx / 12700) {
    console.log('(00)-genTextBody-:new---切行结果----lines:--判断', '相同')
  }
  else {
    console.log('(00)-genTextBody-:new---切行结果----lines:--判断', '不同！', 'cx:', cx)
  }

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
    console.log('(00)-genTextBody-:new---切行结果----lines:[textBoxWidthPx, charWidth]:', textBoxWidthPx, charWidth)
    console.log('(00)-genTextBody-:new---切行结果----lines:', lines)

    result.push({
      text: fullText,
      lines: lines,
    })
  }

  return result
}


function parsePPTTextToLines(shapeData, width, getStyleUseInfo, paragraphInfo) {
  if (!shapeData || typeof shapeData !== 'object') return []

  const spPr = shapeData['p:spPr']
  if (!spPr) return []

  const xfrm = spPr['a:xfrm']
  if (!xfrm) return []

  const ext = xfrm['a:ext']
  if (!ext || !ext.attrs || ext.attrs.cx === null) return []

  const cx = Number(ext.attrs.cx)
  if (isNaN(cx) || cx <= 0) return []
  const textBoxWidthPx = width

  const txBody = shapeData['p:txBody']
  if (!txBody) return []

  if (!getStyleUseInfo) return []

  const {
    pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj
  } = getStyleUseInfo
  const abs = false
  if (abs) {
    const rNodeItem = {}
    const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
    console.log('(00)styleInfo', styleInfo)
  }

  let paragraphs = txBody['a:p']
  if (!paragraphs) return []
  paragraphs = Array.isArray(paragraphs) ? paragraphs : [paragraphs]

  const result = []

  for (const p of paragraphs) {
    if (!p) continue
    let runs = p['a:r']
    if (!runs) continue
    
    runs = Array.isArray(runs) ? runs : [runs]

    let fullText = ''
    let realFontSize = 36
    const runList = []

    for (const r of runs) {
      if (r && typeof r['a:t'] === 'string') {
        const text = r['a:t']
        const start = fullText.length
        fullText += text
        const end = fullText.length
        const rNodeItem = r
        const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
        console.log('(00)-pptxtojson-text-parsePPTTextToLines-styleInfo:', styleInfo)
        const style = getRunStyle(r, realFontSize)
        console.log('(00)-pptxtojson-text-parsePPTTextToLines-styleInfo:---style:', style)
        let styleText
        if (styleInfo) {
          styleText = styleInfo.styleText
        }
        runList.push({ start, end, text, ...style, styleText })
      }

      if (r && r['a:rPr'] && r['a:rPr'].attrs && r['a:rPr'].attrs.sz) {
        const sz = Number(r['a:rPr'].attrs.sz)
        if (!isNaN(sz) && sz > 0) realFontSize = sz / 100
      }
    }

    const charWidth = realFontSize
    const lines = wrapTextProfessional(fullText, textBoxWidthPx - 10, charWidth)

    // ===================== 正确生成 HTML =====================
    const lineSpans = []
    let currentPos = 0

    for (const line of lines) {
      const lineLen = line.length
      const lineStart = currentPos
      const lineEnd = currentPos + lineLen
      currentPos = lineEnd

      const parts = []
      for (const run of runList) {
        if (run.end <= lineStart || run.start >= lineEnd) continue
        const s = Math.max(run.start, lineStart) - run.start
        const e = Math.min(run.end, lineEnd) - run.start
        const txt = run.text.slice(s, e)

        let styleStr = ''
        for (const key in run.style) {
          styleStr += key + ':' + run.style[key] + ';'
        }
        if (run.styleText) {
          parts.push('<span style="' + run.styleText + '">' + txt + '</span>')
        }
        else {
          parts.push('<span style="' + styleStr + '">' + txt + '</span>')
        }
      }

      // 行 span 无任何样式！
      lineSpans.push('<span style=" white-space: pre ">' + parts.join('') + '</span>')
    }

    const paraStyle = {}
    // const paragraphHtml = '<p style="">' + lineSpans.join('') + '</p>'
    const linesString = ` <span style=" white-space: pre-wrap ">` + lineSpans.join('') + '</span>'
    let paragraphHtml = '<p style="">' + linesString + '</p>'
    if (paragraphInfo && paragraphInfo.start && paragraphInfo.end) {
      paragraphHtml = paragraphInfo.start + linesString + paragraphInfo.end
      console.log('(00)-genTextBody-:new---切行结果----text--------------------:paragraphInfo:', paragraphInfo, paragraphHtml)
      console.log('(00)-genTextBody-:new---切行结果----text--------------------:linesString:', linesString)
      console.log('(00)-genTextBody-:new---切行结果----text--------------------:lines:', lines)
      
    }

    result.push({
      text: fullText,
      lines: lines,
      lineSpans: lineSpans,
      paragraphHtml: paragraphHtml,
      paraStyle: paraStyle
    })
  }
  console.log('(00)-genTextBody-:new---切行结果----text--------------------:result:--result:', result)
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
 * 真正正确的 PPT 中文自动换行
 * 汉字 = 全宽
 * 空格/标点 = 半宽
 * 严格按像素计算
 * 标点不出现在行首
 */
function wrapTextProfessional(text, maxLineWidthPx, fullCharWidth) {
  if (typeof text !== 'string' || text === '') return []
  if (maxLineWidthPx <= 0 || fullCharWidth <= 0) return []

  const halfCharWidth = fullCharWidth * 0.5
  const lines = []
  let currentLine = ''
  let currentWidth = 0

  // 行首禁止出现的标点
  const NO_LINE_START = new Set([
    '，', '。', '、', '；', '：', '）', '”', '！', '？', '…'
  ])

  // 空格、英文标点使用半宽
  const HALF_WIDTH_CHARS = new Set([
    ' ', ' ', ' ', ' ', '\t',
    ',', '.', ';', ':', '!', '?', '"', ')', ']'
  ])

  for (const char of text) {
    // ✅ 关键修复：空格 != 汉字宽度
    const charW = HALF_WIDTH_CHARS.has(char) ? halfCharWidth : fullCharWidth

    // 超宽判断
    if (currentWidth + charW > maxLineWidthPx) {
      // 标点不能放行首
      if (NO_LINE_START.has(char)) {
        currentLine += char
        currentWidth += charW
        lines.push(currentLine)
        currentLine = ''
        currentWidth = 0
      }
      else {
        lines.push(currentLine)
        currentLine = char
        currentWidth = charW
      }
    }
    else {
      currentLine += char
      currentWidth += charW
    }
  }

  if (currentLine !== '') {
    lines.push(currentLine)
  }

  return lines
}

// export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj, width) {
//   if (!textBodyNode) return ''
//   let text = ''

//   console.log('(00)-genTextBody-:textBodyNode:', textBodyNode)
//   // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -textBodyNode:', textBodyNode)
//   const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])
  
//   const pNode = textBodyNode['a:p']
//   // console.log('(00)-genTextBody-:pNode:', pNode)
//   const pNodes = pNode.constructor === Array ? pNode : [pNode]
  
//   const listTypes = []
//   console.log('(00)-genTextBody-:new---spNode:', spNode)
//   console.log('(00)-genTextBody-:new---pNodes:', pNodes)
//   const getStyleInfo = {
//     pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj
//   }
//   const cutLineResult = parsePPTTextToLines(spNode, width, getStyleInfo)
//   // const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode)
//   console.log('(00)-genTextBody-:new---切行结果----cutLineResult:', cutLineResult)
//   for (let i = 0; i < cutLineResult.length;i++ ) {
//     console.log('(00)-genTextBody-:new---切行结果----cutLineResult:.paragraphHtml', cutLineResult[i].paragraphHtml)
//     text += cutLineResult[i].paragraphHtml
//   }
//   if (text.length > 0) {
//     console.log('(00)-genTextBody-:new---切行结果----text', text)
//     return text
//   }
  
//   for (const pNode of pNodes) {
//     // console.log('(00)-genTextBody-:pNode:', pNode)
//     let rNode = pNode['a:r']
//     console.log('(00)-genTextBody-:rNode:', rNode)
//     let fldNode = pNode['a:fld']
//     let brNode = pNode['a:br']
//     if (rNode) {
//       rNode = (rNode.constructor === Array) ? rNode : [rNode]

//       if (fldNode) {
//         fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode]
//         rNode = rNode.concat(fldNode)
//       }
//       if (brNode) {
//         brNode = (brNode.constructor === Array) ? brNode : [brNode]
//         brNode.forEach(item => item.type = 'br')
  
//         if (brNode.length > 1) brNode.shift()
//         rNode = rNode.concat(brNode)
//         rNode.sort((a, b) => {
//           if (!a.attrs || !b.attrs) return true
//           return a.attrs.order - b.attrs.order
//         })
//       }
//     }

//     // 增加整个textbody对齐参信息的获取-解决文本框对齐问题
//     const lstStyle = textBodyNode['a:lstStyle']
//     const lstStyle_align = getTextByPathList(lstStyle, ['a:lvl1pPr', 'attrs', 'algn'])
//     const align = getHorizontalAlign(pNode, spNode, type, warpObj, lstStyle_align)
//     const spacing = getParagraphSpacing(pNode)
//     let styleText = `text-align: ${align};`
//     if (spacing) {
//       // console.log('(00)=-=====>spacing.lineSpacing:', spacing.lineSpacing)
//       if (spacing.lineSpacing) styleText += `line-height: ${spacing.lineSpacing};`
//       if (spacing.spaceBefore) styleText += `margin-top: ${spacing.spaceBefore};`
//       if (spacing.spaceAfter) styleText += `margin-bottom: ${spacing.spaceAfter};`
//     }

//     const listType = getListType(pNode)
//     const listLevel = getListLevel(pNode)

//     if (listType) {
//       while (listTypes.length > listLevel + 1) {
//         const closedListType = listTypes.pop()
//         text += `</${closedListType}>`
//       }

//       if (listTypes[listLevel] === undefined) {
//         text += `<${listType}>`
//         listTypes[listLevel] = listType
//       }
//       else if (listTypes[listLevel] !== listType) {
//         text += `</${listTypes[listLevel]}>`
//         text += `<${listType}>`
//         listTypes[listLevel] = listType
//       }
//       text += `<li style="${styleText}">`
//     }
//     else {
//       while (listTypes.length > 0) {
//         const closedListType = listTypes.pop()
//         text += `</${closedListType}>`
//       }
//       text += `<p style="${styleText}">`
//     }
    
//     if (!rNode) {
//       text += genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
//       console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText--!rNode', genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj))
//     } 
//     else {
//       let prevStyleInfo = null
//       let accumulatedText = ''
//       let defaultLineHight = false
//       let defaultLineHight11 = false
//       const lineHight11 = rNode.length > 1
//       for (const rNodeItem of rNode) {
//         const styleInfo = getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)

//         if (!prevStyleInfo || prevStyleInfo.styleText !== styleInfo.styleText || prevStyleInfo.hasLink !== styleInfo.hasLink || styleInfo.hasLink) {
//           if (accumulatedText) {
//             // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
//             // const processedText = accumulatedText.replace(/\s/g, '&nbsp;')
//             const processedText = accumulatedText
//             text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
//             accumulatedText = ''
//           }

//           if (styleInfo.hasLink) {
//             // const processedText = styleInfo.text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
//             // const processedText = styleInfo.text.replace(/\s/g, '&nbsp;')
//             const processedText = styleInfo.text
//             text += `<span style="${styleInfo.styleText}"><a href="${styleInfo.linkURL}" target="_blank">${processedText}</a></span>`
//             prevStyleInfo = null
//           } 
//           else {
//             prevStyleInfo = styleInfo
//             accumulatedText = styleInfo.text
//           }
//         } 
//         else accumulatedText += styleInfo.text

//         // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -inFor:', accumulatedText)
//         defaultLineHight = styleInfo.lineHight115
//         defaultLineHight11 = styleInfo.lineHight11
//       }
//       if (accumulatedText && prevStyleInfo) {
//         // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
//         const processedText = accumulatedText
//         text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
//       }
//       if (abs) {
//         defaultLineHight
//         lineHight11
//       }
//       // if (defaultLineHight || !(spacing && spacing.lineSpacing)) {
//       if (defaultLineHight) {
//         text = addStyleToTag(text, 'span', 'line-height: 1.15')
//         console.log('(00)-getSpanStyleInfo----aRpr:-----accumulatedText-------------------------------------------------:', accumulatedText)
//       }
//       if (defaultLineHight11) {
//         text = addStyleToTag(text, 'span', 'line-height: 1.1')
//         console.log('(00)-getSpanStyleInfo----aRpr:-----accumulatedText-------------------------------------------------:', accumulatedText)
//       }
//       // else if (lineHight11) {
//       //   text = addStyleToTag(text, 'span', 'line-height: 1.1')
//       // }

//     }

//     if (listType) text += '</li>'
//     else text += '</p>'
//   }
//   while (listTypes.length > 0) {
//     const closedListType = listTypes.pop()
//     text += `</${closedListType}>`
//   }
//   const abs = false
//   if (abs) {
//     text = replaceMultiNbspBlocks(text)
//     text = addStyleToSpans(text, ' line-height: inherit; vertical-align: middle;')
//   }
//   else {
//     text = replaceNbspByLimit(text, 3)
//   }
//   // if (!text.includes('&nbsp;')) {
//   //   text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; word-break: keep-all; white-space: nowrap')
//   //   // console.log('(00)-pptxtojson-genTextBody---text:', '有&nbsp;')
//   // }
//   // else {
//   //   text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; word-break: keep-all;')
//   //   // console.log('(00)-pptxtojson-genTextBody---text:', '没有&nbsp;')
//   // }
//   if (!abs) {
//     text = addStyleToTag(text, 'p', ' margin: 0; padding: 0;')
//     // console.log('(00)-pptxtojson-genTextBody---text:', text)
//     const result = checkSpanLastCharIsTonePinyin(text)
//     if (abs) {
//       // console.log('(00)---checkSpanLastCharIsTonePinyinresult:', result)
//     }
//     // // console.log('(00)-dsfjslkjflas:-result:', result)
//     if (result && result.length === 1 && result[0].isTonePinyin) {
//       text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; line-break: strict; word-break: keep-all; overflow-wrap: break-word; white-space: nowrap')
//       // // console.log('(00)-dsfjslkjflas:-result:23412', '强制不换行')
//     }
//     else {
//       // text = addStyleToTag(text, 'span', 'white-space: pre-wrap; line-height: 2')
//       text = addStyleToTag(text, 'span', 'line-break: strict; overflow-wrap: break-word; white-space: pre-wrap')
//     }
//     // // console.log('(00)-pptxtojson-genTextBody---text:---最终:', text)
//   }
//   // text = addStyleToTag(text, 'span', 'line-height: 1.15')
//   // text = addStyleToTag(text, 'span', 'white-space: pre-wrap;')
//   if (text.includes('&nbsp;')) {
//     console.log('(00)-------text:有&nbsp:', text)
//     text = text.replace('&nbsp;', ' ')
//   }

//   text = replaceSpaceInUnderlineSpan(text)
//   // if (!text.includes('line-height')) {
//   //   text = addStyleToTag(text, 'span', 'line-height: 1.1')
//   // }
//   return text
// }
export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj, width) {
  if (!textBodyNode) return ''
  let text = ''

  console.log('(00)-genTextBody-:textBodyNode:', textBodyNode)
  // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -textBodyNode:', textBodyNode)
  const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])
  
  const pNode = textBodyNode['a:p']
  // console.log('(00)-genTextBody-:pNode:', pNode)
  const pNodes = pNode.constructor === Array ? pNode : [pNode]
  
  const listTypes = []
  console.log('(00)-genTextBody-:new---spNode:', spNode)
  console.log('(00)-genTextBody-:new---pNodes:', pNodes)
  
  for (const pNode of pNodes) {
    // console.log('(00)-genTextBody-:pNode:', pNode)
    let rNode = pNode['a:r']
    console.log('(00)-genTextBody-:rNode:', rNode)
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
    if (spacing) {
      // console.log('(00)=-=====>spacing.lineSpacing:', spacing.lineSpacing)
      if (spacing.lineSpacing) styleText += `line-height: ${spacing.lineSpacing};`
      if (spacing.spaceBefore) styleText += `margin-top: ${spacing.spaceBefore};`
      if (spacing.spaceAfter) styleText += `margin-bottom: ${spacing.spaceAfter};`
    }

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
      console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText--!rNode', genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj))
    } 
    else {
      let prevStyleInfo = null
      let accumulatedText = ''
      let defaultLineHight = false
      let defaultLineHight11 = false
      const lineHight11 = rNode.length > 1

      let newDeal = false
      if (!newDeal) {
        text = ''
        const getStyleInfo = {
          pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj
        }
        const cutLineResultNew = parsePPTTextToLinesNew(spNode, width)
        const cutLineResult = parsePPTTextToLines(spNode, width, getStyleInfo, paragraphInfo)
        // const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode)
        console.log('(00)-genTextBody-:new---切行结果----cutLineResultNew:', cutLineResultNew)
        console.log('(00)-genTextBody-:new---切行结果----cutLineResult:', cutLineResult)
        for (let i = 0; i < cutLineResult.length;i++ ) {
          // console.log('(00)-genTextBody-:new---切行结果----cutLineResult:.paragraphHtml', cutLineResult[i].paragraphHtml)
          console.log('(00)-genTextBody-:new---切行结果----cutLineResult:.lines', i, '-', cutLineResult[i].lines)
          text += cutLineResult[i].paragraphHtml
          newDeal = true
        }
      }

      if (newDeal) {
        console.log('(00)-genTextBody-:new---切行结果----text:::', text)
        // return text
        const abs = false
        if (abs) {
          text = replaceMultiNbspBlocks(text)
          text = addStyleToSpans(text, ' line-height: inherit; vertical-align: middle;')
        }
        else {
          text = replaceNbspByLimit(text, 3)
        }
        if (!abs) {
          text = addStyleToTag(text, 'p', ' margin: 0; padding: 0;')
          const result = checkSpanLastCharIsTonePinyin(text)
          if (abs) {
            // console.log('(00)---checkSpanLastCharIsTonePinyinresult:', result)
            result
          }
          // if (result && result.length === 1 && result[0].isTonePinyin) {
          //   text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; line-break: strict; word-break: keep-all; overflow-wrap: break-word; white-space: nowrap')
          // }
          // else {
          //   text = addStyleToTag(text, 'span', 'line-break: strict; overflow-wrap: break-word; white-space: pre-wrap')
          // }

        }

        if (text.includes('&nbsp;')) {
          console.log('(00)-------text:有&nbsp:', text)
          text = text.replace('&nbsp;', ' ')
        }
        console.log('(00)-genTextBody-:new---切行结果----text:::返回值------:', text)
        return text
      }
      
      else if (!newDeal) {
        //    
        for (const rNodeItem of rNode) {
          console.log('(00)-genTextBody-:new---切行结果----rNodeItem', rNodeItem)
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

          // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -inFor:', accumulatedText)
          defaultLineHight = styleInfo.lineHight115
          defaultLineHight11 = styleInfo.lineHight11
        }
        if (accumulatedText && prevStyleInfo) {
        // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
          const processedText = accumulatedText
          text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
        }
      }

      


      if (abs) {
        defaultLineHight
        lineHight11
      }
      // if (defaultLineHight || !(spacing && spacing.lineSpacing)) {
      if (defaultLineHight) {
        text = addStyleToTag(text, 'span', 'line-height: 1.15')
        console.log('(00)-getSpanStyleInfo----aRpr:-----accumulatedText-------------------------------------------------:', accumulatedText)
      }
      if (defaultLineHight11) {
        text = addStyleToTag(text, 'span', 'line-height: 1.1')
        console.log('(00)-getSpanStyleInfo----aRpr:-----accumulatedText-------------------------------------------------:', accumulatedText)
      }
      // else if (lineHight11) {
      //   text = addStyleToTag(text, 'span', 'line-height: 1.1')
      // }

    }

    if (listType) text += '</li>'
    else text += '</p>'
  }
  while (listTypes.length > 0) {
    const closedListType = listTypes.pop()
    text += `</${closedListType}>`
  }
  const abs = false
  if (abs) {
    text = replaceMultiNbspBlocks(text)
    text = addStyleToSpans(text, ' line-height: inherit; vertical-align: middle;')
  }
  else {
    text = replaceNbspByLimit(text, 3)
  }
  // if (!text.includes('&nbsp;')) {
  //   text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; word-break: keep-all; white-space: nowrap')
  //   // console.log('(00)-pptxtojson-genTextBody---text:', '有&nbsp;')
  // }
  // else {
  //   text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; word-break: keep-all;')
  //   // console.log('(00)-pptxtojson-genTextBody---text:', '没有&nbsp;')
  // }
  if (!abs) {
    text = addStyleToTag(text, 'p', ' margin: 0; padding: 0;')
    // console.log('(00)-pptxtojson-genTextBody---text:', text)
    const result = checkSpanLastCharIsTonePinyin(text)
    if (abs) {
      // console.log('(00)---checkSpanLastCharIsTonePinyinresult:', result)
    }
    // // console.log('(00)-dsfjslkjflas:-result:', result)
    if (result && result.length === 1 && result[0].isTonePinyin) {
      text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; line-break: strict; word-break: keep-all; overflow-wrap: break-word; white-space: nowrap')
      // // console.log('(00)-dsfjslkjflas:-result:23412', '强制不换行')
    }
    else {
      // text = addStyleToTag(text, 'span', 'white-space: pre-wrap; line-height: 2')
      text = addStyleToTag(text, 'span', 'line-break: strict; overflow-wrap: break-word; white-space: pre-wrap')
    }
    // // console.log('(00)-pptxtojson-genTextBody---text:---最终:', text)
  }
  // text = addStyleToTag(text, 'span', 'line-height: 1.15')
  // text = addStyleToTag(text, 'span', 'white-space: pre-wrap;')
  if (text.includes('&nbsp;')) {
    console.log('(00)-------text:有&nbsp:', text)
    text = text.replace('&nbsp;', ' ')
  }

  text = replaceSpaceInUnderlineSpan(text)
  // if (!text.includes('line-height')) {
  //   text = addStyleToTag(text, 'span', 'line-height: 1.1')
  // }
  console.log('(00)-genTextBody-:new---切行结果----text:::返回值------:--!!!!!', text)
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

  console.log('(00)-replaceSpaceInUnderlineSpan--tempContainer.innerHTML:', tempContainer.innerHTML)
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

function addStyleToSpans(htmlStr, addStyle) {
  const style = addStyle.trim().endsWith(';') ? addStyle : addStyle + ';'
  return htmlStr.replace(/<span([^>]*)style="([^"]*)"/gi, (match, attr, old) => {
    return `<span${attr}style="${old} ${style}"`
  }).replace(/<span(?!.*style=)/gi, `<span style="${style}"`)
}

function replaceMultiNbspBlocks(str) {
  // 匹配 任意位置 连续 2个及以上的 &nbsp; 全局替换
  return str.replace(/(&nbsp;){2,}/g, (match) => {
    // 计算当前这个区块有多少个 &nbsp;
    const nbspCount = (match.match(/&nbsp;/g) || []).length

    // 规则：首尾变成普通空格，中间保留 &nbsp;
    if (nbspCount === 1) {
      return match
    }
    if (nbspCount === 2) {
      return '  ' // 2个 → 两个普通空格
    }
    // 3个及以上：首尾普通空格 + 中间剩下的 &nbsp;
    return ' ' + '&nbsp;'.repeat(nbspCount - 2) + ' '
  })
}

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
  const lstStyle = textBodyNode['a:lstStyle']
  const slideMasterTextStyles = warpObj['slideMasterTextStyles']
  // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText-text:', '1111111')
  // console.log('(00)-genTextBody-getSpanStyleInfo--lstStyle:', lstStyle)
  let lvl = 1
  const pPrNode = pNode['a:pPr']
  const lvlNode = getTextByPathList(pPrNode, ['attrs', 'lvl'])
  if (lvlNode !== undefined) lvl = parseInt(lvlNode) + 1

  let text = node['a:t']
  // console.log('(00)-pptxtojson-getSpanStyleInfo---text:', text)
  // console.log('(00)-pptxtojson-genSpanEl---node:', node)
  // console.log('(00)-pptxtojson-genSpanEl---text:', text)
  if (typeof text !== 'string') text = getTextByPathList(node, ['a:fld', 'a:t'])
  // if (typeof text !== 'string') text = '&nbsp;'
  if (typeof text !== 'string') text = ' '
  // // console.log('(00)-pptxtojson-getSpanStyleInfo---text:.includes(/\s/g):', text.includes(/\s/g))
  if (text.includes('\t')) {
    // console.log('(00)-pptxtojson-getSpanStyleInfo---text:.includes(/\s/g):', '有回车')
  }
  if (text.includes('\s')) {
    // console.log('(00)-pptxtojson-getSpanStyleInfo---text:.includes(/\s/g):', '有空格')
  }

  let styleText = ''
  const fontColor = getFontColor(node, pNode, lstStyle, pFontStyle, lvl, warpObj)
  const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, textBodyNode, pNode)
  console.log('(00)-genTextBody-:new-----fontSize:)', fontSize)
  const fontType = getFontType(node, type, warpObj, slideLayoutSpNode, slideMasterSpNode, slideMasterTextStyles)
  const fontBold = getFontBold(node)
  const fontItalic = getFontItalic(node)
  const fontDecoration = getFontDecoration(node)
  const fontDecorationLine = getFontDecorationLine(node)
  const fontSpace = getFontSpace(node)
  const shadow = getFontShadow(node, warpObj)
  const subscript = getFontSubscript(node)

  if (fontDecoration) {
    console.log('(00)-genTextBody-:rNode:--getSpanStyleInfo[fontDecoration]', fontDecoration)
  }

  if (fontDecorationLine) {
    console.log('(00)-genTextBody-:rNode:--getSpanStyleInfo[fontDecorationLine]', fontDecorationLine)
  }

  if (fontColor) {
    if (typeof fontColor === 'string') styleText += `color: ${fontColor};`
    else if (fontColor.colors) {
      const { colors, rot } = fontColor
      const stops = colors.map(item => `${item.color} ${item.pos}`).join(', ')
      const gradientStyle = `linear-gradient(${rot + 90}deg, ${stops})`
      styleText += `background: ${gradientStyle}; background-clip: text; color: transparent;`
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

  const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
  const hasLink = linkID && warpObj['slideResObj'][linkID]

  const aRpr = getTextByPathList(node, ['a:rPr', 'attrs'])
  let lineHight115
  let lineHight11
  if (aRpr && aRpr.b) {
    // lineHight115 = aRpr.sz === '2800' && text.length > 40 && aRpr.b === '1'
    lineHight115 = aRpr.sz === '2800' && aRpr.b === '1' && text.length > 40
    lineHight11 = lineHight115 || aRpr.sz === '2800' && aRpr.b === '1' && text.length > 20
    // lineHight115 = aRpr.b === '1'
    // lineHight115 = aRpr.sz === '2800' 
    // if (lineHight115) {
    // console.log('(00)-getSpanStyleInfo----aRpr:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
    // }
    // if (text.length > 10) {
    //   console.log('(00)-getSpanStyleInfo----aRpr:------>10:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
    // }
    // console.log('(00)-getSpanStyleInfo----aRpr:------>10:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
  }

  return {
    styleText,
    text,
    hasLink,
    linkURL: hasLink ? warpObj['slideResObj'][linkID]['target'] : null,
    lineHight115,
    lineHight11
  }
}