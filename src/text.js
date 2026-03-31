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

export function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
  if (!textBodyNode) return ''
  let text = ''

  console.log('(00)-genTextBody-:textBodyNode:', textBodyNode)
  // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -textBodyNode:', textBodyNode)
  const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])

  const pNode = textBodyNode['a:p']
  // console.log('(00)-genTextBody-:pNode:', pNode)
  const pNodes = pNode.constructor === Array ? pNode : [pNode]

  const listTypes = []

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
    }
    else {
      while (listTypes.length > 0) {
        const closedListType = listTypes.pop()
        text += `</${closedListType}>`
      }
      text += `<p style="${styleText}">`
    }
    
    if (!rNode) {
      text += genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
      console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText--!rNode', genSpanElement(pNode, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj))
    } 
    else {
      let prevStyleInfo = null
      let accumulatedText = ''
      let defaultLineHight = false
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

        // console.log('(00)-pptxtojson-genTextBody---text:--accumulatedText -inFor:', accumulatedText)
        defaultLineHight = styleInfo.lineHight115
      }
      if (accumulatedText && prevStyleInfo) {
        // const processedText = accumulatedText.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, '&nbsp;')
        const processedText = accumulatedText
        text += `<span style="${prevStyleInfo.styleText}">${processedText}</span>`
      }
      if (defaultLineHight) {
        text = addStyleToTag(text, 'span', 'line-height: 1.15')
        console.log('(00)-getSpanStyleInfo----aRpr:-----accumulatedText-------------------------------------------------:', accumulatedText)
      }
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
      text = addStyleToTag(text, 'span', ' line-height: inherit; vertical-align: middle; word-break: keep-all; white-space: nowrap')
      // // console.log('(00)-dsfjslkjflas:-result:23412', '强制不换行')
    }
    else {
      // text = addStyleToTag(text, 'span', 'white-space: pre-wrap; line-height: 2')
      text = addStyleToTag(text, 'span', 'white-space: pre-wrap')
    }
    // // console.log('(00)-pptxtojson-genTextBody---text:---最终:', text)
  }
  // text = addStyleToTag(text, 'span', 'line-height: 1.15')
  // text = addStyleToTag(text, 'span', 'white-space: pre-wrap;')
  if (text.includes('&nbsp;')) {
    console.log('(00)-------text:有&nbsp:', text)
    text = text.replace('&nbsp;', ' ')
  }
  return text
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
  if (aRpr && aRpr.sz) {
    lineHight115 = aRpr.sz === '2800' && text.length > 40
    if (lineHight115) {
      console.log('(00)-getSpanStyleInfo----aRpr:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
    }
    // if (text.length > 10) {
    //   console.log('(00)-getSpanStyleInfo----aRpr:------>10:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
    // }
    console.log('(00)-getSpanStyleInfo----aRpr:------>10:', aRpr.sz, '--lineHight115:', lineHight115, '---text:', text)
  }

  return {
    styleText,
    text,
    hasLink,
    linkURL: hasLink ? warpObj['slideResObj'][linkID]['target'] : null,
    lineHight115
  }
}