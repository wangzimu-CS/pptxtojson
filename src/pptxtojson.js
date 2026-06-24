import JSZip from 'jszip'
import { readXmlFile, simplifyLostLess } from './readXmlFile'
import { getBorder } from './border'
import { getSlideBackgroundFill, getShapeFill, getSolidFill, getPicFill, getPicFilters, dealGradientFill, getPicFillOpacity } from './fill'
import { getChartInfo, getChartTitle, getChartLegend, getChartElPrInfo } from './chart'
import { extractProperties } from './chart_wps'
import { getVerticalAlign, getTextAutoFit } from './align'
import { getPosition, getSize } from './position'
import { genTextBody } from './text'
import { getCustomShapePath, identifyShape } from './shape'
import { extractFileExtension, base64ArrayBuffer, getTextByPathList, angleToDegrees, getMimeType, isVideoLink, escapeHtml, hasValidText, numberToFixed } from './utils'
import { getShadow } from './shadow'
import { getGlow, getSoftEdge } from './glow'
import { getTableBorders, getTableCellParams, getTableRowParams } from './table'
import { RATIO_EMUs_Points } from './constants'
import { findOMath, latexFormart, parseOMath } from './math'
import { getShapePath } from './shapePath'
import { parseTransition, findTransitionNode } from './animation'
import { getSmartArtTextData } from './diagram'
import * as txml from 'txml/dist/txml.mjs'

export async function parse(file) {
  const slides = []
  
  const zip = await JSZip.loadAsync(file)

  const filesInfo = await getContentTypes(zip)
  const { width, height, defaultTextStyle } = await getSlideInfo(zip)
  const { themeContent, themeColors } = await getTheme(zip)

  for (const filename of filesInfo.slides) {
    const singleSlide = await processSingleSlide(zip, filename, themeContent, defaultTextStyle)
    slides.push(singleSlide)
  }

  return {
    slides,
    themeColors,
    size: {
      width,
      height,
    },
  }
}

async function getContentTypes(zip) {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson['Types']['Override']
  let slidesLocArray = []
  let slideLayoutsLocArray = []

  for (const item of subObj) {
    switch (item['attrs']['ContentType']) {
      case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
        slidesLocArray.push(item['attrs']['PartName'].substr(1))
        break
      case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
        slideLayoutsLocArray.push(item['attrs']['PartName'].substr(1))
        break
      default:
    }
  }
  
  const sortSlideXml = (p1, p2) => {
    const n1 = +/(\d+)\.xml/.exec(p1)[1]
    const n2 = +/(\d+)\.xml/.exec(p2)[1]
    return n1 - n2
  }
  slidesLocArray = slidesLocArray.sort(sortSlideXml)
  slideLayoutsLocArray = slideLayoutsLocArray.sort(sortSlideXml)
  
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  }
}

async function getSlideInfo(zip) {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
  const defaultTextStyle = content['p:presentation']['p:defaultTextStyle']
  return {
    width: parseInt(sldSzAttrs['cx']) * RATIO_EMUs_Points,
    height: parseInt(sldSzAttrs['cy']) * RATIO_EMUs_Points,
    defaultTextStyle,
  }
}

async function getTheme(zip) {
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent['Relationships']['Relationship']
  let themeURI

  if (relationshipArray.constructor === Array) {
    for (const relationshipItem of relationshipArray) {
      if (relationshipItem['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
        themeURI = relationshipItem['attrs']['Target']
        break
      }
    }
  } 
  else if (relationshipArray['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
    themeURI = relationshipArray['attrs']['Target']
  }

  const themeContent = await readXmlFile(zip, 'ppt/' + themeURI)

  const themeColors = []
  const clrScheme = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme'])
  if (clrScheme) {
    for (let i = 1; i <= 6; i++) {
      if (clrScheme[`a:accent${i}`] === undefined) break
      const color = getTextByPathList(clrScheme, [`a:accent${i}`, 'a:srgbClr', 'attrs', 'val'])
      if (color) themeColors.push('#' + color)
    }
  }
  return { themeContent, themeColors }
}

async function processSingleSlide(zip, sldFileName, themeContent, defaultTextStyle) {
  const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
  const resContent = await readXmlFile(zip, resName)
  let relationshipArray = resContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  
  let noteFilename = ''
  let layoutFilename = ''
  let masterFilename = ''
  let themeFilename = ''
  let diagramFilename = ''
  const diagramFiles = {}
  const slideResObj = {}
  const layoutResObj = {}
  const masterResObj = {}
  const themeResObj = {}
  const diagramResObj = {}

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    if (relType.includes('/www.wps.cn')) {
      console.log('(00)-pptxtojson-[chartEL]:-pie:-an:--relTarget:', relTarget)
      console.log('(00)-pptxtojson-[chartEL]:-pie:-an:--relType:', relType)
    }
    const isExternal = relationshipArrayItem['attrs']['TargetMode'] === 'External'
    if (!isExternal) {
      if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
      else relTarget = 'ppt/slides/' + relTarget
    }

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
        layoutFilename = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        noteFilename = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
        diagramFilename = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors':
        diagramFiles[relationshipArrayItem['attrs']['Id']] = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
      default:
        if (relType.includes('/www.wps.cn')) {
          const wpsOBJ = await readXmlFile(zip, relTarget)
          console.log('(00)-pptxtojson-[chartEL]:-pie:-an:--wpsOBJ:', wpsOBJ)
        }
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }
  
  const slideNotesContent = await readXmlFile(zip, noteFilename)
  const note = getNote(slideNotesContent)

  const slideLayoutContent = await readXmlFile(zip, layoutFilename)
  const slideLayoutTables = await indexNodes(slideLayoutContent)
  const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
  const slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename)
  relationshipArray = slideLayoutResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
    else relTarget = 'ppt/slideLayouts/' + relTarget

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
        masterFilename = relTarget
        break
      default:
        layoutResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }

  const slideMasterContent = await readXmlFile(zip, masterFilename)
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
  const slideMasterTables = indexNodes(slideMasterContent)
  const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
  const slideMasterResContent = await readXmlFile(zip, slideMasterResFilename)
  relationshipArray = slideMasterResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
    else relTarget = 'ppt/slideMasters/' + relTarget

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
        themeFilename = relTarget
        break
      default:
        masterResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }

  if (themeFilename) {
    const themeName = themeFilename.split('/').pop()
    const themeResFileName = themeFilename.replace(themeName, '_rels/' + themeName) + '.rels'
    console.log('(00)-pptxtojson-[chartEL]:-genChart-[refName]:-themeResFileName:', '【', themeResFileName, '】【', themeFilename, '】')
    const themeResContent = await readXmlFile(zip, themeResFileName)
    if (themeResContent) {
      relationshipArray = themeResContent['Relationships']['Relationship']
      if (relationshipArray) {
        if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
        for (const relationshipArrayItem of relationshipArray) {
          themeResObj[relationshipArrayItem['attrs']['Id']] = {
            'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
        }
      }
    }
  }

  const diagramContent = {
    data: null,
    layout: null,
    quickStyle: null,
    colors: null,
    drawing: null
  }
  let digramFileContent = {}
  if (diagramFilename) {
    const diagName = diagramFilename.split('/').pop()
    const diagramResFileName = diagramFilename.replace(diagName, '_rels/' + diagName) + '.rels'
    console.log('(00)-pptxtojson-[chartEL]:-genChart-[refName]:-diagramResFileName:', diagramResFileName)
    digramFileContent = await readXmlFile(zip, diagramFilename)
    if (digramFileContent) {
      const digramFileContentObjToStr = JSON.stringify(digramFileContent).replace(/dsp:/g, 'p:')
      digramFileContent = JSON.parse(digramFileContentObjToStr)
    }
    const digramResContent = await readXmlFile(zip, diagramResFileName)
    if (digramResContent) {
      relationshipArray = digramResContent['Relationships']['Relationship']
      if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
      for (const relationshipArrayItem of relationshipArray) {
        diagramResObj[relationshipArrayItem['attrs']['Id']] = {
          'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
        }
      }
    }
  }

  if (Object.values(diagramFiles).length > 0) {
    for (const filePath of Object.values(diagramFiles)) {
      const content = await readXmlFile(zip, filePath)
      if (filePath.includes('/data')) diagramContent.data = content
      else if (filePath.includes('/layout')) diagramContent.layout = content
      else if (filePath.includes('/quickStyle')) diagramContent.quickStyle = content
      else if (filePath.includes('/colors')) diagramContent.colors = content
    }
  }

  const tableStyles = await readXmlFile(zip, 'ppt/tableStyles.xml')

  const slideContent = await readXmlFile(zip, sldFileName)
  const slide1Xml = await zip.file(sldFileName).async('text')
  let spNodes
  if (slide1Xml) {
    spNodes = parseXMLData(slide1Xml, 'spTree')
  }
  // let nodes = slideContent['p:sld']['p:cSld']['p:spTree']
  let nodes 
  if (slideContent) {
    nodes = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:spTree'])
  }
  if (spNodes.length > 0) {
    nodes = getXMLNodeData(spNodes[0], ['p:spTree'])
  }
  const warpObj = {
    zip,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    tableStyles,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent,
    themeResObj,
    digramFileContent,
    diagramResObj,
    diagramContent,
    defaultTextStyle,
  }
  const layoutElements = await getLayoutElements(warpObj)
  const fill = await getSlideBackgroundFill(warpObj)

  const elements = []
  for (const nodeKey in nodes) {
    if (nodes[nodeKey].constructor !== Array) nodes[nodeKey] = [nodes[nodeKey]]
    for (const node of nodes[nodeKey]) {
      const ret = await processNodesInSlide(nodeKey, node, warpObj, 'slide')
      if (ret) elements.push(ret)
    }
  }

  let transitionNode = findTransitionNode(slideContent, 'p:sld')
  if (!transitionNode) transitionNode = findTransitionNode(slideLayoutContent, 'p:sldLayout')
  if (!transitionNode) transitionNode = findTransitionNode(slideMasterContent, 'p:sldMaster')

  const transition = parseTransition(transitionNode)

  return {
    fill,
    elements,
    layoutElements,
    note,
    transition,
  }
}

function parseXMLData(slideXml, key) {
  if (!slideXml) {
    return
  } 
  const ns = {
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main' 
  }
  const parser = new DOMParser()
  const xmlDoc = parser.parseFromString(slideXml, 'text/xml')
  if (!xmlDoc) return
  // 获取所有动画节点
  // const ctns = xmlDoc.getElementsByTagNameNS(ns.p, 'cTn')
  const spList = xmlDoc.getElementsByTagNameNS(ns.p, key)
  // const animNodes = Array.from(ctns).filter((v) => v.getAttribute('presetClass'))
  return spList
}

function getXMLNode(node) {
  // 创建 XML 序列化器
  const serializer = new XMLSerializer()
  // 将节点序列化为 XML 字符串
  const animEffectXmlString = serializer.serializeToString(node)
  const nodeObj = dealXmlData(animEffectXmlString)
  return nodeObj
}

function getXMLNodeData(node, keys) {
  const nodeObj = getXMLNode(node)
  const value = getTextByPathList(nodeObj, keys)
  return value
}

/**
 * 处理XML字符串数据
 * @param data XML字符串
 * @returns 简化后的XML解析结果
 */
export function dealXmlData(data) {
  if (data) {
    const xmlData = txml.parse(data, {
      keepWhitespace: true // 禁用首尾空白修剪
    }) 
      
    // const DeletSpaceData = JSON.parse(JSON.stringify(xmlData).replace('"\\r\\n",', ''))
    // return simplifyLostLess(DeletSpaceData)
    return simplifyLostLess(xmlData)
  }
    
  return null
}


function getHyperlinkFromCNvPr(cNvPr, warpObj) {
  const hlinkClick = getTextByPathList(cNvPr, ['a:hlinkClick', 'attrs'])
  if (!hlinkClick) return null

  const linkId = hlinkClick['r:id']
  if (!linkId) return null

  const res = warpObj['slideResObj'][linkId]
  if (!res) return null

  if (res['type'] !== 'hyperlink') return null

  const target = res['target']
  if (!target || !/^https?:\/\//.test(target)) return null

  return target
}

function getNote(noteContent) {
  let text = ''
  let spNodes = getTextByPathList(noteContent, ['p:notes', 'p:cSld', 'p:spTree', 'p:sp'])
  if (!spNodes) return ''

  if (spNodes.constructor !== Array) spNodes = [spNodes]
  for (const spNode of spNodes) {
    const phType = getTextByPathList(spNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
    if (phType !== 'body') continue

    const textBody = getTextByPathList(spNode, ['p:txBody'])
    if (!textBody) continue

    let pNode = textBody['a:p']
    if (!pNode) continue
    if (pNode.constructor !== Array) pNode = [pNode]

    const listTypes = []

    for (const p of pNode) {
      const pPr = p['a:pPr']
      const algn = getTextByPathList(pPr, ['attrs', 'algn'])
      let align = 'left'
      if (algn) {
        switch (algn) {
          case 'r': align = 'right'; break
          case 'ctr': align = 'center'; break
          case 'just': case 'dist': align = 'justify'; break
          default: break
        }
      }

      let listType = ''
      if (pPr) {
        if (pPr['a:buChar']) listType = 'ul'
        else if (pPr['a:buAutoNum']) listType = 'ol'
      }
      const lvlNode = getTextByPathList(pPr, ['attrs', 'lvl'])
      const listLevel = lvlNode !== undefined ? parseInt(lvlNode) : 0
      if (listType) {
        while (listTypes.length > listLevel + 1) {
          text += `</${listTypes.pop()}>`
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
        text += `<li style="text-align:${align};">`
      }
      else {
        while (listTypes.length > 0) {
          text += `</${listTypes.pop()}>`
        }
        text += `<p style="text-align:${align};">`
      }

      let rNodes = p['a:r']
      if (rNodes) {
        if (rNodes.constructor !== Array) rNodes = [rNodes]
        for (const r of rNodes) {
          const t = getTextByPathList(r, ['a:t'])
          if (t && typeof t === 'string') text += t
        }
      }

      if (listType) text += '</li>'
      else text += '</p>'
    }
    while (listTypes.length > 0) {
      text += `</${listTypes.pop()}>`
    }
  }
  return text
}

async function getLayoutElements(warpObj) {
  const elements = []
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  const nodesSldLayout = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree'])
  const nodesSldMaster = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree'])
  // 获取是否阻断集成的信息
  const isPreserve = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'preserve']) === '1'
  // 母版
  const sldLayOutElementList = []
  const sldMasterElementList = []
  const showMasterSp = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp'])
  if (nodesSldLayout) {
    for (const nodeKey in nodesSldLayout) {
      if (nodesSldLayout[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
          const ph = getTextByPathList(nodesSldLayout[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
          if (!ph) {
            const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], warpObj, 'slideLayoutBg')
            if (ret) sldLayOutElementList.push(ret) // elements.push(ret)
          }
        }
      } 
      else {
        const ph = getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
        if (!ph) {
          const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], warpObj, 'slideLayoutBg')
          if (ret) sldLayOutElementList.push(ret) // elements.push(ret)
        }
      }
    }
  }
  if (nodesSldMaster && showMasterSp !== '0') {
    for (const nodeKey in nodesSldMaster) {
      if (nodesSldMaster[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
          const ph = getTextByPathList(nodesSldMaster[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
          if (!ph) {
            const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], warpObj, 'slideMasterBg')
            if (ret) sldMasterElementList.push(ret)// elements.push(ret)
          }
        }
      } 
      else {
        const ph = getTextByPathList(nodesSldMaster[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
        if (!ph) {
          const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], warpObj, 'slideMasterBg')
          if (ret) sldMasterElementList.push(ret)// elements.push(ret)
        }
      }
    }
  }
  const abs = false
  if (abs) {
    findSameSubObjects(sldLayOutElementList, sldMasterElementList)
  }
  for (let i = 0;i < sldLayOutElementList.length;i++) {
    const target = sldLayOutElementList[i]
    if (!isPreserve) {
      elements.push(target)
    }
    else {
      // 阻断集成的LayOut需要判断元素是否继承自母版
      if (!isObjectInArrayById(target, sldMasterElementList)) {
        elements.push(target)
      }
    }
  }
  return elements
}
function isObjectInArrayById(target, arr) {
  return arr.some(item => item.id === target.id)
}
/**
 * 深度对比两个对象，找出完全相同的子属性/子对象
 * @param {Object} obj1 第一个对象
 * @param {Object} obj2 第二个对象
 * @returns {Array} 完全相同的键名数组（支持嵌套路径）
 */
function findSameSubObjects(obj1, obj2) {
  const sameKeys = []

  // 深度对比工具函数
  function deepCompare(target1, target2, currentPath = '') {
    // 类型不同 → 直接不相等
    if (typeof target1 !== typeof target2) return false
    
    // 处理 null / undefined
    if (target1 === null && target2 === null) return true
    if (target1 === undefined && target2 === undefined) return true

    // 基础类型（字符串/数字/布尔）直接对比值
    if (typeof target1 !== 'object') {
      return target1 === target2
    }

    // 数组对比
    if (Array.isArray(target1) && Array.isArray(target2)) {
      if (target1.length !== target2.length) return false
      for (let i = 0; i < target1.length; i++) {
        if (!deepCompare(target1[i], target2[i], `${currentPath}[${i}]`)) {
          return false
        }
      }
      return true
    }

    // 对象对比
    const keys1 = Object.keys(target1)
    const keys2 = Object.keys(target2)

    // 键数量不同 → 不相等
    if (keys1.length !== keys2.length) return false

    let isEqual = true
    for (const key of keys1) {
      const path = currentPath ? `${currentPath}.${key}` : key
      const val1 = target1[key]
      const val2 = target2[key]

      // 递归对比子值
      const childEqual = deepCompare(val1, val2, path)
      if (!childEqual) isEqual = false

      // 子对象/子值完全相等 → 记录路径
      if (childEqual) {
        sameKeys.push(path)
      }
    }

    return isEqual
  }

  deepCompare(obj1, obj2)
  return sameKeys
}

function indexNodes(content) {
  const keys = Object.keys(content)
  const spTreeNode = content[keys[0]]['p:cSld']['p:spTree']
  const idTable = {}
  const idxTable = {}
  const typeTable = {}

  for (const key in spTreeNode) {
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue

    const targetNode = spTreeNode[key]

    if (targetNode.constructor === Array) {
      for (const targetNodeItem of targetNode) {
        const nvSpPrNode = targetNodeItem['p:nvSpPr']
        const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
        const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
        const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

        if (id) idTable[id] = targetNodeItem
        if (idx) idxTable[idx] = targetNodeItem
        if (type) typeTable[type] = targetNodeItem
      }
    } 
    else {
      const nvSpPrNode = targetNode['p:nvSpPr']
      const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
      const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
      const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

      if (id) idTable[id] = targetNode
      if (idx) idxTable[idx] = targetNode
      if (type) typeTable[type] = targetNode
    }
  }

  return { idTable, idxTable, typeTable }
}

async function processNodesInSlide(nodeKey, nodeValue, warpObj, source, groupHierarchy = []) {
  let json

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = await processSpNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'p:cxnSp': // Shape, Text
      json = await processCxnSpNode(nodeValue, warpObj, source)
      break
    case 'p:pic': // Image, Video, Audio
      json = await processPicNode(nodeValue, warpObj, source)
      console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:', nodeValue)
      const webExtensionRefNode = getTextByPathList(nodeValue, ['p:spPr', 'a:extLst', 'a:ext', 'wpswe:webExtensionRef', 'attrs'])
      console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:', webExtensionRefNode, warpObj)
      if (webExtensionRefNode) {
        const rid = getTextByPathList(webExtensionRefNode, ['r:id'])
        if (rid) {
          //
          const slideResObj = getTextByPathList(warpObj, ['slideResObj'])
          const targetNode = getTextByPathList(slideResObj, [`${rid}`])
          console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-rid-slideResObj:', rid, slideResObj)
          console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-rid:', rid, targetNode)
          const {type, target} = targetNode
          if (type.includes('/www.wps.cn')) {
            const useTarget = target.replace('..', 'ppt')
            const targetNode = await readXmlFile(warpObj['zip'], useTarget)
            console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-targetNode:', targetNode)
            const exTensionId = getTextByPathList(targetNode, ['wpswe:webExtension', 'wpswe:extSource', 'attrs', 'id'])
            if (exTensionId === 'webchart') {
              const dealNode = extractProperties(targetNode)
              // json = genWPSWebChart(nodeValue, warpObj, source)
              // json = genWPSWebChart(dealNode, warpObj, source)
              // genWPSWebChart(dealNode, warpObj, source)
              const genWPSWebChartNode = genWPSWebChart(json, dealNode)
              json = genWPSWebChartNode
              // console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-dealNode:', dealNode, targetNode, genWPSWebChartNode)
              console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-genWPSWebChartNode:', genWPSWebChartNode)
            }
          }
        }
      }
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source)
      console.log('(00)-pptxtojson-[chartEL]:-json:', json)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'mc:AlternateContent':
      if (getTextByPathList(nodeValue, ['mc:Fallback', 'p:grpSpPr', 'a:xfrm'])) {
        json = await processGroupSpNode(getTextByPathList(nodeValue, ['mc:Fallback']), warpObj, source, groupHierarchy)
      }
      else if (getTextByPathList(nodeValue, ['mc:Choice'])) {
        json = await processMathNode(nodeValue, warpObj, source)
      }
      break
    default:
  }
  if (['p:sp', 'p:cxnSp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'mc:AlternateContent'].includes(nodeKey)) {
    let targetKey = nodeKey.replace('p:', '')
    const parserFirst = (str) => {
      if (!str) return str
      const [first, ...rest] = str 
      return first.toUpperCase() + rest.join('')
    }
    targetKey = parserFirst(targetKey)
    if ( nodeKey === 'p:cxnSp' ) {
      targetKey = 'Sp'
    }
    const id = getTextByPathList(nodeValue, [`p:nv${targetKey}Pr`, 'p:cNvPr', 'attrs', 'id'])
    const pr = getTextByPathList(nodeValue, ['p:spPr'])
    const useBgFill = getTextByPathList(nodeValue, ['attrs', 'useBgFill'])
    if (json && id) {
      json.id = id
    }
    if (json && pr) {
      const newPropertySettingsObj = (useBgFill ? {...pr, useBgFill: true } : pr)
      // console.log('(00)---getProperty:--json.propertySettings[', json.propertySettings, 'newPropertySettingsObj:', newPropertySettingsObj)
      const oriPropertySettings = getTextByPathList(json, ['propertySettings', 'a:effectLst'])
      if (oriPropertySettings) {
        !newPropertySettingsObj['a:effectLst'] ? newPropertySettingsObj['a:effectLst'] = json.propertySettings['a:effectLst']
          : newPropertySettingsObj['a:effectLst'] = {...json.propertySettings['a:effectLst'], ...newPropertySettingsObj['a:effectLst']}
        // json.propertySettings = {...json.propertySettings}
      } 
      json.propertySettings = newPropertySettingsObj
      // console.log('(00)---getProperty:--json.propertySettings1[', json.propertySettings)
    }
  }
  return json
}


async function processMathNode(node, warpObj, source) {
  const choice = getTextByPathList(node, ['mc:Choice'])
  const fallback = getTextByPathList(node, ['mc:Fallback'])

  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(choice, ['p:sp', 'p:spPr', 'a:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const oMath = findOMath(choice)[0]
  const latex = latexFormart(parseOMath(oMath))

  const blipFill = getTextByPathList(fallback, ['p:sp', 'p:spPr', 'a:blipFill'])
  const picBase64 = await getPicFill(source, blipFill, warpObj)

  let text = ''
  if (getTextByPathList(choice, ['p:sp', 'p:txBody', 'a:p', 'a:r'])) {
    const sp = getTextByPathList(choice, ['p:sp'])
    text = genTextBody(sp['p:txBody'], sp, undefined, undefined, undefined, warpObj)
  }

  let effectData
  console.log('(00)-ppt-math-el:mathElement-[node]:', node) 
  const prNode = getTextByPathList(node, ['mc:Choice', 'p:sp'])
  console.log('(00)-ppt-math-el:mathElement-[prNode]:', prNode) 
  const propertySettings = {}
  {
    const data = {}
    // const propertySettings = {}
    let offectObj = {}
    // const defaultStyleObj = getTextByPathList(prNode, ['p:style'])
    // const defaultEffectRef = getTextByPathList(prNode, ['p:style', 'a:effectRef'])

    let shadow
    const outerShdwNode = getTextByPathList(prNode, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
    if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

    let glow
    const glowNode = getTextByPathList(prNode, ['p:spPr', 'a:effectLst', 'a:glow'])
    if (glowNode) glow = getGlow(glowNode, warpObj)

    // 获取倒影配置
    let reflection
    const reflectionNode = getTextByPathList(prNode, ['p:spPr', 'a:effectLst', 'a:reflection'])
    const effectLstNode = getTextByPathList(prNode, ['p:spPr', 'a:effectLst'])
    console.log('(00)-ppt-math-el:mathElement-[reflectionNode, effectLstNode]:', reflectionNode, effectLstNode) 
    if (effectLstNode && !reflectionNode) {
    //
    }
    else if (!effectLstNode && !reflectionNode) {
      const effectStyleLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:effectStyleLst']
      const idx = getTextByPathList(prNode, ['p:style', 'a:effectRef', 'attrs', 'idx'])
      const effectStyleList = getTextByPathList(effectStyleLst, ['a:effectStyle'])
      const lnIdx = Number(idx) - 1
      if (lnIdx >= 0) {
        const targetEffect = effectStyleList[Number(lnIdx)]
        // const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst', 'a:reflection'])
        const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst'])
        offectObj = {...offectObj, ...defaultReflection}
        if (targetEffect) {
          reflection = defaultReflection
        }
      }
    }
    else if (reflectionNode) {
      reflection = reflectionNode
      reflection
    }
    if (reflection) {
      propertySettings['a:effectLst'] = offectObj
    }
    console.log('(00)-ppt-math-el:mathElement-[propertySettings]:', propertySettings) 

    let softEdge
    const softEdgeNode = getTextByPathList(prNode, ['p:spPr', 'a:effectLst', 'a:softEdge'])
    if (softEdgeNode) softEdge = getSoftEdge(softEdgeNode)

    // const vAlign = getVerticalAlign(prNode, slideLayoutSpNode, slideMasterSpNode, type)

    // console.log('(00)-pptxtosjson-bodyPrValueAttrs:', bodyPrValueAttrs)
    // const vertValue = getTextByPathList(prNode, ['p:txBody', 'a:bodyPr', 'attrs', 'vert'])
    // let textDirectionValue
    // console.log('(00)-pptxtosjson-bodyPrValueAttrs:', vertValue, anchorValue, anchorCtrValue)
    // vertValue ? textDirectionValue = vertValue : ''
    // const vertAttrs = getTextByPathList(prNode, ['p:txBody', 'a:bodyPr', 'attrs'])
    // const isVertical = getTextByPathList(prNode, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
    if (shadow) data.shadow = shadow
    if (glow) data.glow = glow
    if (softEdge) data.softEdge = softEdge
    // if (autoFit) data.autoFit = autoFit
    // if (link) data.link = link
    console.log('(00)-ppt-math-el:data:', data)
    effectData = data
  }
  const { borderColor, borderWidth, borderType, strokeDasharray, borderColorObj } = getBorder(prNode, undefined, warpObj)
  strokeDasharray
  const spPr = getTextByPathList(prNode, ['p:spPr'])
  console.log('(00)-ppt-math-el:mathElement-[spPr]:', spPr)
  console.log('(00)-ppt-math-el:mathElement-[borderColor, borderWidth, borderType, borderColorObj]:', borderColor, borderWidth, borderType, borderColorObj)
  console.log('(00)-ppt-math-el:mathElement-[borderWidth]:', borderWidth)
  console.log('(00)-ppt-math-el:mathElement-[effectData]:', effectData)

  let heightMultiple = 1
  
  if (propertySettings) {
    const usePropertySettings = JSON.parse(JSON.stringify(propertySettings))
    // const useEffectData = JSON.parse(JSON.stringify(effectData))
    const effectLst = usePropertySettings['a:effectLst']
    // const glow = useEffectData['glow']
    // console.log('(00)-ppt-math-el:mathElement-[propertySettings]:', propertySettings)
    // console.log('(00)-ppt-math-el:mathElement-[effectLst]:', effectLst)
    if (effectLst) {
      heightMultiple = 2
    }
  }
  // console.log('(00)-ppt-math-el:mathElement-[heightMultiple]:', heightMultiple)

  return {
    type: 'math',
    propertySettings,
    top,
    left,
    width, 
    height: height * heightMultiple,
    latex,
    picBase64,
    text,
    order,
  }
}

async function processGroupSpNode(node, warpObj, source, parentGroupHierarchy = []) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:grpSpPr', 'a:xfrm'])
  if (!xfrmNode) return null

  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * RATIO_EMUs_Points
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * RATIO_EMUs_Points
  const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * RATIO_EMUs_Points
  const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * RATIO_EMUs_Points
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * RATIO_EMUs_Points
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * RATIO_EMUs_Points
  const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * RATIO_EMUs_Points
  const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * RATIO_EMUs_Points

  const isFlipV = getTextByPathList(xfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(xfrmNode, ['attrs', 'flipH']) === '1'

  let rotate = getTextByPathList(xfrmNode, ['attrs', 'rot']) || 0
  if (rotate) rotate = angleToDegrees(rotate)

  // 计算缩放因子
  const ws = cx / chcx
  const hs = cy / chcy

  // 构建当前组合层级（将当前组合添加到父级层级中）
  const currentGroupHierarchy = [...parentGroupHierarchy, node]

  const elements = []
  const parentGroupInfo = {
    x,
    y,
    chx,
    chy,
    cx,
    cy,
    chcx,
    chcy,
    ws,
    hs
  }
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        warpObj = {...warpObj, isInGroup: true}
        const ret = await processNodesInSlide(nodeKey, {...item, isInGroup: true, groupDeep: node.groupDeep ? node.groupDeep + 1 : 1, parentGroupInfo}, warpObj, source, currentGroupHierarchy)
        if (ret) elements.push(ret)
      }
    }
    else {
      warpObj = {...warpObj, isInGroup: true}
      const ret = await processNodesInSlide(nodeKey, {...node[nodeKey], isInGroup: true, groupDeep: node.groupDeep ? node.groupDeep + 1 : 1, parentGroupInfo}, warpObj, source, currentGroupHierarchy)
      if (ret) elements.push(ret)
    }
  }

  const processedElements = elements.map(element => {

    let left = numberToFixed((element.left - chx) * ws)
    let top = numberToFixed((element.top - chy) * hs)
    
    if (node.groupDeep === 1) {
      if (element.type === 'shape' && element.shapType === 'line') {
        left = numberToFixed((element.left - chx) + 15 )
        top = numberToFixed((element.top - chy) + 20 ) 
      }
      else if (element.type === 'text') {
        // left = numberToFixed((element.left - chx) + 0)
        top = numberToFixed((element.top - chy) + 10) 
      }
      else {
        const parObj = node.parentGroupInfo
        // left = numberToFixed((left - parObj.chx) * parObj.ws)
        // top = numberToFixed((top - parObj.chy) * parObj.hs)
        // left = numberToFixed((left * parObj.cx * 1.11))
        // top = numberToFixed((top * parObj.cy * 1.89))
        // let debugValue = 0
        // if (element.id === '49') {
        //   debugValue = 0
        // }
        // let guessValue = 3
        // let topDiff = 0
        // if (element.type === 'text' || (element.type === 'shape' && element.shapType === 'line')) {
        //   guessValue = 5
        //   topDiff = 8
        // }
        // left = numberToFixed((left * parObj.cx * cy * guessValue / (cx + cy)) + debugValue)
        // top = numberToFixed((top * parObj.cy * cx * guessValue / (cx + cy)) + topDiff)
      
        const x1 = (element.left - chx) * ws
        const y1 = (element.top - chy) * hs

        left = x1 * parObj.ws
        top = y1 * parObj.hs
      }

    }
    // else if (!node.groupDeep) {
    //   const parObj = node.parentGroupInfo
    //   const guessValue = element.type === 'text' || (element.type === 'shape' && element.shapType === 'line') ? 5 : 3
    //   left = numberToFixed((left * parObj.cx * cy * guessValue / (cx + cy)))
    //   top = numberToFixed((top * parObj.cy * cx * guessValue / (cx + cy)))
    // }

    if (element.borderWidth) {
      left = left - element.borderWidth / 2
      // top = top - element.borderWidth
    }
    return {
      ...element,
      elementInGroupDeep: node.groupDeep,
      left: left,
      top: top,
      width: numberToFixed(element.width * ws),
      height: numberToFixed(element.height * hs),
      ...(element.type === 'group' && element.elements ? {
        elements: processNestedGroupElements(element.elements, ws, hs)
      } : {})
    }
  })
  // const processedElements = elements.map(element => ({
  //   ...element,
  //   // left: numberToFixed((element.left - chx) * ws),
  //   // top: numberToFixed((element.top - chy) * hs),
  //   left: node.groupDeep === 1 ? numberToFixed((element.left - chx) + (element.type === 'shape' || element.shapType === 'line' ? 15 : 0)) : numberToFixed((element.left - chx) * ws),
  //   top: node.groupDeep === 1 ? numberToFixed((element.top - chy) + (element.type === 'shape' || element.shapType === 'line' ? 20 : 10)) : numberToFixed((element.top - chy) * hs),
  //   width: numberToFixed(element.width * ws),
  //   height: numberToFixed(element.height * hs),
  //   ...(element.type === 'group' && element.elements ? {
  //     elements: processNestedGroupElements(element.elements, ws, hs)
  //   } : {})
  // }))

  function processNestedGroupElements(elements, ws, hs, depth = 0) {
    if (depth > 0) return elements

    return elements.map(element => {
      const processed = {
        ...element,
        width: numberToFixed(element.width * ws),
        height: numberToFixed(element.height * hs),
      }
      if (element.type === 'group' && element.elements) {
        processed.elements = processNestedGroupElements(element.elements, ws, hs, depth + 1)
      }
      return processed
    })
  }

  return {
    type: 'group',
    top: numberToFixed(y),
    left: numberToFixed(x),
    width: numberToFixed(cx),
    height: numberToFixed(cy),
    rotate,
    order,
    isFlipV,
    isFlipH,
    elements: processedElements,
  }
}

async function processSpNode(node, warpObj, source, groupHierarchy = []) {
  const name = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name'])
  const id = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id'])
  const cNvPr = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr'])
  // const name = getTextByPathList(cNvPr, ['attrs', 'name'])
  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  let type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  const order = getTextByPathList(node, ['attrs', 'order'])

  let slideLayoutSpNode, slideMasterSpNode

  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
      slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
      if (!slideLayoutSpNode) slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      if (!slideMasterSpNode) slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
    else {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
  }
  else if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
    slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
  }

  if (!type) {
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') type = 'text'
  }
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  if (!type) {
    if (source === 'diagramBg') type = 'diagram'
    else type = 'obj'
  }

  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)

  return await genShape(node, slideLayoutSpNode, slideMasterSpNode, name, id, type, order, warpObj, source, link, groupHierarchy, id)
}

async function processCxnSpNode(node, warpObj, source) {
  const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']
  const id = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['id']
  const cNvPr = getTextByPathList(node, ['p:nvCxnSpPr', 'p:cNvPr'])
  // const name = getTextByPathList(cNvPr, ['attrs', 'name'])
  const type = (node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined) ? undefined : node['p:nvCxnSpPr']['p:nvPr']['p:ph']['attrs']['type']
  const order = node['attrs']['order']
  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)

  return await genShape(node, undefined, undefined, name, id, type, order, warpObj, source, link)
}

async function genShape(node, slideLayoutSpNode, slideMasterSpNode, name, id, type, order, warpObj, source, link, groupHierarchy = []) {
  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  const keypoints = {}
  if (shapType) {
    const shapAdjst_ary = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'a:avLst', 'a:gd'])
    if (shapAdjst_ary) {
      const adjList = Array.isArray(shapAdjst_ary) ? shapAdjst_ary : [shapAdjst_ary]
      for (const adj of adjList) {
        const name = getTextByPathList(adj, ['attrs', 'name'])
        const fmla = getTextByPathList(adj, ['attrs', 'fmla'])
        if (name && fmla && fmla.startsWith('val ')) {
          keypoints[name] = parseInt(fmla.substring(4)) / 50000
        }
      }
    }
  }

  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const isFlipV = getTextByPathList(slideXfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(slideXfrmNode, ['attrs', 'flipH']) === '1'

  const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = angleToDegrees(txtXframeRot) + 90
  } 
  else txtRotate = rotate

  let content = ''
  let fullText = []
  // if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj, width, height, true)
  const { borderColor, borderWidth, borderColorObj, borderType, strokeDasharray } = getBorder(node, type, warpObj)
  const fill = await getShapeFill(node, warpObj, source, groupHierarchy)
  if (fill && fill.type === 'gradient') dealGradientFill(fill)

  const isShape = (custShapType && type !== 'diagram') 
                  || (shapType && (type === 'obj' || !type || shapType !== 'rect')) 
                  || (shapType && (fill || borderWidth))
                  // || (shapType && !isHasValidText && (fill || borderWidth))
  const bodyPrValueAttrs = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs'])
  const anchorCtrValue = getTextByPathList(bodyPrValueAttrs, ['anchorCtr'])
  const anchorValue = getTextByPathList(bodyPrValueAttrs, ['anchor'])
  const anchorInfo = {
    anchorValue,
    anchorCtrValue
  }
  const isVertical1 = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
  // console.log('(00)-pptxtosjson-bodyPrValueAttrs:', bodyPrValueAttrs)
  const vertValue = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert'])
  let textDirectionValue
  console.log('(00)-pptxtosjson-bodyPrValueAttrs:', vertValue, anchorValue, anchorCtrValue)
  vertValue ? textDirectionValue = vertValue : ''
  const paramsObj = {
    isVertical: isVertical1,
    textDirectionValue,
    isUseNewDeal: !isShape,
    width,
    height,
  }
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj, paramsObj)
  fullText = []
  const propertySettings = {}
  let offectObj = {}
  // const defaultStyleObj = getTextByPathList(node, ['p:style'])
  // const defaultEffectRef = getTextByPathList(node, ['p:style', 'a:effectRef'])

  let shadow
  const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
  if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

  let glow
  const glowNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:glow'])
  if (glowNode) glow = getGlow(glowNode, warpObj)

  // 获取倒影配置
  let reflection
  const reflectionNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:reflection'])
  const effectLstNode = getTextByPathList(node, ['p:spPr', 'a:effectLst'])
  if (effectLstNode && !reflectionNode) {
    //
  }
  else if (!effectLstNode && !reflectionNode) {
    const effectStyleLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:effectStyleLst']
    const idx = getTextByPathList(node, ['p:style', 'a:effectRef', 'attrs', 'idx'])
    const effectStyleList = getTextByPathList(effectStyleLst, ['a:effectStyle'])
    const lnIdx = Number(idx) - 1
    if (lnIdx >= 0) {
      const targetEffect = effectStyleList[Number(lnIdx)]
      // const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst', 'a:reflection'])
      const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst'])
      offectObj = {...offectObj, ...defaultReflection}
      if (targetEffect) {
        reflection = defaultReflection
      }
    }
  }
  else if (reflectionNode) {
    reflection = reflectionNode
    reflection
  }
  if (reflection) {
    propertySettings['a:effectLst'] = offectObj
  }

  let softEdge
  const softEdgeNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:softEdge'])
  if (softEdgeNode) softEdge = getSoftEdge(softEdgeNode)

  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)

  // const vertAttrs = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs'])
  const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
  const autoFit = getTextAutoFit(node, slideLayoutSpNode, slideMasterSpNode)

  const data = {
    left,
    top,
    width,
    height,
    borderColor,
    borderColorObj,
    propertySettings,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fill,
    content,
    fullText,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    name,
    id,
    order,
  }

  if (shadow) data.shadow = shadow
  if (glow) data.glow = glow
  if (softEdge) data.softEdge = softEdge
  if (autoFit) data.autoFit = autoFit
  if (link) data.link = link

  const isHasValidText = data.content && hasValidText(data.content)

  if (custShapType && type !== 'diagram') {
    const ext = getTextByPathList(slideXfrmNode, ['a:ext', 'attrs'])
    const w = parseInt(ext['cx']) * RATIO_EMUs_Points
    const h = parseInt(ext['cy']) * RATIO_EMUs_Points
    const d = getCustomShapePath(custShapType, w, h)
    if (!isHasValidText) data.content = ''

    return {
      ...data,
      type: 'shape',
      shapType: 'custom',
      path: d,
    }
  }

  let shapePath = ''
  if (shapType) shapePath = getShapePath(shapType, width, height, node)
  if (shapType && (type === 'obj' || !type || shapType !== 'rect')) {
    if (!isHasValidText) data.content = ''
    return {
      ...data,
      type: 'shape',
      shapType,
      path: shapePath,
      keypoints,
    }
  }
  if (shapType && !isHasValidText && (fill || borderWidth)) {
    return {
      ...data,
      type: 'shape',
      content: '',
      shapType,
      path: shapePath,
      keypoints,
    }
  }
  return {
    ...data,
    type: 'text',
    isVertical,
    textDirectionValue,
    anchorInfo,
    rotate: txtRotate,
  }
}

async function processPicNode(node, warpObj, source) {
  let resObj
  if (source === 'slideMasterBg') resObj = warpObj['masterResObj']
  else if (source === 'slideLayoutBg') resObj = warpObj['layoutResObj']
  else resObj = warpObj['slideResObj']

  const cNvPr = getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr'])
  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)
  const order = node['attrs']['order']
  
  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  // const targetInfo = getTextByPathList(node, ['p:blipFill', 'a:blip'])
  // const targetInfo1 = getTextByPathList(node, ['p:blipFill', 'a:blip', 'a:alphaModFix'])
  // const targetInfo2 = getTextByPathList(node, ['p:blipFill', 'a:blip', 'a:lum'])

  const fill = await getShapeFill(node, warpObj, source, [])
  if (fill && fill.type === 'gradient') dealGradientFill(fill)
  console.log('(00)-ppt-image-el:ImageElement-[fill]:', fill)


  const opacity = getPicFillOpacity( node['p:blipFill'])
  // const targetInfo1 = getTextByPathList(node, ['p:blipFill', 'a:blip', 'a:alphaModFix', 'attrs', 'amt'])
  const contrastInfo = getTextByPathList(node, ['p:blipFill', 'a:blip', 'a:lum', 'attrs', 'contrast'])
  let contrast = 1
  if (contrastInfo) {
    // contrast = parseInt(contrastInfo) / 100000
    contrast = parseInt(contrastInfo) / 10000
  }
  // console.log('(00)-ppt-image-el:node:-targetInfo:', opacity, contrast, targetInfo1, targetInfo2)
  const blipFillPr = {
    contrast,
    opacity
  }
  const spPrNode = getTextByPathList(node, ['p:spPr'])
  console.log('(00)-ppt-image-el:glowNode:', spPrNode)
  console.log('(00)-ppt-image-el:node:', node)
  
  console.log('(00)-ppt-image-el:3DInfo-[p:spPr]:', spPrNode)

  const sp3dNode = getTextByPathList(node, ['p:spPr', 'a:sp3d'])
  const sp3dAttrs = getTextByPathList(node, ['p:spPr', 'a:sp3d', 'attrs'])
  const sp3dContourNode = getTextByPathList(node, ['p:spPr', 'a:sp3d', 'a:contourClr'])
  const sp3dExtrusionNode = getTextByPathList(node, ['p:spPr', 'a:sp3d', 'a:extrusionClr'])
  const scene3dNode = getTextByPathList(node, ['p:spPr', 'a:scene3d'])
  // getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)

  console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]:', sp3dNode)
  console.log('(00)-ppt-image-el:3DInfo-[a:scene3d]:', scene3dNode)
  
  // let sp3dPr = {}
  let extrusionHValue
  let contourWValue
  let sp3dExtrusion
  let sp3dContour
  let sp3dPrstMaterial
  console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[sp3dAttrs]:', sp3dAttrs)
  if (sp3dAttrs) {
    const {extrusionH, contourW, prstMaterial} = sp3dAttrs
    console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[extrusionH, contourW, prstMaterial]:', extrusionH, contourW, prstMaterial)
    if (extrusionH) extrusionHValue = numberToFixed(parseInt(extrusionH) * RATIO_EMUs_Points)
    if (contourW) contourWValue = numberToFixed(parseInt(contourW) * RATIO_EMUs_Points)
    sp3dPrstMaterial = prstMaterial 
    console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[extrusionHValue, contourWValue]:', extrusionHValue, contourWValue)
  }

  console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[sp3dContourNode]:', sp3dContourNode)
  console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[sp3dExtrusionNode]:', sp3dExtrusionNode)

  // extrusionH="76200" contourW="63500" prstMaterial="dkEdge"

  if (sp3dContourNode) {
    sp3dContour = getSolidFill(sp3dContourNode, undefined, undefined, warpObj)
  }
  if (sp3dExtrusionNode) {
    sp3dExtrusion = getSolidFill(sp3dExtrusionNode, undefined, undefined, warpObj)
  }
  const sp3dObj = {
    extrusionHValue,
    contourWValue,
    sp3dExtrusion,
    sp3dContour,
    sp3dPrstMaterial,
  }
  const sp3dObj1 = {
    extrusion: {
      color: sp3dExtrusion,
      height: extrusionHValue
    },
    contour: {
      color: sp3dContour,
      width: contourWValue
    }
  }
  console.log('(00)-ppt-image-el:3DInfo-[a:sp3d]-[sp3dObj,sp3dObj1]:', sp3dObj, sp3dObj1)
  
  let effectData
  {
    const data = {}
    const propertySettings = {}
    let offectObj = {}
    // const defaultStyleObj = getTextByPathList(node, ['p:style'])
    // const defaultEffectRef = getTextByPathList(node, ['p:style', 'a:effectRef'])

    let shadow
    const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
    if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

    let glow
    const glowNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:glow'])
    if (glowNode) glow = getGlow(glowNode, warpObj)

    // 获取倒影配置
    let reflection
    const reflectionNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:reflection'])
    const effectLstNode = getTextByPathList(node, ['p:spPr', 'a:effectLst'])
    if (effectLstNode && !reflectionNode) {
    //
    }
    else if (!effectLstNode && !reflectionNode) {
      const effectStyleLst = warpObj['themeContent']['a:theme']['a:themeElements']['a:fmtScheme']['a:effectStyleLst']
      const idx = getTextByPathList(node, ['p:style', 'a:effectRef', 'attrs', 'idx'])
      const effectStyleList = getTextByPathList(effectStyleLst, ['a:effectStyle'])
      const lnIdx = Number(idx) - 1
      if (lnIdx >= 0) {
        const targetEffect = effectStyleList[Number(lnIdx)]
        // const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst', 'a:reflection'])
        const defaultReflection = getTextByPathList(targetEffect, ['a:effectLst'])
        offectObj = {...offectObj, ...defaultReflection}
        if (targetEffect) {
          reflection = defaultReflection
        }
      }
    }
    else if (reflectionNode) {
      reflection = reflectionNode
      reflection
    }
    if (reflection) {
      propertySettings['a:effectLst'] = offectObj
    }

    let softEdge
    const softEdgeNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:softEdge'])
    if (softEdgeNode) softEdge = getSoftEdge(softEdgeNode)

    // const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)

    // console.log('(00)-pptxtosjson-bodyPrValueAttrs:', bodyPrValueAttrs)
    // const vertValue = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert'])
    // let textDirectionValue
    // console.log('(00)-pptxtosjson-bodyPrValueAttrs:', vertValue, anchorValue, anchorCtrValue)
    // vertValue ? textDirectionValue = vertValue : ''
    // const vertAttrs = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs'])
    // const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
    if (shadow) data.shadow = shadow
    if (glow) data.glow = glow
    if (softEdge) data.softEdge = softEdge
    // if (autoFit) data.autoFit = autoFit
    if (link) data.link = link
    console.log('(00)-ppt-image-el:data:', data)
    effectData = data
  }
  const { borderColor, borderWidth, borderType, strokeDasharray, borderColorObj } = getBorder(node, undefined, warpObj)
  const spPr = getTextByPathList(node, ['p:spPr'])
  console.log('(00)-ppt-image-el:ImageElement-[spPr]:', spPr)
  console.log('(00)-ppt-image-el:ImageElement-[borderColor, borderWidth, borderType, borderColorObj]:', borderColor, borderWidth, borderType, borderColorObj)
  console.log('(00)-ppt-image-el:ImageElement-[borderWidth]:', borderWidth)

  const imgName = resObj[rid]['target']
  const imgFileExt = extractFileExtension(imgName).toLowerCase()
  const zip = warpObj['zip']
  const imgArrayBuffer = await zip.file(imgName).async('arraybuffer')

  let xfrmNode = node['p:spPr']['a:xfrm']
  if (!xfrmNode) {
    const idx = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
    if (idx) xfrmNode = getTextByPathList(warpObj['slideLayoutTables'], ['idxTable', idx, 'p:spPr', 'a:xfrm'])
  }

  const mimeType = getMimeType(imgFileExt)
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  const src = `data:${mimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`

  const isFlipV = getTextByPathList(xfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(xfrmNode, ['attrs', 'flipH']) === '1'

  let rotate = 0
  const rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
  if (rotateNode) rotate = angleToDegrees(rotateNode)

  const videoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
  let videoRid, videoFile, videoFileExt, videoMimeType, uInt8ArrayVideo, videoBlob
  let isVdeoLink = false

  if (videoNode) {
    videoRid = videoNode['attrs']['r:link']
    videoFile = resObj[videoRid]['target']
    if (isVideoLink(videoFile)) {
      videoFile = escapeHtml(videoFile)
      isVdeoLink = true
    } 
    else {
      videoFileExt = extractFileExtension(videoFile).toLowerCase()
      if (videoFileExt === 'mp4' || videoFileExt === 'webm' || videoFileExt === 'ogg') {
        uInt8ArrayVideo = await zip.file(videoFile).async('arraybuffer')
        videoMimeType = getMimeType(videoFileExt)
        videoBlob = URL.createObjectURL(new Blob([uInt8ArrayVideo], {
          type: videoMimeType
        }))
      }
    }
  }

  const audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
  let audioRid, audioFile, audioFileExt, uInt8ArrayAudio, audioBlob
  if (audioNode) {
    audioRid = audioNode['attrs']['r:link']
    audioFile = resObj[audioRid]['target']
    audioFileExt = extractFileExtension(audioFile).toLowerCase()
    if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
      uInt8ArrayAudio = await zip.file(audioFile).async('arraybuffer')
      audioBlob = URL.createObjectURL(new Blob([uInt8ArrayAudio]))
    }
  }

  if (videoNode && !isVdeoLink) {
    return {
      type: 'video',
      blipFillPr,
      fill,
      borderColor,
      borderWidth,
      borderType,
      borderColorObj,
      ...effectData,
      top,
      left,
      width, 
      height,
      rotate,
      blob: videoBlob,
      order,
    }
  } 
  if (videoNode && isVdeoLink) {
    return {
      type: 'video',
      blipFillPr,
      fill,
      borderColor,
      borderWidth,
      borderType,
      borderColorObj,
      ...effectData,
      top,
      left,
      width, 
      height,
      rotate,
      src: videoFile,
      order,
    }
  }
  if (audioNode) {
    return {
      type: 'audio',
      top,
      left,
      width, 
      height,
      rotate,
      blob: audioBlob,
      order,
    }
  }

  let rect
  const srcRectAttrs = getTextByPathList(node, ['p:blipFill', 'a:srcRect', 'attrs'])
  if (srcRectAttrs && (srcRectAttrs.t || srcRectAttrs.b || srcRectAttrs.l || srcRectAttrs.r)) {
    rect = {}
    if (srcRectAttrs.t) rect.t = srcRectAttrs.t / 1000
    if (srcRectAttrs.b) rect.b = srcRectAttrs.b / 1000
    if (srcRectAttrs.l) rect.l = srcRectAttrs.l / 1000
    if (srcRectAttrs.r) rect.r = srcRectAttrs.r / 1000
  }
  let geom = 'rect'
  const prstGeom = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custGeom = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  if (prstGeom) {
    geom = prstGeom
  }
  else if (custGeom) {
    geom = identifyShape(custGeom)
    if (geom !== 'custom') geom = `custom:${geom}`
  }

  // const { borderColor, borderWidth, borderType, strokeDasharray, borderColorObj } = getBorder(node, undefined, warpObj)
  // const spPr = getTextByPathList(node, ['p:spPr'])
  // console.log('(00)-ppt-image-el:ImageElement-[spPr]:', spPr)
  // console.log('(00)-ppt-image-el:ImageElement-[borderColor, borderWidth, borderType, borderColorObj]:', borderColor, borderWidth, borderType, borderColorObj)

  const filters = getPicFilters(node['p:blipFill'])

  const imageData = {
    type: 'image',
    blipFillPr,
    fill,
    borderColorObj,
    ...effectData,
    top,
    left,
    width,
    height,
    rotate,
    src,
    isFlipV,
    isFlipH,
    order,
    rect,
    geom,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
  }
  console.log('(00)-ppt-image-el:ImageElement-[imageData]:', imageData)


  if (filters) imageData.filters = filters
  if (link) imageData.link = link
  // console.log('(00)-ppt-image-el:imageData:', imageData)
  return imageData
}

async function processGraphicFrameNode(node, warpObj, source) {
  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])
  
  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = await genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj, source)
      // console.log('(00)-pptxtojson-[chartEL]:-graphicTypeUri:', graphicTypeUri)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = await genDiagram(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
      let oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj'])
      if (!oleObjNode) oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj'])
      if (oleObjNode) result = await processGroupSpNode(oleObjNode, warpObj, source)
      break
    default:
  }
  return result
}

async function genTable(node, warpObj) {
  const order = node['attrs']['order']
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const getTblPr = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr'])
  let getColsGrid = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblGrid', 'a:gridCol'])
  const getTblGrid = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblGrid'])
  console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]- width, height:', width, height)
  console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]-getTblGrid:', getTblGrid)
  console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]-getColsGrid:', getColsGrid)
  if (getColsGrid.constructor !== Array) getColsGrid = [getColsGrid]

  const colWidths = []
  if (getColsGrid) {
    for (const item of getColsGrid) {
      const colWidthParam = getTextByPathList(item, ['attrs', 'w']) || 0
      const colWidth = parseInt(colWidthParam) * RATIO_EMUs_Points
      colWidths.push(colWidth)
    }
  }
  console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]-colWidths:', colWidths)

  const firstRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstRow'] : undefined
  const firstColAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstCol'] : undefined
  const lastRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastRow'] : undefined
  const lastColAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastCol'] : undefined
  const bandRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandRow'] : undefined
  const bandColAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandCol'] : undefined
  const tblStylAttrObj = {
    isFrstRowAttr: (firstRowAttr && firstRowAttr === '1') ? 1 : 0,
    isFrstColAttr: (firstColAttr && firstColAttr === '1') ? 1 : 0,
    isLstRowAttr: (lastRowAttr && lastRowAttr === '1') ? 1 : 0,
    isLstColAttr: (lastColAttr && lastColAttr === '1') ? 1 : 0,
    isBandRowAttr: (bandRowAttr && bandRowAttr === '1') ? 1 : 0,
    isBandColAttr: (bandColAttr && bandColAttr === '1') ? 1 : 0,
  }

  let thisTblStyle
  const tbleStyleId = getTblPr['a:tableStyleId']
  if (tbleStyleId) {
    const tbleStylList = warpObj['tableStyles']['a:tblStyleLst']['a:tblStyle']
    if (tbleStylList) {
      if (tbleStylList.constructor === Array) {
        for (let k = 0; k < tbleStylList.length; k++) {
          if (tbleStylList[k]['attrs']['styleId'] === tbleStyleId) {
            thisTblStyle = tbleStylList[k]
          }
        }
      } 
      else {
        if (tbleStylList['attrs']['styleId'] === tbleStyleId) {
          thisTblStyle = tbleStylList
        }
      }
    }
  }
  if (thisTblStyle) thisTblStyle['tblStylAttrObj'] = tblStylAttrObj

  let borders = {}
  const tblStyl = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle'])
  const tblBorderStyl = getTextByPathList(tblStyl, ['a:tcBdr'])
  if (tblBorderStyl) borders = getTableBorders(tblBorderStyl, warpObj)

  let tbl_bgcolor = ''
  let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:tblBg', 'a:fillRef'])
  if (tbl_bgFillschemeClr) {
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (tbl_bgFillschemeClr === undefined) {
    tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }

  let trNodes = tableNode['a:tr']
  if (trNodes.constructor !== Array) trNodes = [trNodes]
  
  const data = []
  const rowHeights = []
  for (let i = 0; i < trNodes.length; i++) {
    const trNode = trNodes[i]
    
    const rowHeightParam = getTextByPathList(trNodes[i], ['attrs', 'h']) || 0
    const rowHeight = parseInt(rowHeightParam) * RATIO_EMUs_Points
    rowHeights.push(rowHeight)
    const curWidth = colWidths[i]
    console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]-curWidth,rowHeight:', curWidth, rowHeight)

    const {
      fillColor,
      fontColor,
      fontBold,
    } = getTableRowParams(trNodes, i, tblStylAttrObj, thisTblStyle, warpObj)

    const tcNodes = trNode['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (let j = 0; j < tcNodes.length; j++) {
        const tcNode = tcNodes[j]
        let a_sorce
        if (j === 0 && tblStylAttrObj['isFrstColAttr'] === 1) {
          a_sorce = 'a:firstCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:seCell'])) {
            a_sorce = 'a:seCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 &&
            getTextByPathList(thisTblStyle, ['a:neCell'])) {
            a_sorce = 'a:neCell'
          }
        } 
        else if (
          (j > 0 && tblStylAttrObj['isBandColAttr'] === 1) &&
          !(tblStylAttrObj['isFrstColAttr'] === 1 && i === 0) &&
          !(tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1)) &&
          j !== (tcNodes.length - 1)
        ) {
          if ((j % 2) !== 0) {
            let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
            if (aBandNode === undefined) {
              aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
              if (aBandNode) a_sorce = 'a:band2V'
            } 
            else a_sorce = 'a:band2V'
          }
        }
        if (j === (tcNodes.length - 1) && tblStylAttrObj['isLstColAttr'] === 1) {
          a_sorce = 'a:lastCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:swCell'])) {
            a_sorce = 'a:swCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 && getTextByPathList(thisTblStyle, ['a:nwCell'])) {
            a_sorce = 'a:nwCell'
          }
        }

        const tcPrNode = getTextByPathList(tcNode['a:tcPr'], ['attrs', 'vert'])
        const tcPrVertNode = getTextByPathList(tcNode['a:tcPr'], ['attrs', 'vert'])
        // const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
        const isVertical = tcPrVertNode ? tcPrVertNode.includes('vert') || tcPrVertNode.includes('Vert') : false

        const anchorInfo = getTextByPathList(tcNode['a:tcPr'], ['attrs', 'anchor'])
        // t、b、ctr
        // const anchorCheckList = ['t', 'b', 'ctr']
        const anchorDict = {'t': 'flex-start', 'b': 'flex-end', 'ctr': 'center'}
        const anchor = anchorInfo && anchorDict[`${anchorInfo}`] ? anchorDict[`${anchorInfo}`] : anchorDict['t']

        let textDirectionValue
        console.log('(00)-tableEl-text-[isVertical]:', isVertical)
        
        if (tcPrVertNode) {
          console.log('(00)-tableEl-text-[tcPrNode]:', tcPrNode)
          console.log('(00)-tableEl-text-[tcPrNode]-[tcPrVertNode]:', tcPrVertNode)
          textDirectionValue = tcPrVertNode
        }
        const paramsObj = {
          isVertical,
          anchor,
          textDirectionValue,
          isUseNewDeal: false || isVertical,
          width: curWidth,
          height: rowHeight,
        }
        const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, undefined, warpObj, paramsObj)
        // const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, undefined, warpObj)
        const cell = await getTableCellParams(tcNode, thisTblStyle, a_sorce, warpObj)
        const td = { text, isVertical, textDirectionValue, anchor }
        if (cell.rowSpan) td.rowSpan = cell.rowSpan
        if (cell.colSpan) td.colSpan = cell.colSpan
        if (cell.vMerge) td.vMerge = cell.vMerge
        if (cell.hMerge) td.hMerge = cell.hMerge
        if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
        if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
        if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
        if (cell.borders) td.borders = cell.borders
        if (cell.slashObj && JSON.stringify(cell.slashObj) !== '{}') td.slashObj = cell.slashObj
        console.log('(00)-tableEl-pptxtojson-anlysis-[td.fillColor]:', td.fillColor)

        tr.push(td)
        console.log('(00)-tableEl-text-[cellData]-td:', td)
      }
    } 
    else {
      let a_sorce
      if (tblStylAttrObj['isFrstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:firstCol'
      } 
      else if (tblStylAttrObj['isBandColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
        if (!aBandNode) {
          aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
          if (aBandNode) a_sorce = 'a:band2V'
        } 
        else a_sorce = 'a:band2V'
      }
      if (tblStylAttrObj['isLstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:lastCol'
      }

      const tcPrNode = getTextByPathList(tcNodes['a:tcPr'], ['attrs', 'vert'])
      const tcPrVertNode = getTextByPathList(tcNodes['a:tcPr'], ['attrs', 'vert'])
      const isVertical = tcPrVertNode ? tcPrVertNode.includes('vert') || tcPrVertNode.includes('Vert') : false
      let textDirectionValue
      // console.log('(00)-tableEl-text-[isVertical]:', isVertical)
      const anchorInfo = getTextByPathList(tcNodes['a:tcPr'], ['attrs', 'anchor'])
      // t、b、ctr
      // const anchorCheckList = ['t', 'b', 'ctr']
      const anchorDict = {'t': 'flex-start', 'b': 'flex-end', 'ctr': 'center'}
      const anchor = anchorInfo && anchorDict[`${anchorInfo}`] ? anchorDict[`${anchorInfo}`] : anchorDict['t']
      // console.log('(00)-tableEl-text-[spPr]---pptxtojson-[anchor]:', anchor, tcNodes['a:tcPr'])


        
      if (tcPrVertNode) {
        // console.log('(00)-tableEl-text-[tcPrNode]:', tcPrNode)
        // console.log('(00)-tableEl-text-[tcPrNode]-[tcPrVertNode]:', tcPrVertNode)
        textDirectionValue = tcPrVertNode
      }
      const paramsObj = {
        isVertical,
        anchor,
        textDirectionValue,
        isUseNewDeal: false || isVertical,
        width: curWidth,
        height: rowHeight,
        type: 'table',
      }
      const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, undefined, warpObj, paramsObj)
      // const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, undefined, warpObj)

      const cell = await getTableCellParams(tcNodes, thisTblStyle, a_sorce, warpObj)
      console.log('(00)-tableEl-text-[useNewDeal]-[parsePPTTextToLines]-[ok]--cell:', cell)
      const td = { text, isVertical, textDirectionValue, anchor }
      if (cell.rowSpan) td.rowSpan = cell.rowSpan
      if (cell.colSpan) td.colSpan = cell.colSpan
      if (cell.vMerge) td.vMerge = cell.vMerge
      if (cell.hMerge) td.hMerge = cell.hMerge
      if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
      if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
      if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
      if (cell.borders) td.borders = cell.borders
      console.log('(00)-tableEl-pptxtojson-anlysis-[td.fillColor]:', td.fillColor)

      tr.push(td)
      console.log('(00)-tableEl-text-[cellData]-td_in_tcNodes:', td)
    }
    data.push(tr)
    console.log('(00)-tableEl-text-[cellData]-tr:', tr)
  }
  console.log('(00)-tableEl-text-[cellData]-data:', data)

  let actualTableWidth = colWidths.reduce((sum, width) => sum + width, 0)
  if (actualTableWidth) actualTableWidth = numberToFixed(actualTableWidth)

  return {
    type: 'table',
    top,
    left,
    width: actualTableWidth || width,
    height,
    data,
    order,
    borders,
    rowHeights,
    colWidths,
  }
}

async function genChart(node, warpObj, source) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  // console.log('(00)-pptxtojson-[chartEL]:-genChart-[rid]:', rid)
  let refName = getTextByPathList(warpObj['slideResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['layoutResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['masterResObj'], [rid, 'target'])
  if (!refName) return {}

  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  const type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  let slideLayoutSpNode, slideMasterSpNode

  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
      slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
      if (!slideLayoutSpNode) slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      if (!slideMasterSpNode) slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
    else {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
  }
  else if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
    slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
  }
  let orithemeElements = null
  if (refName) {
    const themeName = refName.split('/').pop()
    const dealSamePathFilePath = (target) => {
      if (!target.includes('/')) {
        return refName.replace(themeName, '') + target
      }
      return target
    }
    const getXMlFileContent = async (target) => {
      // console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj]:-getXMlFileContent-target:', target)
      const useTarget = dealSamePathFilePath(target)
      // console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj]:-getXMlFileContent-useTarget:', useTarget)
      if (!useTarget.endsWith('.xml')) return null
      const content = await readXmlFile(warpObj['zip'], useTarget)
      return content
    }
    const themeResFileName = refName.replace(themeName, '_rels/' + themeName) + '.rels'
    const themeResContent = await readXmlFile(warpObj['zip'], themeResFileName)
    const themeResObj = {}
    if (themeResContent) {
      let relationshipArray = themeResContent['Relationships']['Relationship']
      if (relationshipArray) {
        if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
        for (const relationshipArrayItem of relationshipArray) {
          themeResObj[relationshipArrayItem['attrs']['Id']] = {
            'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': dealSamePathFilePath(relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')),
            'targetContent': await getXMlFileContent(relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'))
          }
        }
      }
    }

    warpObj.themeResObj = themeResObj
    // 获取图表主题
    // const themeOverrideContent = Object.values(themeResObj).filter(item => item.type === 'themeOverride')
    const themeOverrideContent = Object.values(themeResObj).find(item => item.type === 'themeOverride')
    const contentObj = getTextByPathList(themeOverrideContent, ['targetContent', 'a:themeOverride'])
    if (contentObj && getTextByPathList(warpObj, ['themeContent', 'a:theme', 'a:themeElements'])) {
      const themeElements = warpObj['themeContent']['a:theme']['a:themeElements']
      orithemeElements = JSON.parse(JSON.stringify(themeElements))
      warpObj['themeContent']['a:theme']['a:themeElements'] = {
        ...warpObj['themeContent']['a:theme']['a:themeElements'],
        ...contentObj
      }
    }
    // for (const key in plotArea) {

    // }
  }

  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])
  const otherParams = {
    slideLayoutSpNode, 
    slideMasterSpNode
  }
  const chart = getChartInfo(plotArea, warpObj, source, otherParams)

  if (!chart) return {}
  const plotTitle = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:title'])
  const chartTitleInfo = getChartTitle(plotTitle, warpObj, source)
  const plotLegend = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:legend'])
  const blipFill = getTextByPathList(content, ['c:chartSpace', 'c:spPr', 'a:blipFill'])
  const picBase64 = await getPicFill('themeBg', blipFill, warpObj)
  const background = {type: 'image', src: picBase64}

  const spPrNode = getTextByPathList(content, ['c:chartSpace', 'c:spPr'])
  const prNode = await getChartElPrInfo(spPrNode, warpObj, source)
  
  const importProperty = {
    background,
    ...prNode,
  }
  console.log('(00)-devAnalysisPPT-[chartEL]:-genChart--getChartInfo-[chart]:-deal-prNode:', spPrNode, prNode, importProperty)
  console.log('(00)-devAnalysisPPT-[chartEL]:-showChart-props.elementInfo:-chart.key:', chart.key)

  const chartSpPr = {
    plotAreaSpPr: chart ? chart.plotAreaSpPrNode : {},
    chartTypeInfo: {
      type: chart.type,
      typeKey: chart.key,
      barDir: chart.barDir,
      grouping: chart.grouping,
    },
    ...chart.chartPr
  }
  const { borderColor, borderWidth, borderType, strokeDasharray, borderColorObj } = getBorder(spPrNode, undefined, warpObj)
  const borderObj = {
    borderColor,
    borderColorObj,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
  }
  
  const chartLegendInfo = getChartLegend(plotLegend, warpObj, source)
  if (orithemeElements) {
    warpObj['themeContent']['a:theme']['a:themeElements'] = orithemeElements
  }
  const data = {
    type: 'chart',
    title: chartTitleInfo,
    table: chart.chartTable,
    spPrNode: chart.plotAreaSpPrNode,
    importProperty,
    chartSpPr,
    ...borderObj,
    legend: chartLegendInfo,
    top,
    left,
    width,
    height,
    data: chart.data,
    colors: chart.colors,
    chartType: chart.type,
    order,
  }
  if (chart.marker !== undefined) data.marker = chart.marker
  if (chart.barDir !== undefined) data.barDir = chart.barDir
  if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
  if (chart.grouping !== undefined) data.grouping = chart.grouping
  if (chart.style !== undefined) data.style = chart.style

  console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-dealNode:-data:', data)
  return data
}
// async function genWPSWebChart(node, warpObj, source) {
//   const order = node['attrs']['order']
//   const xfrmNode = getTextByPathList(node, ['p:xfrm'])
//   const { top, left } = getPosition(xfrmNode, undefined, undefined)
//   const { width, height } = getSize(xfrmNode, undefined, undefined)

//   const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
//   // console.log('(00)-pptxtojson-[chartEL]:-genChart-[rid]:', rid)
//   let refName = getTextByPathList(warpObj['slideResObj'], [rid, 'target'])
//   if (!refName) refName = getTextByPathList(warpObj['layoutResObj'], [rid, 'target'])
//   if (!refName) refName = getTextByPathList(warpObj['masterResObj'], [rid, 'target'])
//   if (!refName) return {}

//   const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
//   const type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

//   let slideLayoutSpNode, slideMasterSpNode

//   if (type) {
//     if (idx) {
//       slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
//       slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
//       if (!slideLayoutSpNode) slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
//       if (!slideMasterSpNode) slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
//     }
//     else {
//       slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
//       slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
//     }
//   }
//   else if (idx) {
//     slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
//     slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
//   }
//   let orithemeElements = null
//   if (refName) {
//     const themeName = refName.split('/').pop()
//     const dealSamePathFilePath = (target) => {
//       if (!target.includes('/')) {
//         return refName.replace(themeName, '') + target
//       }
//       return target
//     }
//     const getXMlFileContent = async (target) => {
//       // console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj]:-getXMlFileContent-target:', target)
//       const useTarget = dealSamePathFilePath(target)
//       // console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj]:-getXMlFileContent-useTarget:', useTarget)
//       if (!useTarget.endsWith('.xml')) return null
//       const content = await readXmlFile(warpObj['zip'], useTarget)
//       return content
//     }
//     const themeResFileName = refName.replace(themeName, '_rels/' + themeName) + '.rels'
//     const themeResContent = await readXmlFile(warpObj['zip'], themeResFileName)
//     console.log('(00)-pptxtojson-[chartEL]:-pie:-an:--themeResObj-warpObj:', warpObj)
//     // console.log('(00)-pptxtojson-[chartEL]:-genChart-[refName]:-themeResFileName:', '【', themeResFileName, '】【', refName, '】【', themeResContent, '】', warpObj)
//     const themeResObj = {}
//     if (themeResContent) {
//       let relationshipArray = themeResContent['Relationships']['Relationship']
//       if (relationshipArray) {
//         if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
//         for (const relationshipArrayItem of relationshipArray) {
//           themeResObj[relationshipArrayItem['attrs']['Id']] = {
//             'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
//             'target': dealSamePathFilePath(relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')),
//             'targetContent': await getXMlFileContent(relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'))
//           }
//           console.log('(00)-pptxtojson-[chartEL]:-pie:-an:--themeResObj:', relationshipArrayItem['attrs']['Id'], '--:', themeResObj[relationshipArrayItem['attrs']['Id']])
//         }
//       }
//     }

//     warpObj.themeResObj = themeResObj
//     // console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj,warpObj]:', themeResObj, warpObj)
//     // 获取图表主题
//     // const themeOverrideContent = Object.values(themeResObj).filter(item => item.type === 'themeOverride')
//     const themeOverrideContent = Object.values(themeResObj).find(item => item.type === 'themeOverride')
//     console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj,warpObj]:-themeOverrideContent', themeOverrideContent)
//     const contentObj = getTextByPathList(themeOverrideContent, ['targetContent', 'a:themeOverride'])
//     console.log('(00)-pptxtojson-[chartEL]:-genChart-get-[themeResObj,warpObj]:-contentObj:', contentObj)
//     if (contentObj && getTextByPathList(warpObj, ['themeContent', 'a:theme', 'a:themeElements'])) {
//       const themeElements = warpObj['themeContent']['a:theme']['a:themeElements']
//       orithemeElements = JSON.parse(JSON.stringify(themeElements))
//       warpObj['themeContent']['a:theme']['a:themeElements'] = {
//         ...warpObj['themeContent']['a:theme']['a:themeElements'],
//         ...contentObj
//       }
//     }
//     // for (const key in plotArea) {

//     // }
//   }

//   const content = await readXmlFile(warpObj['zip'], refName)
//   const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])
//   const otherParams = {
//     slideLayoutSpNode, 
//     slideMasterSpNode
//   }
//   const chart = getChartInfo(plotArea, warpObj, source, otherParams)

//   if (!chart) return {}
//   const plotTitle = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:title'])
//   const chartTitleInfo = getChartTitle(plotTitle, warpObj, source)
//   const plotLegend = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:legend'])
//   // console.log('(00)-pptxtojson-[chartEL]:-genChart-[refName]:', refName)  
//   const blipFill = getTextByPathList(content, ['c:chartSpace', 'c:spPr', 'a:blipFill'])
//   const picBase64 = await getPicFill('themeBg', blipFill, warpObj)
//   const background = {type: 'image', src: picBase64}
//   const importProperty = {
//     background,
//     plotAreaSpPr: chart ? chart.plotAreaSpPrNode : {},
//     chartTypeInfo: {
//       type: chart.type,
//       barDir: chart.barDir,
//       grouping: chart.grouping,
//     },
//     chartPr: chart.chartPr
//   }
  
//   const chartLegendInfo = getChartLegend(plotLegend, warpObj, source)
//   // console.log('(00)-pptxtojson-[chartEL]:-genChart-[chart]:', chart)
//   console.log('(00)-pptxtojson-[chartEL]:-genChart-[chart]-chart.type:', chart.type, chart)
//   if (orithemeElements) {
//     warpObj['themeContent']['a:theme']['a:themeElements'] = orithemeElements
//   }
//   const data = {
//     type: 'chart',
//     title: chartTitleInfo,
//     table: chart.chartTable,
//     spPrNode: chart.plotAreaSpPrNode,
//     importProperty,
//     legend: chartLegendInfo,
//     top,
//     left,
//     width,
//     height,
//     data: chart.data,
//     colors: chart.colors,
//     chartType: chart.type,
//     order,
//   }
//   if (chart.marker !== undefined) data.marker = chart.marker
//   if (chart.barDir !== undefined) data.barDir = chart.barDir
//   if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
//   if (chart.grouping !== undefined) data.grouping = chart.grouping
//   if (chart.style !== undefined) data.style = chart.style
//   console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-dealNode:-data:', data)
//   return data
// }

/**
 * 将 WPS 专有图表数据转为与 genChart 输出一致的格式
 * @param {Object} picInfo  - 已解析的 p:pic 元素信息
 * @param {Object} props    - extractProperties(webExtensionXML) 的结果
 * @returns {Object|null}
 */
function genWPSWebChart(picInfo, props) {
  if (!picInfo || !props) return null

  const demoData = props.demoData
  if (!demoData || !Array.isArray(demoData.data) || demoData.data.length < 2) return null

  const style = props.style || {}
  const extStyle = props.extStyle || {}
  const wpsType = props.type || '2d-pie'

  const DEFAULT_COLORS = ['#4874CB', '#EE822F', '#F2BA02', '#75BD42', '#30C0B4', '#E54C5E']
  const colors = (style.seriesThemeColor || (props.sourceTheme || {}).colors || []).filter(Boolean)
  if (!colors.length) colors.push(...DEFAULT_COLORS)

  const TYPE_MAP = {
    '2d-pie': 'pieChart',
    '2d-bar': 'barChart',
    '2d-line': 'lineChart',
    '2d-polar': 'polarChart',
    '2d-radar': 'radarChart',
    '2d-scatter': 'scatterChart',
    '2d-funnel': 'funnelChart',
  }
  console.log('(00)-pptxtojson-[chartEL]:-pie:-an:-p:pic-nodeValue:-extNode:-genWPSWebChartNode:-wpsType:props.type:', wpsType, props.type)
  const chartType = TYPE_MAP[wpsType] || 'customChart'

  const rawData = demoData.data
  const header = rawData[0] || []
  const dataRows = rawData.slice(1)
  const numSeries = Math.max(1, header.length - 1)

  const noBorder = {
    borderColor: '#000000',
    borderWidth: 0,
    borderType: 'solid',
    strokeDasharray: '0',
  }

  // ---- spPrNode / plotAreaSpPr ----
  const spPrNode = {
    border: {
      borderColor: picInfo.borderColor || '#000000',
      borderWidth: picInfo.borderWidth || 0,
      borderType: picInfo.borderType || 'solid',
      strokeDasharray: String(picInfo.borderStrokeDasharray || '0'),
    },
  }

  // ---- title ----
  const titleConf = style.title || {}
  const tts = titleConf.textStyle || {}
  const titleFontFamily = (tts.fontFamily && tts.fontFamily.name) || (tts.font && tts.font.name) || 'Arial'
  const titleFontSize = tts.fontSize || 14
  const titleFontWeight = tts.fontWeight || 'bold'

  let title = {}
  if (titleConf.show !== false && titleConf.text) {
    title = {
      text:
        '<p style="text-align: center;line-height: 1;margin: 0; padding: 0;">' +
        '<span style="font-size: ' + titleFontSize + 'pt;font-family: ' + titleFontFamily + ';' +
        'font-weight: ' + titleFontWeight + '; line-break: strict; overflow-wrap: break-word; white-space: pre-wrap;">' +
        titleConf.text + '</span></p>',
      layout: { xMode: 'edge', yMode: 'edge', x: '0.5', y: '0' },
      overlay: '0',
      spPr: { border: noBorder },
    }
  }

  // ---- legend ----
  const legendConf = style.legend || {}
  const LEGEND_POS = {
    topCenter: 't',
    bottomCenter: 'b',
    topLeft: 'tl',
    topRight: 'tr',
    bottomLeft: 'bl',
    bottomRight: 'br',
    left: 'l',
    right: 'r',
    center: 'ctr',
  }
  const legend = {
    legendPos: LEGEND_POS[legendConf.position] || 'b',
    overlay: '0',
    spPr: { border: noBorder },
    textpr: {
      'a:bodyPr': {
        attrs: { rot: '0', vertOverflow: 'ellipsis', vert: 'horz', wrap: 'square', anchor: 'ctr', anchorCtr: '1' },
      },
      'a:lstStyle': { attrs: {} },
      'a:p': {
        'a:pPr': {
          'a:defRPr': {
            attrs: {
              lang: 'zh-CN',
              sz: String((legendConf.textStyle && legendConf.textStyle.fontSize) || 900),
              b: '0', i: '0', u: 'none', strike: 'noStrike',
            },
          },
          attrs: {},
        },
        attrs: {},
      },
      attrs: {},
    },
  }

  // ---- data ----
  const isMultiColor = style.isMultiColorScheme !== false
  const extSeriesArr = Array.isArray(extStyle.series) ? extStyle.series : extStyle.series ? [extStyle.series] : []
  const extSeries0 = extSeriesArr[0] || {}
  const dPtBorderColor = (extSeries0.itemStyle && extSeries0.itemStyle.borderColor) || '#FFFFFF'
  const dPtBorderWidth = (extSeries0.itemStyle && extSeries0.itemStyle.borderWidth) || 2

  const chartData = []

  for (let s = 0; s < numSeries; s++) {
    const seriesName = header[s + 1] || '系列' + (s + 1)
    const values = []
    const xlabels = {}
    const dPt = []

    for (let i = 0; i < dataRows.length; i++) {
      const rawVal = dataRows[i][s + 1]
      const numVal = typeof rawVal === 'number' ? rawVal : parseFloat(rawVal) || 0
      values.push({ x: String(i), y: numVal })
      xlabels[String(i)] = String(dataRows[i][0] || '类别' + i)

      const colorIdx = isMultiColor ? i : s
      dPt.push({
        idx: String(i),
        bubble3D: '0',
        spPr: {
          border: {
            borderColor: dPtBorderColor,
            borderWidth: dPtBorderWidth,
            borderType: 'solid',
            strokeDasharray: '0',
          },
          fill: { type: 'color', value: colors[colorIdx % colors.length] },
        },
      })
    }

    chartData.push({
      key: seriesName,
      values,
      xlabels,
      spPr: {
        spPrNode: { border: noBorder },
        dLbls: { delete: !(style.label && style.label.show) },
        dPt,
      },
    })
  }

  // ---- importProperty ----
  const dLbls = {
    showLegendKey: false,
    showVal: !!(style.label && style.label.show),
    showCatName: false,
    showSerName: false,
    showPercent: false,
    showBubbleSize: false,
    showLeaderLines: true,
    delete: !(style.label && style.label.show),
  }

  const importProperty = {
    // background: { type: 'image', src: picInfo.src || '' },
    plotAreaSpPr: spPrNode,
    chartTypeInfo: {
      type: chartType,
      wpsChartType: wpsType,
      wpsRenderer: props.renderer || 'echarts',
    },
    chartPr: { dLbls },
  }

  const chartSpPr = {
    // background: { type: 'image', src: picInfo.src || '' },
    plotAreaSpPr: spPrNode,
    chartTypeInfo: {
      type: chartType,
      wpsChartType: wpsType,
      wpsRenderer: props.renderer || 'echarts',
    },
    chartPr: { dLbls },
  }

  // ---- 组装 ----
  const result = {
    type: 'chart',
    title,
    spPrNode,
    importProperty,
    chartSpPr,
    legend,
    top: picInfo.top,
    left: picInfo.left,
    width: picInfo.width,
    height: picInfo.height,
    data: chartData,
    chartType,
    order: picInfo.order,
    id: String(picInfo.id !== undefined && picInfo.id !== null ? picInfo.id : ''),
  }

  // 可选字段
  if (wpsType === '2d-pie' && Array.isArray(style.radius) && parseInt(style.radius[0]) > 0) {
    result.holeSize = parseInt(style.radius[0])
  }
  if (wpsType === '2d-bar') result.barDir = 'col'
  if (wpsType === '2d-line') result.marker = true

  return result
}

async function genDiagram(node, warpObj) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  
  const dgmDrwSpArray = getTextByPathList(warpObj['digramFileContent'], ['p:drawing', 'p:spTree', 'p:sp'])
  const elements = []
  let textList = []
  if (dgmDrwSpArray) {
    const spList = Array.isArray(dgmDrwSpArray) ? dgmDrwSpArray : [dgmDrwSpArray]

    for (const item of spList) {
      const el = await processSpNode(item, warpObj, 'diagramBg')
      if (el) elements.push(el)
    }
  }
  else if (warpObj.diagramContent && warpObj.diagramContent.data) {
    textList = getSmartArtTextData(warpObj.diagramContent.data)
  }

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
    textList,
    order,
  }
}