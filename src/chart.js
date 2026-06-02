import { eachElement, getTextByPathList } from './utils'
import { applyTint } from './color'
import { getShadow } from './shadow'
import { getGlow, getSoftEdge } from './glow'
import { getBorder } from './border'
import { genTextBody } from './text'
import { getShapeFill, getSolidFill } from './fill'

function extractChartColors(serNode, warpObj) {
  if (serNode && serNode.constructor !== Array) serNode = [serNode]
  if (!serNode) return
  const schemeClrs = []
  for (const node of serNode) {
    let schemeClr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:schemeClr'])
    if (!schemeClr) schemeClr = getTextByPathList(node, ['c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])
    if (!schemeClr) schemeClr = getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])

    let clr = getTextByPathList(schemeClr, ['attrs', 'val'])
    if (clr) {
      clr = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:clrScheme', `a:${clr}`, 'a:srgbClr', 'attrs', 'val'])
      const tint = getTextByPathList(schemeClr, ['a:tint', 'attrs', 'val']) / 100000
      if (clr && !isNaN(tint)) {
        clr = applyTint(clr, tint)
      }
    }
    else clr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])

    if (clr) clr = '#' + clr
    schemeClrs.push(clr)
  }
  return schemeClrs
}

function extractChartData(serNode) {
  const dataMat = []
  if (!serNode) return dataMat

  if (serNode['c:xVal']) {
    let dataRow = []
    eachElement(serNode['c:xVal']['c:numRef']['c:numCache']['c:pt'], innerNode => {
      dataRow.push(parseFloat(innerNode['c:v']))
      return ''
    })
    dataMat.push(dataRow)
    dataRow = []
    eachElement(serNode['c:yVal']['c:numRef']['c:numCache']['c:pt'], innerNode => {
      dataRow.push(parseFloat(innerNode['c:v']))
      return ''
    })
    dataMat.push(dataRow)
  } 
  else {
    eachElement(serNode, (innerNode, index) => {
      const dataRow = []
      const colName = getTextByPathList(innerNode, ['c:tx', 'c:strRef', 'c:strCache', 'c:pt', 'c:v']) || index

      const rowNames = {}
      if (getTextByPathList(innerNode, ['c:cat', 'c:strRef', 'c:strCache', 'c:pt'])) {
        eachElement(innerNode['c:cat']['c:strRef']['c:strCache']['c:pt'], innerNode => {
          rowNames[innerNode['attrs']['idx']] = innerNode['c:v']
          return ''
        })
      } 
      else if (getTextByPathList(innerNode, ['c:cat', 'c:numRef', 'c:numCache', 'c:pt'])) {
        eachElement(innerNode['c:cat']['c:numRef']['c:numCache']['c:pt'], innerNode => {
          rowNames[innerNode['attrs']['idx']] = innerNode['c:v']
          return ''
        })
      }

      if (getTextByPathList(innerNode, ['c:val', 'c:numRef', 'c:numCache', 'c:pt'])) {
        eachElement(innerNode['c:val']['c:numRef']['c:numCache']['c:pt'], innerNode => {
          dataRow.push({
            x: innerNode['attrs']['idx'],
            y: parseFloat(innerNode['c:v']),
          })
          return ''
        })
      }

      dataMat.push({
        key: colName,
        values: dataRow,
        xlabels: rowNames,
      })
      return ''
    })
  }

  return dataMat
}

export function getChartInfo(plotArea, warpObj, source) {
  let chart = null
  for (const key in plotArea) {
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[key]:', key)
    switch (key) {
      case 'c:lineChart':
        chart = {
          type: 'lineChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          marker: plotArea[key]['c:marker'] ? true : false,
        }
        break
      case 'c:line3DChart':
        chart = {
          type: 'line3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:barChart':
        chart = {
          type: 'barChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          barDir: getTextByPathList(plotArea[key], ['c:barDir', 'attrs', 'val']),
        }
        break
      case 'c:bar3DChart':
        chart = {
          type: 'bar3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          barDir: getTextByPathList(plotArea[key], ['c:barDir', 'attrs', 'val']),
        }
        break
      case 'c:pieChart':
        chart = {
          type: 'pieChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
        }
        break
      case 'c:pie3DChart':
        chart = {
          type: 'pie3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
        }
        break
      case 'c:doughnutChart':
        chart = {
          type: 'doughnutChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
          holeSize: getTextByPathList(plotArea[key], ['c:holeSize', 'attrs', 'val']),
        }
        break
      case 'c:areaChart':
        chart = {
          type: 'areaChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:area3DChart':
        chart = {
          type: 'area3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:scatterChart':
        chart = {
          type: 'scatterChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          style: getTextByPathList(plotArea[key], ['c:scatterStyle', 'attrs', 'val']),
        }
        break
      case 'c:bubbleChart':
        chart = {
          type: 'bubbleChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:radarChart':
        chart = {
          type: 'radarChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          style: getTextByPathList(plotArea[key], ['c:radarStyle', 'attrs', 'val']),
        }
        break
      case 'c:surfaceChart':
        chart = {
          type: 'surfaceChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:surface3DChart':
        chart = {
          type: 'surface3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:stockChart':
        chart = {
          type: 'stockChart',
          data: extractChartData(plotArea[key]['c:ser']),
          colors: [],
        }
        break
      default:
    }
  }
  // 其它属性解析
  // console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[plotArea]:', plotArea)

  const dTableNode = getTextByPathList(plotArea, ['c:dTable'])
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[dTableNode]:', dTableNode)
  if (dTableNode) chart.chartTable = getChartTableInfo(dTableNode, warpObj, source)

  const spPrNode = getTextByPathList(plotArea, ['c:spPr'])
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[spPrNode]:', spPrNode)
  if (spPrNode) chart.spPrNode = getPrInfo(spPrNode, warpObj, source)
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[chart]:', chart)

  return chart
}

export function getChartTitle(cTitleNode, warpObj, source) {
  // const titleInfo = null
  const titleInfo = {}
  warpObj
  for (const key in cTitleNode) {
    // console.log('(00)-pptxtojson-[chartEL]:-genChart-[content]-getChartTitle-[key]:', key)
    // console.log('(00)-pptxtojson-[chartEL]:-genChart-[content]-getChartTitle-[key:value]:', key, cTitleNode[`${key}`])
    const curNode = cTitleNode[`${key}`]
    switch (key) {
      case 'c:tx':
        const richNode = getTextByPathList(curNode, ['c:rich'])
        // 处理成文本能处理的对象
        const dealUseObj = {}
        dealUseObj[`p:txBody`] = richNode
        const text = genTextBody(dealUseObj['p:txBody'], dealUseObj, undefined, undefined, undefined, warpObj)
        titleInfo.text = text || ''
        break
      case 'c:layout':
        const manualLayoutNode = getTextByPathList(curNode, ['c:manualLayout'])
        // const layoutNode = curNode
        if (manualLayoutNode) {
          const layoutInfo = {}
          for (const subKey in manualLayoutNode) {
            const curSubItemNode = manualLayoutNode[`${subKey}`]
            if (subKey.includes('c:')) {
              const curSubItemVal = getTextByPathList(curSubItemNode, ['attrs', 'val'])
              layoutInfo[`${subKey.replace('c:', '')}`] = curSubItemVal
            }
          }
          titleInfo.layout = layoutInfo
        }
        break
      case 'c:overlay':
        const overlayVal = getTextByPathList(curNode, ['attrs', 'val'])
        if (overlayVal) titleInfo.overlay = overlayVal
        // const overlayNode = curNode
        break
      case 'c:spPr':
        // const spPrNode = getTextByPathList(curNode, ['c:spPr'])
        const spPrNode = curNode
        const spPrInfo = getPrInfo(spPrNode, warpObj, source)
        if (spPrInfo) titleInfo.spPr = spPrInfo
        break
      default:
    }
  }
  // console.log('(00)-pptxtojson-[chartEL]:-genChart-[content]-getChartTitle-[titleInfo]111111111:', titleInfo)

  return titleInfo
}

export function getChartLegend(cLegendNode, warpObj, source, chartData) {
  // const legendInfo = null
  const legendInfo = {}
  warpObj
  console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[cLegendNode]:', cLegendNode)
  console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[chartData]:', chartData)
  for (const key in cLegendNode) {
    // console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[key]:', key)
    console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[key:value]:', key, cLegendNode[`${key}`])
    const curNode = cLegendNode[`${key}`]
    const curSubItemVal = getTextByPathList(curNode, ['attrs', 'val'])
    let isOnlyAttrs = true
    for (const subKey in curNode) {
      if (subKey !== 'attrs') isOnlyAttrs = false ; break
    }
    if (curSubItemVal && key.includes('c:') && isOnlyAttrs) legendInfo[`${key.replace('c:', '')}`] = curSubItemVal
    else if (!isOnlyAttrs && key.includes('c:')) {
      console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[未解析]-----key:', key, cLegendNode[`${key}`])
      switch (key) {
        case 'c:txPr':
          const txPrNode = curNode
          // 处理成文本能处理的对象
          console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[未解析]-----key-[txPrNode]:', txPrNode)
          const attrsNode = getNodeAttrsObj(txPrNode)
          console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[未解析]-----key-[attrsNode]:', attrsNode)
          const dealUseObj = {}
          dealUseObj[`p:txBody`] = txPrNode
          const text = genTextBody(dealUseObj['p:txBody'], dealUseObj, undefined, undefined, undefined, warpObj)
          console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[未解析]-----key-[text]:', text)
          legendInfo.textpr = txPrNode
          break
        case 'c:layout':
          const manualLayoutNode = getTextByPathList(curNode, ['c:manualLayout'])
          // const layoutNode = curNode
          if (manualLayoutNode) {
            const layoutInfo = {}
            for (const subKey in manualLayoutNode) {
              const curSubItemNode = manualLayoutNode[`${subKey}`]
              if (subKey.includes('c:')) {
                const curSubItemVal = getTextByPathList(curSubItemNode, ['attrs', 'val'])
                layoutInfo[`${subKey.replace('c:', '')}`] = curSubItemVal
              }
            }
            legendInfo.layout = layoutInfo
          }
          break
        case 'c:overlay':
          const overlayVal = getTextByPathList(curNode, ['attrs', 'val'])
          if (overlayVal) legendInfo.overlay = overlayVal
          // const overlayNode = curNode
          break
        case 'c:spPr':
        // const spPrNode = getTextByPathList(curNode, ['c:spPr'])
          const spPrNode = curNode
          const spPrInfo = getPrInfo(spPrNode, warpObj, source)
          if (spPrInfo) legendInfo.spPr = spPrInfo
          break
        default:
      }
    }
    console.log('(00)-pptxtojson-[chartEL]:-genChart-[plotLegend]-getChartLegend-[legendInfo]:', legendInfo)

  }
  // console.log('(00)-pptxtojson-[chartEL]:-genChart-[content]-getChartLegend-[legendInfo]111111111:', legendInfo)

  return legendInfo
}

function getChartTableInfo(dTableNode, warpObj, source) {
  const chartTable = {}
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[dTableNode]:', dTableNode)
  for (const key in dTableNode) {
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-getChartTableInfo-[key]:', key)
    if (key.includes('c:show')) chartTable[`${key.replace('c:', '')}`] = getTextByPathList(dTableNode, [`${key}`, 'attrs', 'val'])

    if (key === 'c:spPr') {
      //
      const spPrNode = getTextByPathList(dTableNode, [`${key}`])
      if (spPrNode) chartTable.spPrNode = spPrNode
      if (spPrNode) chartTable.spPrInfo = getPrInfo(spPrNode, warpObj, source)

    }

    if (key === 'c:txPr') {
      //
      const txPrNode = getTextByPathList(dTableNode, [`${key}`])
      if (txPrNode) chartTable.txPrNode = txPrNode
    }
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-getChartTableInfo-[key]:', key)
  }
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[chartTable]:', chartTable)
  return chartTable 
}

function getPrInfo(prNode, warpObj, source, tag = 'a:') {
  tag
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-getPrInfo-[prNode]:', prNode)
  const propertySettings = {}
  let effectData
  {
    const data = {}
    // const propertySettings = {}
    let offectObj = {}
    // const defaultStyleObj = getTextByPathList(prNode, ['p:style'])
    // const defaultEffectRef = getTextByPathList(prNode, ['p:style', 'a:effectRef'])
  
    let shadow
    const outerShdwNode = getTextByPathList(prNode, ['a:effectLst', 'a:outerShdw'])
    if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)
  
    let glow
    const glowNode = getTextByPathList(prNode, ['a:effectLst', 'a:glow'])
    if (glowNode) glow = getGlow(glowNode, warpObj)
  
    // 获取倒影配置
    let reflection
    const reflectionNode = getTextByPathList(prNode, ['a:effectLst', 'a:reflection'])
    const effectLstNode = getTextByPathList(prNode, ['a:effectLst'])
    if (effectLstNode && !reflectionNode) {
      //
    }
    else if (!effectLstNode && !reflectionNode) {
      const effectStyleLst = getTextByPathList(warpObj, ['themeContent', 'a:theme', 'a:themeElements', 'a:fmtScheme', 'a:effectStyleLst'])
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
  
    let softEdge
    const softEdgeNode = getTextByPathList(prNode, ['a:effectLst', 'a:softEdge'])
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
    effectData = data
  }
  const tableBorder = getBorder(prNode, '', warpObj)
  propertySettings.border = tableBorder

  const fill = getSolidFill(prNode, undefined, undefined, warpObj)
  propertySettings.fill = fill
  return {...effectData, ...propertySettings}
}

export function dealListItem(list, target) {
  list, target
  // if(){}
}

export function getNodeAttrsObj(node) {
  const attrsNode = getTextByPathList(node, ['attrs'])
  for (const key in attrsNode) {
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-getNodeAttrsObj-[key]:', key)
  }
  return attrsNode
}
