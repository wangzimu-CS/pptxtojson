import { eachElement, getTextByPathList } from './utils'
import { applyTint } from './color'
import { getShadow } from './shadow'
import { getGlow, getSoftEdge } from './glow'
import { getBorder } from './border'

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

export function getChartInfo(plotArea, warpObj) {
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
  const dTableNode = getTextByPathList(plotArea, ['c:dTable'])
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[dTableNode]:', dTableNode)
  if (dTableNode) chart.chartTable = getChartTableInfo(dTableNode, warpObj)

  return chart
}

function getChartTableInfo(dTableNode, warpObj) {
  const chartTable = {}
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[dTableNode]:', dTableNode)
  for (const key in dTableNode) {
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-getChartTableInfo-[key]:', key)
    if (key.includes('c:show')) chartTable[`${key.replace('c:', '')}`] = getTextByPathList(dTableNode, [`${key}`, 'attrs', 'val'])

    if (key === 'c:spPr') {
      //
      const spPrNode = getTextByPathList(dTableNode, [`${key}`])
      if (spPrNode) chartTable.spPrNode = spPrNode
      if (spPrNode) chartTable.spPrInfo = getPrInfo(spPrNode, warpObj)

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

function getPrInfo(prNode, warpObj) {
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
    console.log('(00)-ppt-math-el:data:', data)
    effectData = data
  }
  const tableBorder = getBorder(prNode, '', warpObj)
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-{tableBorder}:', tableBorder)
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-{...effectData, ...propertySettings}:', {...effectData, ...propertySettings})
  return {...effectData, ...propertySettings}
}

