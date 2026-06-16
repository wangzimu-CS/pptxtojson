import { eachElement, getTextByPathList } from './utils'
import { applyTint } from './color'
import { getShadow } from './shadow'
import { getGlow, getSoftEdge } from './glow'
import { getBorder } from './border'
import { genTextBody, getSpanStyleInfo } from './text'
import { getSolidFill, getGradientFill } from './fill'

function extractChartColors(serNode, warpObj) {
  if (serNode && serNode.constructor !== Array) serNode = [serNode]
  if (!serNode) return
  const schemeClrs = []
  for (const node of serNode) {
    // let schemeClr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:schemeClr'])
    // if (!schemeClr) schemeClr = getTextByPathList(node, ['c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])
    // if (!schemeClr) schemeClr = getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])

    // let clr = getTextByPathList(schemeClr, ['attrs', 'val'])
    // if (clr) {
    //   clr = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:clrScheme', `a:${clr}`, 'a:srgbClr', 'attrs', 'val'])
    //   const tint = getTextByPathList(schemeClr, ['a:tint', 'attrs', 'val']) / 100000
    //   if (clr && !isNaN(tint)) {
    //     clr = applyTint(clr, tint)
    //   }
    // }
    // else clr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])

    // if (clr) clr = '#' + clr
    // schemeClrs.push(clr)

    {
      const seriesColorObj = {
        solidFill: getSolidFill(getTextByPathList(node, ['c:spPr', 'a:solidFill']), undefined, undefined, warpObj),
        lnSolidFill: getSolidFill(getTextByPathList(node, ['c:spPr', 'a:ln', 'a:solidFill']), undefined, undefined, warpObj),
        markerLnSolidFill: getSolidFill(getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:solidFill']), undefined, undefined, warpObj),
        gradFill: getGradientFill(getTextByPathList(node, ['c:spPr', 'a:gradFill']), warpObj),
        lnGradFill: getGradientFill(getTextByPathList(node, ['c:spPr', 'a:ln', 'a:gradFill']), warpObj),
        markerLnGradFill: getGradientFill(getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:gradFill']), warpObj),
      }
      // console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]-seriesColorObj:', seriesColorObj)

      let schemeClr = getTextByPathList(node, ['c:spPr', 'a:solidFill'])
      if (!schemeClr) schemeClr = getTextByPathList(node, ['c:spPr', 'a:ln', 'a:solidFill'])
      if (!schemeClr) schemeClr = getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:solidFill'])
    
      let clr = ''
      console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]-node:', node, schemeClr)

      const clrObj = {}

      if (schemeClr) {
        clrObj.type = 'color'
        clrObj.value = getSolidFill(schemeClr, undefined, undefined, warpObj)
      }
      const abs = true
      if (JSON.stringify(clrObj) === '{}' || abs) {
        // 颜色获取失败，考虑获取渐变色
        let schemeClrINfoObj = getTextByPathList(node, ['c:spPr', 'a:gradFill'])
        if (!schemeClrINfoObj) schemeClrINfoObj = getTextByPathList(node, ['c:spPr', 'a:ln', 'a:gradFill'])
        if (!schemeClrINfoObj) schemeClrINfoObj = getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:gradFill'])
        const shpFill = schemeClrINfoObj
        if (shpFill) {
          clrObj.type = 'gradient'
          clrObj.value = getGradientFill(shpFill, warpObj)
        }
      }

      clr = clrObj
      clr.seriesColorObj = seriesColorObj
      schemeClrs.push(clr)
      console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]:', schemeClr)
    }
    console.log('')
    console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClrs]:', schemeClrs)
  }
  return schemeClrs
}

function extractChartData(serNode, warpObj, source, otherParams) {
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

      const spPr = getSeriesItemPr(innerNode, warpObj, source, otherParams)
      console.log('(00)-pptxtojson-[chartEL]:-pie:-anylsis[spPr]:', spPr, innerNode, dataRow)

      dataMat.push({
        key: colName,
        values: dataRow,
        xlabels: rowNames,
        spPr
      })
      return ''
    })
  }

  console.log('(00)-pptxtojson-[chartEL]:-anylisChart-dataMat:', dataMat)
  return dataMat
}

function getSeriesItemPr(innerNode, warpObj, source, otherParams) {
  // console.log('(00)-pptxtojson-[chartEL]:-anylisChart-innerNode:', innerNode)
  const spPr = {}
  // c:spPr
  const spPrNode = getTextByPathList(innerNode, ['c:spPr'])
  if (spPrNode) spPr.spPrNode = getPrInfo(spPrNode, warpObj, source)
  // c:smooth
  const smoothNode = getTextByPathList(innerNode, ['c:smooth'])
  if (smoothNode) {
    const {val} = getNodeAttrsObj(smoothNode)
    spPr.smooth = val === '1' ? true : false
  }
  // c:marker
  const markerNode = getTextByPathList(innerNode, ['c:marker', 'c:symbol'])
  if (markerNode) {
    const {val} = getNodeAttrsObj(markerNode)
    spPr.markerSymbol = val
  }
  // c:dLbls
  const dLblsNode = getTextByPathList(innerNode, ['c:dLbls'])
  // if (dLblsNode) spPr.dLblsNode = dLblsNode
  if (dLblsNode) spPr.dLbls = getDLblsInfo(dLblsNode, warpObj, source, otherParams)
  //
  // console.log('(00)-pptxtojson-[chartEL]:-anylisChart-spPr:', spPr)
  const gapWidth = getTextByPathList(innerNode, ['c:gapWidth', 'attrs', 'val'])
  if (gapWidth) spPr.gapWidth = Number(gapWidth)

  const overlap = getTextByPathList(innerNode, ['c:overlap', 'attrs', 'val'])
  if (overlap) spPr.overlap = Number(overlap)
    
  const dPt = getTextByPathList(innerNode, ['c:dPt'])
  if (dPt && Array.isArray(dPt)) {
    // spPr.dPt = Number(dPt)
    for (const index in dPt) {
      const item = dPt[index]
    }
  }
  console.log('(00)-pptxtojson-[chartEL]:-pie:-anylsis[dPt]:', dPt)
    
}

function getChartPr(chartNode, warpObj, source, otherParams, plotArea) {
  console.log('(00)-pptxtojson-[chartEL]:-serNode:', chartNode)
  console.log('(00)-pptxtojson-[chartEL]:-serNode:-plotArea:', plotArea)

  if (!chartNode) return null
  
  // console.log('(00)-pptxtojson-[chartEL]:-anylisChart-innerNode:', innerNode)
  const innerNode = chartNode
  const spPr = {}
  // c:spPr
  const spPrNode = getTextByPathList(innerNode, ['c:spPr'])
  if (spPrNode) spPr.spPrNode = getPrInfo(spPrNode, warpObj, source)
  // c:smooth
  const smoothNode = getTextByPathList(innerNode, ['c:smooth'])
  if (smoothNode) {
    const {val} = getNodeAttrsObj(smoothNode)
    spPr.smooth = val === '1' ? true : false
  }
  // c:marker
  const markerNode = getTextByPathList(innerNode, ['c:marker', 'c:symbol'])
  if (markerNode) {
    const {val} = getNodeAttrsObj(markerNode)
    spPr.markerSymbol = val
  }
  // c:dLbls
  const dLblsNode = getTextByPathList(innerNode, ['c:dLbls'])
  // if (dLblsNode) spPr.dLblsNode = dLblsNode
  if (dLblsNode) spPr.dLbls = getDLblsInfo(dLblsNode, warpObj, source, otherParams)
  // console.log('(00)-pptxtojson-[chartEL]:-anylisChart-spPr:', spPr)
  const gapWidth = getTextByPathList(innerNode, ['c:gapWidth', 'attrs', 'val'])
  if (gapWidth) spPr.gapWidth = Number(gapWidth)

  const overlap = getTextByPathList(innerNode, ['c:overlap', 'attrs', 'val'])
  if (overlap) spPr.overlap = Number(overlap)

  console.log('(00)-pptxtojson-[chartEL]:-serNode:--result:', spPr)
  return spPr  
}
function getDLblsInfo(node, warpObj, source, otherParams) {
  const dLblsInfo = {}
  for (const key in node) {
    const curNode = node[`${key}`]
    const curSubItemVal = getTextByPathList(curNode, ['attrs', 'val'])
    let isOnlyAttrs = true
    for (const subKey in curNode) {
      if (subKey !== 'attrs') isOnlyAttrs = false ; break
    }
    if (curSubItemVal && key.includes('c:') && isOnlyAttrs) {
      dLblsInfo[`${key.replace('c:', '')}`] = curSubItemVal === '0' || curSubItemVal === '1' ? (curSubItemVal === '1' ? true : false) : curSubItemVal
    }
    else if (!isOnlyAttrs && key.includes('c:')) {
      //
      // dLblsInfo[`${key.replace('c:', '')}`] = curNode
      switch (key) {
        case 'c:spPr':
          if (curNode) dLblsInfo.spPr = getPrInfo(curNode, warpObj, source)
          break
        case 'c:txPr':
          const dealUseObj = {}
          dealUseObj[`p:txBody`] = curNode
          console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]-get-txPr:-curNode:', curNode)
          const {
            slideLayoutSpNode, 
            slideMasterSpNode
          } = otherParams 
          // const pFontStyle = getTextByPathList(spNode, ['p:style', 'a:fontRef'])
          // const pNode = textBodyNode['a:p']
          // getSpanStyleInfo(rNodeItem, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
          const pFontStyle = null
          const schemeClr = getTextByPathList(curNode, ['a:p', 'a:pPr', 'a:defRPr', 'a:solidFill'])
          const textFill = getSolidFill(schemeClr, undefined, undefined, warpObj)
          console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]-get-txPr:-schemeClr:', schemeClr, textFill)
          const pNode = getTextByPathList(curNode, ['a:p']) 
          const styleObj = getSpanStyleInfo(null, pNode, curNode, pFontStyle, slideLayoutSpNode, slideMasterSpNode, undefined, warpObj)
          console.log('(00)-pptxtojson-[chartEL]:-analysis-[schemeClr]-get-txPr:-styleObj:', styleObj)
          const text = genTextBody(dealUseObj['p:txBody'], dealUseObj, undefined, undefined, undefined, warpObj)
          dLblsInfo.txPr = {
            text,
            textColor: textFill,
            styleObj
          }
          break
      
        default:
          break
      }
    }

  }
  return dLblsInfo 
}

export function getChartInfo(plotArea, warpObj, source, otherParams) {
  let chart = null
  for (const key in plotArea) {
    console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[key]:', key)
    let isChartKey = true
    switch (key) {
      case 'c:lineChart':
        chart = {
          type: 'lineChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          marker: plotArea[key]['c:marker'] ? true : false,
        }
        break
      case 'c:line3DChart':
        chart = {
          type: 'line3DChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:barChart':
        chart = {
          type: 'barChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          barDir: getTextByPathList(plotArea[key], ['c:barDir', 'attrs', 'val']),
        }
        break
      case 'c:bar3DChart':
        chart = {
          type: 'bar3DChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
          barDir: getTextByPathList(plotArea[key], ['c:barDir', 'attrs', 'val']),
        }
        break
      case 'c:pieChart':
        chart = {
          type: 'pieChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
        }
        break
      case 'c:pie3DChart':
        chart = {
          type: 'pie3DChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
        }
        break
      case 'c:doughnutChart':
        chart = {
          type: 'doughnutChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser']['c:dPt'], warpObj),
          holeSize: getTextByPathList(plotArea[key], ['c:holeSize', 'attrs', 'val']),
        }
        break
      case 'c:areaChart':
        chart = {
          type: 'areaChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:area3DChart':
        chart = {
          type: 'area3DChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          grouping: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']),
        }
        break
      case 'c:scatterChart':
        chart = {
          type: 'scatterChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          style: getTextByPathList(plotArea[key], ['c:scatterStyle', 'attrs', 'val']),
        }
        break
      case 'c:bubbleChart':
        chart = {
          type: 'bubbleChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:radarChart':
        chart = {
          type: 'radarChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
          style: getTextByPathList(plotArea[key], ['c:radarStyle', 'attrs', 'val']),
        }
        break
      case 'c:surfaceChart':
        chart = {
          type: 'surfaceChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:surface3DChart':
        chart = {
          type: 'surface3DChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: extractChartColors(plotArea[key]['c:ser'], warpObj),
        }
        break
      case 'c:stockChart':
        chart = {
          type: 'stockChart',
          data: extractChartData(plotArea[key]['c:ser'], warpObj, source, otherParams),
          colors: [],
        }
        break
      default:
        isChartKey = false
    }
    if (isChartKey) chart.chartPr = getChartPr(plotArea[key], warpObj, source, otherParams, plotArea)
  }
  // 其它属性解析
  // console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[plotArea]:', plotArea)

  const dTableNode = getTextByPathList(plotArea, ['c:dTable'])
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[dTableNode]:', dTableNode)
  if (dTableNode && chart) chart.chartTable = getChartTableInfo(dTableNode, warpObj, source)

  const spPrNode = getTextByPathList(plotArea, ['c:spPr'])
  console.log('(00)-pptxtojson-[chartEL]:-genChart--getChartInfo-[spPrNode]:', spPrNode)
  if (spPrNode && chart) chart.plotAreaSpPrNode = getPrInfo(spPrNode, warpObj, source)
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

export function getChartLegend(cLegendNode, warpObj, source) {
  if (!cLegendNode) return null
  const legendInfo = {}
  warpObj
  // c:legend  c:legendPos
  for (const key in cLegendNode) {
    const curNode = cLegendNode[`${key}`]
    const curSubItemVal = getTextByPathList(curNode, ['attrs', 'val'])
    let isOnlyAttrs = true
    for (const subKey in curNode) {
      if (subKey !== 'attrs') isOnlyAttrs = false ; break
    }
    if (curSubItemVal && key.includes('c:') && isOnlyAttrs) legendInfo[`${key.replace('c:', '')}`] = curSubItemVal
    else if (!isOnlyAttrs && key.includes('c:')) {
      switch (key) {
        case 'c:txPr':
          const txPrNode = curNode
          // 处理成文本能处理的对象
          const attrsNode = getNodeAttrsObj(txPrNode)
          const dealUseObj = {}
          dealUseObj[`p:txBody`] = txPrNode
          const text = genTextBody(dealUseObj['p:txBody'], dealUseObj, undefined, undefined, undefined, warpObj)
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

  const solidFillNode = getTextByPathList(prNode, ['a:solidFill'])
  if (solidFillNode) {
    const fill = getSolidFill(solidFillNode, undefined, undefined, warpObj)
    propertySettings.fill = {
      type: 'color',
      value: fill
    }
  }

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
