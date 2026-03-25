import { getSolidFill } from './fill'
import { RATIO_EMUs_Points } from './constants'

export function getGlow(node, warpObj) {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  // const dir = outerShdwAttrs['rad'] ? (parseInt(outerShdwAttrs['rad']) / 60000) : 0
  const dist = outerShdwAttrs['rad'] ? parseInt(outerShdwAttrs['rad']) * RATIO_EMUs_Points : 0
  // const blurRad = outerShdwAttrs['blurRad'] ? parseInt(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points : ''
  // const vx = dist * Math.sin(dir * Math.PI / 180)
  // const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    rad: dist,
    color: chdwClrNode,
  }
}

export function getSoftEdge(node) {
  // const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  // console.log('(00)-pptxtojson-getSoftEdge--outerShdwAttrs:', outerShdwAttrs)
  // const dir = outerShdwAttrs['rad'] ? (parseInt(outerShdwAttrs['rad']) / 60000) : 0
  const rad = outerShdwAttrs['rad'] ? parseInt(outerShdwAttrs['rad']) * RATIO_EMUs_Points : 0
  // const blurRad = outerShdwAttrs['blurRad'] ? parseInt(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points : ''
  // const vx = dist * Math.sin(dir * Math.PI / 180)
  // const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    rad: rad,
  }
}