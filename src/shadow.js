import { getSolidFill } from './fill'
import { RATIO_EMUs_Points } from './constants'

export function getShadow(node, warpObj) {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  const dir = outerShdwAttrs['dir'] ? (parseInt(outerShdwAttrs['dir']) / 60000) : 0
  const dist = outerShdwAttrs['dist'] ? parseInt(outerShdwAttrs['dist']) * RATIO_EMUs_Points : 0
  const blurRad = outerShdwAttrs['blurRad'] ? parseInt(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points : ''
  const vx = dist * Math.sin(dir * Math.PI / 180)
  const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    h: hx,
    v: vx,
    blur: blurRad,
    color: chdwClrNode,
  }
}

export function getGlow(node, warpObj) {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  const dir = outerShdwAttrs['rad'] ? (parseInt(outerShdwAttrs['rad']) / 60000) : 0
  const dist = outerShdwAttrs['rad'] ? parseInt(outerShdwAttrs['rad']) * RATIO_EMUs_Points : 0
  const blurRad = outerShdwAttrs['blurRad'] ? parseInt(outerShdwAttrs['blurRad']) * RATIO_EMUs_Points : ''
  // const vx = dist * Math.sin(dir * Math.PI / 180)
  // const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    h: dir,
    v: dist,
    blur: blurRad,
    color: chdwClrNode,
  }
}