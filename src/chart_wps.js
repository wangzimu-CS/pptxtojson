// import { eachElement, getTextByPathList } from './utils'
// import { applyTint } from './color'
// import { getShadow } from './shadow'
// import { getGlow, getSoftEdge } from './glow'
// import { getBorder } from './border'
// import { genTextBody, getSpanStyleInfo } from './text'
// import { getSolidFill, getGradientFill } from './fill'

// ================================================================
// 工具函数
// ================================================================

// ================================================================
// 安全取值辅助函数（替代可选链 ?.）
// ================================================================
export function safe(obj, a, b, c, d, e) {
  if (obj === null) return undefined
  let r = obj[a]
  if (r === null) return undefined
  if (b !== undefined) {
    r = r[b]; if (r === null) return undefined 
  }
  if (c !== undefined) {
    r = r[c]; if (r === null) return undefined 
  }
  if (d !== undefined) {
    r = r[d]; if (r === null) return undefined 
  }
  if (e !== undefined) {
    r = r[e]; if (r === null) return undefined 
  }
  return r
}

export function unescapeHtml(s) {
  if (typeof s !== 'string') return s
  return s.replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'")
}

export function extractProperties(extJson) {
  const props = {}
  const ext = extJson['wpswe:webExtension'] || extJson
  const rawItems = safe(ext, 'wpswe:properties', 'wpswe:property')
  if (!rawItems) return props
  const items = Array.isArray(rawItems) ? rawItems : [rawItems]
  for (let i = 0; i < items.length; i++) {
    const item = items[i]
    const k = safe(item, 'attrs', 'key')
    const v = safe(item, 'attrs', 'value')
    if (!k || v === undefined) continue
    try {
      props[k] = JSON.parse(unescapeHtml(v)) 
    }
    catch (e) {
      props[k] = unescapeHtml(v) 
    }
  }
  return props
}

export function wpsColor(c) {
  if (!c) return '#000'
  if (typeof c === 'string') return c
  if (c.rgb) return c.rgb
  return '#000'
}

export function wpsTextStyle(ts) {
  if (!ts) return {}
  const result = {}
  result.color = wpsColor(ts.color)
  const fname = safe(ts, 'fontFamily', 'name') || safe(ts, 'font', 'name') || 'sans-serif'
  result.fontFamily = fname
  result.fontSize = ts.fontSize || 12
  result.fontWeight = ts.fontWeight === 'bold' ? 'bold' : 'normal'
  return result
}

export function parseNum(v) {
  return typeof v === 'number' ? v : (parseFloat(v) || 0)
}

export function toArray(arr) {
  return Array.isArray(arr) ? arr : [arr]
}
