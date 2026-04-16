import JSZip from 'jszip'
import type { SlideBackground } from '@/types/slides'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const PRESENTATION_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const REL_TYPES = {
  slideLayout: `${PRESENTATION_REL_NS}/slideLayout`,
  slideMaster: `${PRESENTATION_REL_NS}/slideMaster`,
  theme: `${PRESENTATION_REL_NS}/theme`,
} as const

type ThemeData = {
  colorScheme: Record<string, string>
  bgFillStyles: Element[]
}

const COLOR_NODE_NAMES = ['srgbClr', 'schemeClr', 'sysClr'] as const
const COLOR_MAP_FALLBACK: Record<string, string> = {
  tx1: 'dk1',
  tx2: 'dk2',
  bg1: 'lt1',
  bg2: 'lt2',
}

const sortSlideEntries = (a: string, b: string) => {
  const aIndex = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || '0')
  const bIndex = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || '0')
  return aIndex - bIndex
}

const resolvePartTarget = (sourcePartPath: string, target: string) => {
  const sourceSegments = sourcePartPath.split('/')
  sourceSegments.pop()

  const resolvedSegments = sourceSegments
    .concat(target.split('/'))
    .reduce<string[]>((segments, segment) => {
      if (!segment || segment === '.') return segments
      if (segment === '..') {
        segments.pop()
        return segments
      }
      segments.push(segment)
      return segments
    }, [])

  return resolvedSegments.join('/')
}

const getDirectChild = (element: Element | null, localName: string) => {
  if (!element) return null
  return Array.from(element.children).find(child => child.localName === localName) || null
}

const getChildByPath = (element: Element | null, path: string[]) => {
  let current: Element | null = element
  for (const name of path) {
    current = getDirectChild(current, name)
    if (!current) return null
  }
  return current
}

const getRelationshipTarget = (relsDoc: Document | null, relType: string) => {
  if (!relsDoc) return null

  const relationships = Array.from(relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship'))
  const match = relationships.find(item => item.getAttribute('Type') === relType)
  return match?.getAttribute('Target') || null
}

const getElementAttributes = (element: Element | null) => {
  if (!element) return null

  return Array.from(element.attributes).reduce<Record<string, string>>((attrs, attr) => {
    attrs[attr.name] = attr.value
    return attrs
  }, {})
}

const getThemeColorScheme = (themeRoot: Element) => {
  const colorScheme = getChildByPath(themeRoot, ['themeElements', 'clrScheme'])
  if (!colorScheme) return {}

  return Array.from(colorScheme.children).reduce<Record<string, string>>((scheme, child) => {
    const srgb = getDirectChild(child, 'srgbClr')?.getAttribute('val')
    const sys = getDirectChild(child, 'sysClr')?.getAttribute('lastClr')
    const color = srgb || sys
    if (color) scheme[`a:${child.localName}`] = `#${color}`
    return scheme
  }, {})
}

const getThemeBgFillStyles = (themeRoot: Element) => {
  const bgFillStyleList = getChildByPath(themeRoot, ['themeElements', 'fmtScheme', 'bgFillStyleLst'])
  if (!bgFillStyleList) return []

  return Array.from(bgFillStyleList.children).sort((a, b) => {
    const aOrder = Number(a.getAttribute('order') || 0)
    const bOrder = Number(b.getAttribute('order') || 0)
    return aOrder - bOrder
  })
}

const getColorNode = (element: Element | null) => {
  if (!element) return null
  if (COLOR_NODE_NAMES.includes(element.localName as typeof COLOR_NODE_NAMES[number])) return element
  return Array.from(element.children).find(child =>
    COLOR_NODE_NAMES.includes(child.localName as typeof COLOR_NODE_NAMES[number])
  ) || null
}

const getMappedSchemeName = (schemeName: string, colorMap?: Record<string, string> | null) => {
  if (colorMap && colorMap[schemeName]) return colorMap[schemeName]
  return COLOR_MAP_FALLBACK[schemeName] || schemeName
}

const getColorMap = (
  slideRoot: Element | null,
  layoutRoot: Element | null,
  masterRoot: Element | null,
  level: 'slide' | 'layout' | 'master',
) => {
  if (level === 'slide') {
    return (
      getElementAttributes(getChildByPath(slideRoot, ['clrMapOvr', 'overrideClrMapping'])) ||
      getElementAttributes(getChildByPath(layoutRoot, ['clrMapOvr', 'overrideClrMapping'])) ||
      getElementAttributes(getChildByPath(masterRoot, ['clrMap']))
    )
  }

  if (level === 'layout') {
    return (
      getElementAttributes(getChildByPath(layoutRoot, ['clrMapOvr', 'overrideClrMapping'])) ||
      getElementAttributes(getChildByPath(masterRoot, ['clrMap']))
    )
  }

  return getElementAttributes(getChildByPath(masterRoot, ['clrMap']))
}

const resolveColor = (
  source: Element | null,
  themeData: ThemeData | null,
  colorMap?: Record<string, string> | null,
  phColor?: string | null,
) => {
  const colorNode = getColorNode(source)
  if (!colorNode) return null

  if (colorNode.localName === 'srgbClr') {
    const color = colorNode.getAttribute('val')
    return color ? `#${color}` : null
  }

  if (colorNode.localName === 'sysClr') {
    const color = colorNode.getAttribute('lastClr')
    return color ? `#${color}` : null
  }

  if (colorNode.localName !== 'schemeClr') return null

  const schemeName = colorNode.getAttribute('val')
  if (!schemeName) return null
  if (schemeName === 'phClr') return phColor || null

  const mappedName = getMappedSchemeName(schemeName, colorMap)
  return themeData?.colorScheme[`a:${mappedName}`] || null
}

const resolveSolidBackgroundFromBgPr = (
  bgPr: Element | null,
  themeData: ThemeData | null,
  colorMap?: Record<string, string> | null,
) => {
  const solidFill = getDirectChild(bgPr, 'solidFill')
  if (!solidFill) return null

  const color = resolveColor(solidFill, themeData, colorMap)
  if (!color) return null

  return {
    type: 'solid',
    color,
  } as SlideBackground
}

const resolveSolidBackgroundFromBgRef = (
  bgRef: Element | null,
  themeData: ThemeData | null,
  colorMap?: Record<string, string> | null,
) => {
  if (!bgRef || !themeData) return null

  const idx = Number(bgRef.getAttribute('idx') || 0)
  if (!idx || idx <= 1000) return null

  const phColor = resolveColor(bgRef, themeData, colorMap)
  const fillStyle = themeData.bgFillStyles[idx - 1001]
  if (!fillStyle || fillStyle.localName !== 'solidFill') return null

  const color = resolveColor(fillStyle, themeData, colorMap, phColor)
  if (!color) return null

  return {
    type: 'solid',
    color,
  } as SlideBackground
}

const getBackgroundElements = (root: Element | null) => {
  const bg = getChildByPath(root, ['cSld', 'bg'])
  return {
    bgPr: getDirectChild(bg, 'bgPr'),
    bgRef: getDirectChild(bg, 'bgRef'),
  }
}

const resolveBackgroundFromRoot = (
  root: Element | null,
  level: 'slide' | 'layout' | 'master',
  themeData: ThemeData | null,
  slideRoot: Element | null,
  layoutRoot: Element | null,
  masterRoot: Element | null,
) => {
  if (!root) return null

  const { bgPr, bgRef } = getBackgroundElements(root)
  const colorMap = getColorMap(slideRoot, layoutRoot, masterRoot, level)

  return (
    resolveSolidBackgroundFromBgPr(bgPr, themeData, colorMap) ||
    resolveSolidBackgroundFromBgRef(bgRef, themeData, colorMap)
  )
}

/**
 * 读取 PPTX 中每页实际使用主题后的纯色背景，用于修正导入时的主题偏差。
 */
export const getSlideBackgroundOverrideMap = async (file: File) => {
  const buffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)

  const xmlCache = new Map<string, Document | null>()
  const themeCache = new Map<string, ThemeData | null>()

  const getXmlDocument = async (path: string) => {
    if (xmlCache.has(path)) return xmlCache.get(path) || null

    const xml = await zip.file(path)?.async('text')
    if (!xml) {
      xmlCache.set(path, null)
      return null
    }

    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    xmlCache.set(path, doc)
    return doc
  }

  const getThemeData = async (themePath: string | null) => {
    if (!themePath) return null
    if (themeCache.has(themePath)) return themeCache.get(themePath) || null

    const themeDoc = await getXmlDocument(themePath)
    const themeRoot = themeDoc?.documentElement || null
    if (!themeRoot) {
      themeCache.set(themePath, null)
      return null
    }

    const themeData = {
      colorScheme: getThemeColorScheme(themeRoot),
      bgFillStyles: getThemeBgFillStyles(themeRoot),
    }
    themeCache.set(themePath, themeData)
    return themeData
  }

  const slidePaths = Object.keys(zip.files)
    .filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort(sortSlideEntries)

  const backgrounds: (SlideBackground | null)[] = []

  for (const slidePath of slidePaths) {
    const slideDoc = await getXmlDocument(slidePath)
    const slideRoot = slideDoc?.documentElement || null
    if (!slideRoot) {
      backgrounds.push(null)
      continue
    }

    const slideRelsPath = slidePath.replace(/\/([^/]+)\.xml$/, '/_rels/$1.xml.rels')
    const slideRelsDoc = await getXmlDocument(slideRelsPath)
    const slideLayoutTarget = getRelationshipTarget(slideRelsDoc, REL_TYPES.slideLayout)
    const slideLayoutPath = slideLayoutTarget ? resolvePartTarget(slidePath, slideLayoutTarget) : null

    const layoutDoc = slideLayoutPath ? await getXmlDocument(slideLayoutPath) : null
    const layoutRoot = layoutDoc?.documentElement || null

    const layoutRelsPath = slideLayoutPath?.replace(/\/([^/]+)\.xml$/, '/_rels/$1.xml.rels')
    const layoutRelsDoc = layoutRelsPath ? await getXmlDocument(layoutRelsPath) : null
    const slideMasterTarget = getRelationshipTarget(layoutRelsDoc, REL_TYPES.slideMaster)
    const slideMasterPath = slideMasterTarget && slideLayoutPath
      ? resolvePartTarget(slideLayoutPath, slideMasterTarget)
      : null

    const masterDoc = slideMasterPath ? await getXmlDocument(slideMasterPath) : null
    const masterRoot = masterDoc?.documentElement || null

    const masterRelsPath = slideMasterPath?.replace(/\/([^/]+)\.xml$/, '/_rels/$1.xml.rels')
    const masterRelsDoc = masterRelsPath ? await getXmlDocument(masterRelsPath) : null
    const themeTarget = getRelationshipTarget(masterRelsDoc, REL_TYPES.theme)
    const themePath = themeTarget && slideMasterPath
      ? resolvePartTarget(slideMasterPath, themeTarget)
      : null

    const themeData = await getThemeData(themePath)

    const background =
      resolveBackgroundFromRoot(slideRoot, 'slide', themeData, slideRoot, layoutRoot, masterRoot) ||
      resolveBackgroundFromRoot(layoutRoot, 'layout', themeData, slideRoot, layoutRoot, masterRoot) ||
      resolveBackgroundFromRoot(masterRoot, 'master', themeData, slideRoot, layoutRoot, masterRoot)

    backgrounds.push(background)
  }

  return backgrounds
}
