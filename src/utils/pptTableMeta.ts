import JSZip from 'jszip'

export interface PPTTableCellLayout {
  valign?: 'top' | 'middle' | 'bottom'
  margin?: [number, number, number, number]
}

export type PPTTableLayout = PPTTableCellLayout[][]
export type PPTSlideTableLayouts = PPTTableLayout[]

const EMU_PER_PT = 12700

const mapTableAnchor = (anchor?: string): PPTTableCellLayout['valign'] => {
  if (!anchor) return undefined
  if (anchor === 'ctr') return 'middle'
  if (anchor === 'b') return 'bottom'
  if (anchor === 't') return 'top'
  return undefined
}

export const getTableCellLayoutMap = async (file: File, ratio: number) => {
  const buffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)

  const slideEntries = Object.keys(zip.files)
    .filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort((a, b) => {
      const aIndex = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || '0')
      const bIndex = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || '0')
      return aIndex - bIndex
    })

  const tableLayoutMap: PPTSlideTableLayouts[] = []

  for (const slidePath of slideEntries) {
    const slideXml = await zip.file(slidePath)?.async('text')
    if (!slideXml) {
      tableLayoutMap.push([])
      continue
    }

    const parser = new DOMParser()
    const xmlDoc = parser.parseFromString(slideXml, 'application/xml')
    const graphicFrames = Array.from(xmlDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'graphicFrame'))

    const slideTables: PPTSlideTableLayouts = []
    for (const frame of graphicFrames) {
      const table = frame.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tbl')[0]
      if (!table) continue

      const rows = Array.from(table.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tr'))
      const tableRows: PPTTableLayout = rows.map(row => {
        const cells = Array.from(row.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tc'))
        return cells.map(cell => {
          const tcPr = cell.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'tcPr')[0]
          if (!tcPr) return {}

          const margin = [
            Number(tcPr.getAttribute('marL') || 0) / EMU_PER_PT * ratio,
            Number(tcPr.getAttribute('marT') || 0) / EMU_PER_PT * ratio,
            Number(tcPr.getAttribute('marR') || 0) / EMU_PER_PT * ratio,
            Number(tcPr.getAttribute('marB') || 0) / EMU_PER_PT * ratio,
          ] as [number, number, number, number]

          return {
            valign: mapTableAnchor(tcPr.getAttribute('anchor') || undefined),
            margin: margin.some(item => item) ? margin.map(item => +item.toFixed(2)) as [number, number, number, number] : undefined,
          }
        })
      })

      slideTables.push(tableRows)
    }

    tableLayoutMap.push(slideTables)
  }

  return tableLayoutMap
}
