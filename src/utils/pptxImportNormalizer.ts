import JSZip from 'jszip'

const GROUP_TRANSFORM_XML_PATH = /^ppt\/(slides|slideLayouts|slideMasters)\/.*\.xml$/
const GROUP_TRANSFORM_REGEX = /<a:xfrm>([\s\S]*?)<\/a:xfrm>/g

const patchMissingGroupOffset = (xml: string) => {
  let changed = false

  const normalizedXml = xml.replace(GROUP_TRANSFORM_REGEX, (match, innerXml: string) => {
    if (!innerXml.includes('<a:chExt') || innerXml.includes('<a:chOff')) return match

    changed = true
    const patchedInnerXml = innerXml.replace('<a:chExt', '<a:chOff x="0" y="0" /><a:chExt')
    return `<a:xfrm>${patchedInnerXml}</a:xfrm>`
  })

  return {
    changed,
    xml: normalizedXml,
  }
}

export const normalizePPTXImportBuffer = async (buffer: ArrayBuffer) => {
  const zip = await JSZip.loadAsync(buffer)
  let changed = false

  for (const path of Object.keys(zip.files)) {
    if (!GROUP_TRANSFORM_XML_PATH.test(path)) continue

    const file = zip.file(path)
    if (!file) continue

    const xml = await file.async('text')
    const normalized = patchMissingGroupOffset(xml)
    if (!normalized.changed) continue

    zip.file(path, normalized.xml)
    changed = true
  }

  if (!changed) return { buffer, changed: false }

  return {
    buffer: await zip.generateAsync({
      type: 'arraybuffer',
      compression: 'DEFLATE',
    }),
    changed: true,
  }
}
