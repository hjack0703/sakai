// Render Fluent Emoji Flat (MIT) icons to PNG data URIs for use as simple illustrations
const sharp = require("sharp");
const { getIconData, iconToSVG, iconToHTML, replaceIDs } = require("@iconify/utils");
const set = require("@iconify-json/fluent-emoji-flat/icons.json");
const cache = new Map();
async function illus(name, px = 512) {
  if (cache.has(name)) return cache.get(name);
  const data = getIconData(set, name);
  if (!data) throw new Error("no illustration " + name);
  const r = iconToSVG(data, { height: px, width: px });
  const svg = iconToHTML(replaceIDs(r.body), r.attributes);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  const uri = "image/png;base64," + buf.toString("base64");
  cache.set(name, uri);
  return uri;
}
module.exports = { illus };
