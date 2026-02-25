import { readFile, writeFile } from "node:fs/promises";

/**
 * Fix duplicate xmlns declarations in the Activity element.
 * The writer sometimes emits two sets of xmlns attributes — deduplicate them.
 */
export async function fixDuplicateXmlns(filePath: string): Promise<void> {
  let content = await readFile(filePath, "utf-8");

  const match = content.match(/(<Activity\s)([\s\S]*?)(>)/);
  if (!match) return;

  const attrsStr = match[2];
  const attrPairs = attrsStr.match(/[\w:]+="[^"]*"/g);
  if (!attrPairs) return;

  // Deduplicate by attribute name (keep last occurrence for completeness)
  const seen = new Map<string, string>();
  for (const attr of attrPairs) {
    const name = attr.split("=")[0];
    seen.set(name, attr);
  }

  // Check if there were duplicates
  if (seen.size === attrPairs.length) return;

  // Rebuild attribute string with stable ordering
  const xmlnsDefault = seen.get("xmlns");
  seen.delete("xmlns");
  const mcIgnorable = seen.get("mc:Ignorable");
  seen.delete("mc:Ignorable");
  const xClass = seen.get("x:Class");
  seen.delete("x:Class");

  const xmlnsAttrs: string[] = [];
  const otherAttrs: string[] = [];
  for (const [key, val] of seen) {
    if (key.startsWith("xmlns:")) xmlnsAttrs.push(val);
    else otherAttrs.push(val);
  }
  xmlnsAttrs.sort();
  otherAttrs.sort();

  const parts: string[] = [];
  if (xmlnsDefault) parts.push(xmlnsDefault);
  parts.push(...xmlnsAttrs);
  if (mcIgnorable) parts.push(mcIgnorable);
  if (xClass) parts.push(xClass);
  parts.push(...otherAttrs);

  const newTag = `<Activity ${parts.join(" ")}>`;
  content = content.substring(0, match.index!) + newTag + content.substring(match.index! + match[0].length);
  await writeFile(filePath, content, "utf-8");
}

/**
 * Try to parse a JSON string. Returns the parsed object on success,
 * or throws with a descriptive message on failure.
 */
export function tryParseJson(text: string): any {
  try {
    return JSON.parse(text);
  } catch (e) {
    const msg = (e as Error).message;
    throw new Error(`Invalid JSON: ${msg}`);
  }
}
