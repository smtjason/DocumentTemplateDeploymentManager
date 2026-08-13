import JSZip from "jszip";

const namespacePattern = /urn:microsoft-crm\/document-template\/([a-zA-Z][a-zA-Z0-9_]*)\/([0-9]+)\//gi;

function bytesFromBase64(value: string): Uint8Array {
  const binary = atob(value); const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
function base64FromBytes(bytes: Uint8Array): string {
  let binary = ""; const size = 0x8000;
  for (let i = 0; i < bytes.length; i += size) binary += String.fromCharCode(...bytes.subarray(i, i + size));
  return btoa(binary);
}
function namespaces(xml: string, table?: string): string[] {
  const values: string[] = []; namespacePattern.lastIndex = 0;
  for (const match of xml.matchAll(namespacePattern)) if (!table || match[1].toLowerCase() === table.toLowerCase()) values.push(match[0]);
  return [...new Set(values)];
}
async function parts(content: string) {
  const zip = await JSZip.loadAsync(bytesFromBase64(content));
  const item = zip.file("customXml/item1.xml"); const document = zip.file("word/document.xml");
  if (!item || !document) throw new Error("DOCX is missing customXml/item1.xml or word/document.xml.");
  return { zip, itemXml: await item.async("string"), documentXml: await document.async("string") };
}
export async function associatedTable(content: string): Promise<string> {
  const { itemXml } = await parts(content); const tables = [...new Set(namespaces(itemXml).map(value => { namespacePattern.lastIndex = 0; return namespacePattern.exec(value)![1].toLowerCase(); }))];
  if (tables.length !== 1) throw new Error(`Expected one associated table in the DOCX; found ${tables.length}.`);
  return tables[0];
}
export async function rebind(content: string, table: string, objectTypeCode: number): Promise<{ content: string; bindings: number }> {
  const { zip, itemXml, documentXml } = await parts(content); const source = namespaces(itemXml, table);
  if (source.length !== 1) throw new Error(`Expected one CRM namespace for '${table}'; found ${source.length}.`);
  const target = `urn:microsoft-crm/document-template/${table}/${objectTypeCode}/`;
  if (!itemXml.includes(source[0]) || !documentXml.includes(source[0])) throw new Error("Required source namespace was not found in both DOCX parts.");
  zip.file("customXml/item1.xml", itemXml.split(source[0]).join(target));
  zip.file("word/document.xml", documentXml.split(source[0]).join(target));
  const output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const rebound = base64FromBytes(output); const validation = await validate(rebound, table, objectTypeCode);
  if (!validation.valid) throw new Error(validation.message);
  return { content: rebound, bindings: validation.bindings };
}
export async function validate(content: string, table: string, objectTypeCode: number) {
  const { itemXml, documentXml } = await parts(content); const expected = `urn:microsoft-crm/document-template/${table}/${objectTypeCode}/`;
  const distinct = namespaces(`${itemXml}\n${documentXml}`, table); const bindings = (documentXml.match(/<w:dataBinding\b/g) ?? []).length;
  const occurrences = `${itemXml}\n${documentXml}`.split(expected).length - 1;
  const valid = distinct.length === 1 && distinct[0] === expected && bindings > 0 && occurrences === bindings + 1;
  return { valid, bindings, message: valid ? "Valid" : `Expected namespace '${expected}' in one custom XML root plus all bindings; found ${occurrences} occurrences, ${bindings} bindings, and ${distinct.length} distinct namespaces.` };
}
