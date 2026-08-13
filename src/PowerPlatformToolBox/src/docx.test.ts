import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { associatedTable, rebind, validate } from "./docx";

async function fixture(table = "sample_request", typeCode = 10001, bindings = 2) {
  const ns = `urn:microsoft-crm/document-template/${table}/${typeCode}/`;
  const zip = new JSZip();
  zip.file("customXml/item1.xml", `<root xmlns:a="${ns}"><a:name/></root>`);
  const controls = Array.from({ length: bindings }, (_, i) => `<w:sdtPr><w:dataBinding w:prefixMappings="xmlns:a='${ns}'" w:xpath="/a:root/a:f${i}"/></w:sdtPr>`).join("");
  zip.file("word/document.xml", `<w:document xmlns:w="word">${controls || `<w:tag w:val="${ns}"/>`}</w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array" }); let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

describe("DOCX rebinding", () => {
  it("reads the associated table and replaces every CRM namespace", async () => {
    const source = await fixture(); expect(await associatedTable(source)).toBe("sample_request");
    const moved = await rebind(source, "sample_request", 20002);
    expect(moved.bindings).toBe(2); expect((await validate(moved.content, "sample_request", 20002)).valid).toBe(true);
    expect((await validate(moved.content, "sample_request", 10001)).valid).toBe(false);
  });
  it("rejects packages without Word bindings", async () => {
    const source = await fixture("sample_request", 10001, 0);
    await expect(rebind(source, "sample_request", 20002)).rejects.toThrow("0 bindings");
  });
});
