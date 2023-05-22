const xml_head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ct_fn = "[Content_Types].xml";
const ct_head = `${xml_head}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;
const ct_tail = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\
</Types>';
const rels_fn = "_rels/.rels";
const rels_dt = `${xml_head}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
const xlrels_fn = "xl/_rels/workbook.xml.rels";
const xlrels_head = `${xml_head}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="r0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
const xlrels_tail = "</Relationships>";
const wb_fn = "xl/workbook.xml";
const wb_head = `${xml_head}
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
<sheets>`;
const wb_tail = "</sheets>\
</workbook>";
const sh_head = `${xml_head}
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>`;
const sh_tail = "</sheetData>\
</worksheet>";
const st_fn = "xl/sharedStrings.xml";
const st_head = `${xml_head}
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;
const st_tail = "</sst>";

function packxlsx(sheets) {
    let ct_dt = ct_head;
    let xlrels_dt = xlrels_head;
    let wb_dt = wb_head;

    const string_table = [];
    const sheet_xmls = [];

    for (let i = 0; i < sheets.length; i++) {
        ct_dt += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        xlrels_dt += `<Relationship Id="r${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
        let sheet = sheets[i];
        if (typeof sheet[0] === "string") {
            wb_dt += `<sheet name="${sheet[0]}" sheetId="${i + 1}" r:id="r${i + 1}"/>`;
            sheet = sheet.slice(1);
        } else {
            wb_dt += `<sheet name="Sheet${i + 1}" sheetId="${i + 1}" r:id="r${i + 1}"/>`;
        }

        let sheet_data = "";
        for (let y = 0; y < sheet.length; y++) {
            sheet_data += "<row>";
            const row = sheet[y];
            if (row)
                for (let x = 0; x < row.length; x++) {
                    const cell = row[x];
                    const typ = typeof cell;
                    const isnum = typ === "number";
                    const isstr = typ === "string";
                    if (!isnum && !isstr)
                        continue;
                    sheet_data += "<c r=\"";
                    if (x >= 26)
                        sheet_data += String.fromCharCode(64 + Math.floor(x / 26));
                    sheet_data += String.fromCharCode(65 + x % 26);
                    sheet_data += y + 1;
                    if (isnum)
                        sheet_data += `"><v>${cell}</v>`;
                    else if (cell.startsWith("="))
                        sheet_data += `"><f>${cell.substring(1)}</f>`;
                    else {
                        let idx = string_table.indexOf(cell);
                        if (idx === -1) {
                            idx = string_table.length;
                            string_table.push(cell);
                        }
                        sheet_data += `" t="s"><v>${idx}</v>`;
                    }
                    sheet_data += "</c>";
                }
            sheet_data += "</row>";
        }
        sheet_xmls.push(sh_head + sheet_data + sh_tail);
    }
    ct_dt += ct_tail;
    xlrels_dt += xlrels_tail;
    wb_dt += wb_tail;
    const date = new Date;
    const te = new TextEncoder();
    return zipstore([
        {name: ct_fn, date, data: te.encode(ct_dt)},
        {name: rels_fn, date, data: te.encode(rels_dt)},
        {name: xlrels_fn, date, data: te.encode(xlrels_dt)},
        {name: wb_fn, date, data: te.encode(wb_dt)},
        ...sheet_xmls.map((xml, idx) => ({name: `xl/worksheets/sheet${idx + 1}.xml`, date, data: te.encode(xml)})),
        {name: st_fn, date, data: te.encode(st_head + string_table.map(s => `<si><t>${s}</t></si>`).join("") + st_tail)}
    ]);
}
