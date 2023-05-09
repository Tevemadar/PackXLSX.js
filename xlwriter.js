const xml_head = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const ct_fn = `[Content_Types].xml`;
const ct_dt = `${xml_head}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;
const rels_fn = `_rels/.rels`;
const rels_dt = `${xml_head}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
const xlrels_fn = `xl/_rels/workbook.xml.rels`;
const xlrels_dt = `${xml_head}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;
const wb_fn = `xl/workbook.xml`;
const wb_dt = `${xml_head}
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="r1"/>
</sheets>
</workbook>`;
const sh_fn = `xl/worksheets/sheet1.xml`;
//const sh_dt=`${xml_head}
//<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//<sheetData>
//<row>
//<c r="A1" t="s"><v>0</v></c>
//<c r="C1" t="s"><v>1</v></c>
//</row>
//<row/>
//<row>
//<c r="A3"><v>1.1000000000000001</v></c>
//<c r="B3"><v>2.2000000000000002</v></c>
//<c r="C3"><f>A3*B3</f></c>
//</row>
//</sheetData>
//</worksheet>`;
const sh_head = `${xml_head}
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>`;
const sh_tail = `</sheetData>
</worksheet>`;
const st_fn = `xl/sharedStrings.xml`;
//const st_dt=`${xml_head}
//<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//<si><t>Hello</t></si>
//<si><t>eh</t></si>
//</sst>`;
const st_head = `${xml_head}
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;
const st_tail = `</sst>`;

//function doThing() {
//    const te = new TextEncoder();
//    const date = new Date;
//    const blob = zipstore([
//        {name:ct_fn, date, data:te.encode(ct_dt)},
//        {name:rels_fn, date, data:te.encode(rels_dt)},
//        {name:xlrels_fn, date, data:te.encode(xlrels_dt)},
//        {name:wb_fn, date, data:te.encode(wb_dt)},
//        {name:sh_fn, date, data:te.encode(sh_dt)},
//        {name:st_fn, date, data:te.encode(st_dt)}
//    ]);
//    const url = URL.createObjectURL(blob);
//    const a = document.createElement("a");
//    a.href = url;
//    a.download = "test.xlsx";
//    a.click();
//    URL.revokeObjectURL(url);
//}

function doXl(sheet) {
    const string_table = [];
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
                    sheet_data += `" t="s"><v>${string_table.length}</v>`;
                    string_table.push(cell);
                }
                sheet_data += "</c>";
            }
        sheet_data += "</row>";
    }
    const date = new Date;
    const te = new TextEncoder();
    return zipstore([
        {name: ct_fn, date, data: te.encode(ct_dt)},
        {name: rels_fn, date, data: te.encode(rels_dt)},
        {name: xlrels_fn, date, data: te.encode(xlrels_dt)},
        {name: wb_fn, date, data: te.encode(wb_dt)},
        {name: sh_fn, date, data: te.encode(sh_head + sheet_data + sh_tail)},
        {name: st_fn, date, data: te.encode(st_head + string_table.map(s => `<si><t>${s}</t></si>`).join("") + st_tail)}
    ]);
}
