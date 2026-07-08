// Native pie-chart injection for the generated report.
//
// ExcelJS (4.4.0) can embed images but cannot emit native Excel charts. The
// client's reference (Shamir) workbook shows live pie charts on the result
// sheet that recalculate as rows are reclassified, so we reproduce that by
// post-processing the workbook buffer: ExcelJS writes the sheets/formulas, then
// this module opens the resulting .xlsx (a zip) and injects the OOXML chart,
// drawing, relationship and content-type parts by hand.
//
// Nothing here changes any cell value, formula, validation or layout — the
// charts only *reference* ranges that already exist on the sheet.
import JSZip from "jszip";

export type ChartAnchor = {
  // 0-based worksheet cell coordinates for the chart's top-left / bottom-right.
  fromCol: number;
  fromRow: number;
  toCol: number;
  toRow: number;
};

export type CategoryPieChart = {
  title: string;
  // Excel range refs, already sheet-qualified and $-anchored, e.g.
  //   'תוצאות השיקוף'!$A$4:$A$9   (subcategory labels)
  //   'תוצאות השיקוף'!$B$4:$B$9   (subcategory values)
  catRef: string;
  valRef: string;
  // "category" (default) = small pie, percent-only labels + legend below.
  // "summary" = large pie, category-name + percent labels around the slices,
  // no legend — used for the two big expenses/income overview pies.
  variant?: "category" | "summary";
  // Explicit placement. When omitted the chart flows into the auto grid.
  anchor?: ChartAnchor;
};

// Teal → coral palette derived from the report theme (reportTheme in
// finance-analyzer). Slice colors cycle through this list so every pie reads as
// the same product as the rest of the workbook.
const slicePalette = [
  "124559", "1B6E8C", "2A9D8F", "3A7CA5", "5FA8D3",
  "E76F51", "EA580C", "B5532A", "F4A261", "6B7B83",
];

const titleColor = "124559";

// Grid layout (in worksheet cells) for the pies, placed to the right of the
// existing A–H data block.
const gridColumns = 3;
const firstChartCol = 9; // column J (0-based), clear of the H-column data
const firstChartRow = 22; // row 23 (0-based) — below the two big summary pies at the top
const chartColSpan = 7; // pie occupies 7 columns…
const chartRowSpan = 16; // …and 16 rows
const colStride = chartColSpan + 1; // +1 column gutter
const rowStride = chartRowSpan + 1; // +1 row gutter

// Only & < > require escaping in XML element text. We deliberately leave
// apostrophes/quotes raw so the chart formula refs (e.g. 'תוצאות השיקוף'!$A$4)
// read exactly the way Excel itself writes them.
function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function buildChartXml(chart: CategoryPieChart) {
  const isSummary = chart.variant === "summary";

  // Summary pies label each slice with the category name AND its percent (like
  // the reference overview pies) and drop the legend; category pies show the
  // percent only and rely on a legend below for the names.
  const dataLabels =
    "<c:dLbls>" +
    '<c:numFmt formatCode="0%" sourceLinked="0"/>' +
    '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' +
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"/></a:pPr><a:endParaRPr lang="he-IL"/></a:p></c:txPr>' +
    `<c:dLblPos val="${isSummary ? "outEnd" : "bestFit"}"/>` +
    '<c:showLegendKey val="0"/><c:showVal val="0"/>' +
    `<c:showCatName val="${isSummary ? "1" : "0"}"/>` +
    '<c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>' +
    (isSummary ? "<c:separator>\n</c:separator>" : "") +
    '<c:showLeaderLines val="1"/>' +
    "</c:dLbls>";
  const legend = isSummary ? "" : '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>';

  // One <c:dPt> per palette color gives the slices our themed colors; Excel
  // cycles them across however many data points the range actually has.
  const dataPoints = slicePalette
    .map(
      (color, index) =>
        `<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
        `<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`,
    )
    .join("");

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    "<c:chart>" +
    "<c:title><c:tx><c:rich>" +
    '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>' +
    "<a:lstStyle/>" +
    `<a:p><a:pPr><a:defRPr sz="1200" b="1" i="0" u="none" strike="noStrike" baseline="0">` +
    `<a:solidFill><a:srgbClr val="${titleColor}"/></a:solidFill></a:defRPr></a:pPr>` +
    `<a:r><a:rPr lang="he-IL"/><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p>` +
    "</c:rich></c:tx><c:overlay val=\"0\"/></c:title>" +
    '<c:autoTitleDeleted val="0"/>' +
    "<c:plotArea><c:layout/>" +
    "<c:pieChart><c:varyColors val=\"1\"/>" +
    "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
    dataPoints +
    dataLabels +
    `<c:cat><c:strRef><c:f>${escapeXml(chart.catRef)}</c:f></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>${escapeXml(chart.valRef)}</c:f></c:numRef></c:val>` +
    "</c:ser>" +
    '<c:firstSliceAng val="0"/>' +
    "</c:pieChart>" +
    "</c:plotArea>" +
    legend +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>' +
    "</c:chart></c:chartSpace>"
  );
}

function buildDrawingXml(charts: CategoryPieChart[]) {
  let anchors = "";
  let gridIndex = 0; // counts only auto-placed (anchor-less) charts
  charts.forEach((chart, index) => {
    let fromCol: number;
    let fromRow: number;
    let toCol: number;
    let toRow: number;
    if (chart.anchor) {
      ({ fromCol, fromRow, toCol, toRow } = chart.anchor);
    } else {
      const gridCol = gridIndex % gridColumns;
      const gridRow = Math.floor(gridIndex / gridColumns);
      fromCol = firstChartCol + gridCol * colStride;
      fromRow = firstChartRow + gridRow * rowStride;
      toCol = fromCol + chartColSpan;
      toRow = fromRow + chartRowSpan;
      gridIndex += 1;
    }
    const frameId = index + 2; // 1 is reserved by convention
    anchors +=
      '<xdr:twoCellAnchor editAs="oneCell">' +
      `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff>` +
      `<xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff>` +
      `<xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr>" +
      `<xdr:cNvPr id="${frameId}" name="Chart ${index + 1}"/>` +
      "<xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr>" +
      "</xdr:nvGraphicFramePr>" +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
      ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${index + 1}"/>` +
      "</a:graphicData></a:graphic></xdr:graphicFrame>" +
      "<xdr:clientData/></xdr:twoCellAnchor>";
  });
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    anchors +
    "</xdr:wsDr>"
  );
}

// One drawing owns several charts; its rels map the drawing's local rIds to the
// global chart part numbers (chart parts are numbered across ALL sheets so two
// drawings never point at the same chartN.xml).
function buildDrawingRels(chartNumbers: number[]) {
  let relationships = "";
  chartNumbers.forEach((chartNumber, index) => {
    relationships +=
      `<Relationship Id="rId${index + 1}"` +
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"' +
      ` Target="../charts/chart${chartNumber}.xml"/>`;
  });
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships +
    "</Relationships>"
  );
}

// Resolve the sheet part path (e.g. xl/worksheets/sheet2.xml) for a sheet by its
// display name, by walking workbook.xml → workbook.xml.rels.
function resolveSheetPath(workbookXml: string, workbookRels: string, sheetName: string) {
  const escaped = escapeXml(sheetName);
  const sheetTag = new RegExp(`<sheet[^>]*name="${escaped.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*r:id="(rId\\d+)"`);
  const match = workbookXml.match(sheetTag);
  if (!match) {
    return null;
  }
  const rId = match[1];
  const relTag = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"`);
  const relMatch = workbookRels.match(relTag);
  if (!relMatch) {
    return null;
  }
  const target = relMatch[1].replace(/^\//, "");
  return target.startsWith("xl/") ? target : `xl/${target}`;
}

// A set of charts to inject onto one worksheet. Each sheet gets its own drawing
// part; the charts only reference ranges that already exist on the workbook, so
// they can live on a different sheet than the data they plot.
export type ChartSheetGroup = {
  sheetName: string;
  charts: CategoryPieChart[];
};

export async function injectCategoryPieCharts(
  buffer: Buffer,
  groups: ChartSheetGroup[],
): Promise<Buffer> {
  const activeGroups = groups.filter((group) => group.charts.length > 0);
  if (activeGroups.length === 0) {
    return buffer;
  }

  const zip = await JSZip.loadAsync(buffer);

  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const workbookRels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !workbookRels) {
    return buffer;
  }

  // Chart parts are numbered globally and drawings sequentially, so injecting
  // onto several sheets never collides on chartN.xml / drawingN.xml.
  let chartCounter = 0;
  let drawingCounter = 0;
  let contentOverrides = "";

  for (const group of activeGroups) {
    const sheetPath = resolveSheetPath(workbookXml, workbookRels, group.sheetName);
    if (!sheetPath) {
      continue;
    }
    const sheetXml = await zip.file(sheetPath)?.async("string");
    if (!sheetXml) {
      continue;
    }

    drawingCounter += 1;
    const drawingIndex = drawingCounter;
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;

    // 1. chart parts (globally-unique numbers)
    const chartNumbers = group.charts.map(() => {
      chartCounter += 1;
      return chartCounter;
    });
    group.charts.forEach((chart, index) => {
      zip.file(`xl/charts/chart${chartNumbers[index]}.xml`, buildChartXml(chart));
    });

    // 2. drawing part + its rels
    zip.file(drawingPath, buildDrawingXml(group.charts));
    zip.file(`xl/drawings/_rels/drawing${drawingIndex}.xml.rels`, buildDrawingRels(chartNumbers));

    // 3. sheet → drawing relationship
    const sheetRelsPath = sheetPath.replace(/worksheets\/([^/]+)$/, "worksheets/_rels/$1.rels");
    const existingSheetRels = await zip.file(sheetRelsPath)?.async("string");
    let drawingRelId = "rId1";
    if (existingSheetRels) {
      const ids = [...existingSheetRels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
      const nextId = (ids.length ? Math.max(...ids) : 0) + 1;
      drawingRelId = `rId${nextId}`;
      const updated = existingSheetRels.replace(
        "</Relationships>",
        `<Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/></Relationships>`,
      );
      zip.file(sheetRelsPath, updated);
    } else {
      zip.file(
        sheetRelsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/>` +
          "</Relationships>",
      );
    }

    // 4. <drawing> element on the worksheet (just before </worksheet>, which
    //    puts it after pageSetup — the schema-correct position).
    if (!/<drawing\s/.test(sheetXml)) {
      const updatedSheet = sheetXml.replace(
        "</worksheet>",
        `<drawing r:id="${drawingRelId}"/></worksheet>`,
      );
      zip.file(sheetPath, updatedSheet);
    }

    // 5. accumulate content-type overrides for this group's new parts
    contentOverrides += `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
    chartNumbers.forEach((chartNumber) => {
      contentOverrides += `<Override PartName="/xl/charts/chart${chartNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
    });
  }

  // 6. content-type overrides (written once for every injected part)
  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath)?.async("string");
  if (contentTypes && contentOverrides) {
    const updatedTypes = contentTypes.replace("</Types>", `${contentOverrides}</Types>`);
    zip.file(contentTypesPath, updatedTypes);
  }

  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return out;
}
