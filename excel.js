// Arma un Excel (.xlsx) con varias hojas y gráficos nativos de Excel, sin librerías externas.
// Uso: crearExcel([{ nombre, anchos:[...], filas:[[celda,...],...], grafico:{...} }]) -> Blob
// celda: texto | número | { v, s } donde s es un estilo de ESTILOS.
(function () {
  const ESTILOS = { normal: 0, encabezado: 1, plata: 2, pct: 3, titulo: 4, nota: 5, fecha: 6, hora: 7, entero: 8, plataFuerte: 9, etiqueta: 10 };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const col = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  // Fecha de JavaScript -> número de serie de Excel (hora local)
  function serial(d) { return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569; }

  function celdaXml(c, ref) {
    if (c === null || c === undefined || c === '') return '';
    let v = c, s = 0;
    if (typeof c === 'object' && !(c instanceof Date)) { v = c.v; s = ESTILOS[c.s] || 0; }
    if (v instanceof Date) v = serial(v);
    if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
  }

  function hojaXml(h, conDibujo) {
    const cols = (h.anchos || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    const filas = h.filas.map((f, r) => `<row r="${r + 1}">${(f || []).map((c, i) => celdaXml(c, col(i) + (r + 1))).join('')}</row>`).join('');
    return XML + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_R}">` +
      `<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
      (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${filas}</sheetData>` +
      `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
      (conDibujo ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
  }

  // g: { titulo, hoja, cat:[c,f1,f2], val:[c,f1,f2], horizontal, resaltar:[índices], etiquetas, formato }
  function graficoXml(g) {
    const ref = ([c, f1, f2]) => `'${g.hoja.replace(/'/g, "''")}'!$${col(c)}$${f1 + 1}:$${col(c)}$${f2 + 1}`;
    const relleno = hex => `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>`;
    const fmt = g.formato || '"$ "#,##0';
    const puntos = (g.resaltar || []).map(i => `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>${relleno('16A34A')}</c:dPt>`).join('');
    const etiquetas = g.etiquetas === false
      ? '<c:dLbls><c:delete val="1"/></c:dLbls>'
      : `<c:dLbls><c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` +
        `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="1F2A24"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es-AR"/></a:p></c:txPr>` +
        `<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
    const texto = (sz, color) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es-AR"/></a:p></c:txPr>`;
    const sinLinea = '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>';
    const grilla = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E3EAE6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
    const catAx = `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="${g.horizontal ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/>` +
      `<c:axPos val="${g.horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
      `<c:tickLblPos val="nextTo"/><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="C9D3CE"/></a:solidFill></a:ln></c:spPr>${texto(900, '5B6B63')}` +
      `<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${g.horizontal ? 1 : 0}"/>` +
      `<c:axPos val="${g.horizontal ? 'b' : 'l'}"/>${g.horizontal ? '' : grilla}<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` +
      `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${sinLinea}${texto(900, '5B6B63')}` +
      `<c:crossAx val="1001"/><c:crosses val="${g.horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
    return XML + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">` +
      `<c:roundedCorners val="0"/><c:chart>` +
      `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1"/></a:pPr><a:r><a:rPr lang="es-AR" sz="1300" b="1"><a:solidFill><a:srgbClr val="1F2A24"/></a:solidFill></a:rPr><a:t>${esc(g.titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
      `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>` +
      `<c:barChart><c:barDir val="${g.horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${esc(g.titulo)}</c:v></c:tx>${relleno(g.resaltar ? '8CCFA3' : '16A34A')}<c:invertIfNegative val="0"/>` +
      `${puntos}${etiquetas}<c:cat><c:strRef><c:f>${ref(g.cat)}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${ref(g.val)}</c:f></c:numRef></c:val></c:ser>` +
      `<c:gapWidth val="${g.horizontal ? 50 : 70}"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>${catAx}${valAx}` +
      `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
      `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
  }

  function dibujoXml(g) {
    const [c1, f1, c2, f2] = g.lugar;
    const p = (c, f) => `<xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${f}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
    return XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<xdr:twoCellAnchor><xdr:from>${p(c1, f1)}</xdr:from><xdr:to>${p(c2, f2)}</xdr:to>` +
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${esc(g.titulo)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
      `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${NS_R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>` +
      `<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  }

  const ESTILOS_XML = XML + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$ &quot;#,##0"/></numFmts>` +
    `<fonts count="5"><font><sz val="11"/><color rgb="FF1F2A24"/><name val="Calibri"/><family val="2"/></font>` +
    `<font><b/><sz val="11"/><color rgb="FF14532D"/><name val="Calibri"/><family val="2"/></font>` +
    `<font><b/><sz val="16"/><color rgb="FF15803D"/><name val="Calibri"/><family val="2"/></font>` +
    `<font><i/><sz val="10"/><color rgb="FF6B7A73"/><name val="Calibri"/><family val="2"/></font>` +
    `<font><b/><sz val="11"/><color rgb="FF1F2A24"/><name val="Calibri"/><family val="2"/></font></fonts>` +
    `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/><bgColor indexed="64"/></patternFill></fill></fills>` +
    `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>` +
    `<border><left/><right/><top/><bottom style="thin"><color rgb="FF9FD3B2"/></bottom><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="11">` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>` +
    `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
    `<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
    `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
    `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
    `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
    `<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
    `<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
    `<xf numFmtId="164" fontId="4" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>` +
    `<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
    `</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  // ---------- ZIP (sin compresión) ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(archivos) {
    const enc = new TextEncoder(), partes = [], central = []; let off = 0;
    const d = new Date(), hora = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const fecha = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const [nombre, texto] of archivos) {
      const n = enc.encode(nombre), datos = enc.encode(texto), crc = crc32(datos);
      const cab = new DataView(new ArrayBuffer(30));
      [[0, 0x04034b50, 4], [4, 20, 2], [6, 0x0800, 2], [8, 0, 2], [10, hora, 2], [12, fecha, 2], [14, crc, 4], [18, datos.length, 4], [22, datos.length, 4], [26, n.length, 2], [28, 0, 2]]
        .forEach(([p, v, l]) => (l === 4 ? cab.setUint32(p, v, true) : cab.setUint16(p, v, true)));
      const cen = new DataView(new ArrayBuffer(46));
      [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0x0800, 2], [10, 0, 2], [12, hora, 2], [14, fecha, 2], [16, crc, 4], [20, datos.length, 4], [24, datos.length, 4],
        [28, n.length, 2], [30, 0, 2], [32, 0, 2], [34, 0, 2], [36, 0, 2], [38, 0, 4], [42, off, 4]]
        .forEach(([p, v, l]) => (l === 4 ? cen.setUint32(p, v, true) : cen.setUint16(p, v, true)));
      partes.push(new Uint8Array(cab.buffer), n, datos); central.push(new Uint8Array(cen.buffer), n);
      off += 30 + n.length + datos.length;
    }
    const tamCentral = central.reduce((a, x) => a + x.length, 0);
    const fin = new DataView(new ArrayBuffer(22));
    [[0, 0x06054b50, 4], [4, 0, 2], [6, 0, 2], [8, archivos.length, 2], [10, archivos.length, 2], [12, tamCentral, 4], [16, off, 4], [20, 0, 2]]
      .forEach(([p, v, l]) => (l === 4 ? fin.setUint32(p, v, true) : fin.setUint16(p, v, true)));
    return new Blob([...partes, ...central, new Uint8Array(fin.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  window.crearExcel = function (hojas) {
    const arch = [], tipos = [], relsLibro = [];
    let nGraf = 0;
    hojas.forEach((h, i) => {
      const n = i + 1;
      arch.push([`xl/worksheets/sheet${n}.xml`, hojaXml(h, !!h.grafico)]);
      tipos.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
      relsLibro.push(`<Relationship Id="rId${n}" Type="${NS_R}/worksheet" Target="worksheets/sheet${n}.xml"/>`);
      if (h.grafico) {
        const k = ++nGraf, g = Object.assign({ hoja: h.nombre }, h.grafico);
        arch.push([`xl/worksheets/_rels/sheet${n}.xml.rels`, XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/drawing" Target="../drawings/drawing${k}.xml"/></Relationships>`]);
        arch.push([`xl/drawings/drawing${k}.xml`, dibujoXml(g)]);
        arch.push([`xl/drawings/_rels/drawing${k}.xml.rels`, XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/chart" Target="../charts/chart${k}.xml"/></Relationships>`]);
        arch.push([`xl/charts/chart${k}.xml`, graficoXml(g)]);
        tipos.push(`<Override PartName="/xl/drawings/drawing${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
          `<Override PartName="/xl/charts/chart${k}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
      }
    });
    const nEst = hojas.length + 1;
    relsLibro.push(`<Relationship Id="rId${nEst}" Type="${NS_R}/styles" Target="styles.xml"/>`);
    return zip([
      ['[Content_Types].xml', XML + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${tipos.join('')}</Types>`],
      ['_rels/.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ['xl/workbook.xml', XML + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_R}"><bookViews><workbookView/></bookViews><sheets>` +
        hojas.map((h, i) => `<sheet name="${esc(h.nombre)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsLibro.join('')}</Relationships>`],
      ['xl/styles.xml', ESTILOS_XML],
      ...arch,
    ]);
  };
})();
