import Foundation

// MARK: - Slide XML Generation

enum SlideXMLGenerator {

    /// Generate complete slide XML
    static func generateSlideXML(slide: Slide, presentation: Presentation) -> String {
        var shapeIndex = 2  // Start at 2 (1 is reserved for the spTree group)
        var spTreeContent = ""

        for element in slide.elements {
            shapeIndex += 1
            if let text = element as? TextElement {
                spTreeContent += generateTextBoxXML(text, shapeId: shapeIndex)
            } else if let image = element as? ImageElement {
                if let rId = image.rId {
                    spTreeContent += generateImageXML(image, shapeId: shapeIndex, rId: rId)
                }
            } else if let shape = element as? ShapeElement {
                spTreeContent += generateShapeXML(shape, shapeId: shapeIndex)
            } else if let table = element as? TableElement {
                spTreeContent += generateTableXML(table, shapeId: shapeIndex, slideWidth: presentation.slideWidth)
            } else if let chart = element as? ChartElement {
                if let rId = chart.rId {
                    spTreeContent += generateChartFrameXML(chart, shapeId: shapeIndex, rId: rId)
                }
            }
        }

        let bgXML = generateBackgroundXML(slide.background)

        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="\(OOXML.nsA)" xmlns:r="\(OOXML.nsR)" xmlns:p="\(OOXML.nsP)">
          <p:cSld>
            \(bgXML)<p:spTree>
              <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
              </p:nvGrpSpPr>
              <p:grpSpPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="0" cy="0"/>
                  <a:chOff x="0" y="0"/>
                  <a:chExt cx="0" cy="0"/>
                </a:xfrm>
              </p:grpSpPr>
        \(spTreeContent)    </p:spTree>
          </p:cSld>
          <p:clrMapOvr>
            <a:masterClrMapping/>
          </p:clrMapOvr>
        </p:sld>
        """
    }

    // MARK: - Background XML

    static func generateBackgroundXML(_ background: SlideBackground?) -> String {
        guard let bg = background else { return "" }

        switch bg.type {
        case .solid(let color):
            return """
                <p:bg>
                  <p:bgPr>
                    <a:solidFill>
                      \(srgbClrXML(color))
                    </a:solidFill>
                    <a:effectLst/>
                  </p:bgPr>
                </p:bg>

            """
        case .gradient(let color1, let color2, let angle):
            let angleVal = Int(angle * 60000)
            return """
                <p:bg>
                  <p:bgPr>
                    <a:gradFill>
                      <a:gsLst>
                        <a:gs pos="0">\(srgbClrXML(color1))</a:gs>
                        <a:gs pos="100000">\(srgbClrXML(color2))</a:gs>
                      </a:gsLst>
                      <a:lin ang="\(angleVal)" scaled="1"/>
                    </a:gradFill>
                    <a:effectLst/>
                  </p:bgPr>
                </p:bg>

            """
        }
    }

    // MARK: - Text Box XML

    static func generateTextBoxXML(_ text: TextElement, shapeId: Int) -> String {
        let rotAttr = text.rotation.map { " rot=\"\(Int($0 * 60000))\"" } ?? ""

        let vertAlignAttr: String
        switch text.verticalAlignment {
        case .top: vertAlignAttr = " anchor=\"t\""
        case .middle: vertAlignAttr = " anchor=\"ctr\""
        case .bottom: vertAlignAttr = " anchor=\"b\""
        }

        let wrapAttr = text.wordWrap ? " wrap=\"square\"" : " wrap=\"none\""

        // Split text by newlines to create multiple paragraphs
        let paragraphs = text.text.components(separatedBy: "\n")
        var paragraphsXML = ""

        for para in paragraphs {
            let bulletXML = text.bullets ? "<a:buChar char=\"&#x2022;\"/>" : "<a:buNone/>"

            let lineSpacingXML: String
            if let spacing = text.lineSpacing {
                lineSpacingXML = "<a:lnSpc><a:spcPts val=\"\(Units.pointsToHundredths(spacing))\"/></a:lnSpc>"
            } else {
                lineSpacingXML = ""
            }

            paragraphsXML += """
                        <a:p>
                          <a:pPr algn="\(text.alignment.rawValue)">\(bulletXML)\(lineSpacingXML)</a:pPr>
                          <a:r>
                            <a:rPr lang="en-US" sz="\(Units.pointsToHundredths(text.fontSize))" b="\(text.bold ? "1" : "0")" i="\(text.italic ? "1" : "0")" u="\(text.underline ? "sng" : "none")" dirty="0">
                              <a:solidFill>\(srgbClrXML(text.fontColor))</a:solidFill>
                              <a:latin typeface="\(xmlEscape(text.fontFace))"/>
                              <a:cs typeface="\(xmlEscape(text.fontFace))"/>
                            </a:rPr>
                            <a:t>\(xmlEscape(para))</a:t>
                          </a:r>
                        </a:p>

            """
        }

        return """
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="\(shapeId)" name="TextBox \(shapeId)"/>
                  <p:cNvSpPr txBox="1"/>
                  <p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm\(rotAttr)>
                    <a:off x="\(text.position.xEMU)" y="\(text.position.yEMU)"/>
                    <a:ext cx="\(text.position.widthEMU)" cy="\(text.position.heightEMU)"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                  <a:noFill/>
                </p:spPr>
                <p:txBody>
                  <a:bodyPr\(wrapAttr)\(vertAlignAttr) rtlCol="0"/>
                  <a:lstStyle/>
        \(paragraphsXML)          </p:txBody>
              </p:sp>

        """
    }

    // MARK: - Image XML

    static func generateImageXML(_ image: ImageElement, shapeId: Int, rId: String) -> String {
        """
              <p:pic>
                <p:nvPicPr>
                  <p:cNvPr id="\(shapeId)" name="Image \(shapeId)"/>
                  <p:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                  </p:cNvPicPr>
                  <p:nvPr/>
                </p:nvPicPr>
                <p:blipFill>
                  <a:blip r:embed="\(rId)"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </p:blipFill>
                <p:spPr>
                  <a:xfrm>
                    <a:off x="\(image.position.xEMU)" y="\(image.position.yEMU)"/>
                    <a:ext cx="\(image.position.widthEMU)" cy="\(image.position.heightEMU)"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                </p:spPr>
              </p:pic>

        """
    }

    // MARK: - Shape XML

    static func generateShapeXML(_ shape: ShapeElement, shapeId: Int) -> String {
        let rotAttr = shape.rotation.map { " rot=\"\(Int($0 * 60000))\"" } ?? ""

        let fillXML: String
        if let fillColor = shape.fillColor {
            fillXML = "<a:solidFill>\(srgbClrXML(fillColor))</a:solidFill>"
        } else {
            fillXML = "<a:noFill/>"
        }

        let lineXML: String
        if let borderColor = shape.borderColor {
            let widthEMU = Units.pointsToEMU(shape.borderWidth)
            lineXML = "<a:ln w=\"\(widthEMU)\"><a:solidFill>\(srgbClrXML(borderColor))</a:solidFill></a:ln>"
        } else {
            lineXML = "<a:ln><a:noFill/></a:ln>"
        }

        let textXML: String
        if let text = shape.text {
            textXML = """
                    <p:txBody>
                      <a:bodyPr wrap="square" anchor="ctr" rtlCol="0"/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr algn="ctr"/>
                        <a:r>
                          <a:rPr lang="en-US" sz="\(Units.pointsToHundredths(shape.textSize))" dirty="0">
                            <a:solidFill>\(srgbClrXML(shape.textColor))</a:solidFill>
                          </a:rPr>
                          <a:t>\(xmlEscape(text))</a:t>
                        </a:r>
                      </a:p>
                    </p:txBody>
            """
        } else {
            textXML = """
                    <p:txBody>
                      <a:bodyPr rtlCol="0"/>
                      <a:lstStyle/>
                      <a:p><a:endParaRPr lang="en-US"/></a:p>
                    </p:txBody>
            """
        }

        return """
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr id="\(shapeId)" name="Shape \(shapeId)"/>
                  <p:cNvSpPr/>
                  <p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm\(rotAttr)>
                    <a:off x="\(shape.position.xEMU)" y="\(shape.position.yEMU)"/>
                    <a:ext cx="\(shape.position.widthEMU)" cy="\(shape.position.heightEMU)"/>
                  </a:xfrm>
                  <a:prstGeom prst="\(shape.shapeType.ooxmlPreset)">
                    <a:avLst/>
                  </a:prstGeom>
                  \(fillXML)
                  \(lineXML)
                </p:spPr>
        \(textXML)
              </p:sp>

        """
    }

    // MARK: - Table XML

    static func generateTableXML(_ table: TableElement, shapeId: Int, slideWidth: Int) -> String {
        guard !table.rows.isEmpty else { return "" }

        let colCount = table.rows.map { $0.count }.max() ?? 0
        guard colCount > 0 else { return "" }

        // Calculate column widths
        let tableWidthEMU = table.position.widthEMU
        let colWidths: [Int]
        if let customWidths = table.columnWidths, customWidths.count == colCount {
            colWidths = customWidths.map { Units.inchesToEMU($0) }
        } else {
            let evenWidth = tableWidthEMU / colCount
            colWidths = Array(repeating: evenWidth, count: colCount)
        }

        let rowHeight = table.position.heightEMU / table.rows.count

        // Build merge lookup
        var mergeMap: [String: MergedCell] = [:]
        var skipCells: Set<String> = []
        for merge in table.mergedCells {
            mergeMap["\(merge.row),\(merge.col)"] = merge
            for r in merge.row..<(merge.row + merge.rowSpan) {
                for c in merge.col..<(merge.col + merge.colSpan) {
                    if r != merge.row || c != merge.col {
                        skipCells.insert("\(r),\(c)")
                    }
                }
            }
        }

        // Build grid columns
        var gridColsXML = ""
        for w in colWidths {
            gridColsXML += "            <a:gridCol w=\"\(w)\"/>\n"
        }

        // Build rows
        var rowsXML = ""
        for (rowIdx, row) in table.rows.enumerated() {
            let isHeader = table.hasHeader && rowIdx == 0
            let isAlternate = !isHeader && rowIdx % 2 == 1

            rowsXML += "          <a:tr h=\"\(rowHeight)\">\n"

            for colIdx in 0..<colCount {
                let cellKey = "\(rowIdx),\(colIdx)"

                if skipCells.contains(cellKey) {
                    // Merged away cell
                    rowsXML += "            <a:tc hMerge=\"1\" vMerge=\"1\"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></a:txBody><a:tcPr/></a:tc>\n"
                    continue
                }

                let cellValue = colIdx < row.count ? row[colIdx] : ""
                let textColor = isHeader ? table.headerTextColor : "333333"

                var mergeAttrs = ""
                if let merge = mergeMap[cellKey] {
                    if merge.colSpan > 1 { mergeAttrs += " gridSpan=\"\(merge.colSpan)\"" }
                    if merge.rowSpan > 1 { mergeAttrs += " rowSpan=\"\(merge.rowSpan)\"" }
                }

                let bgColor: String
                if isHeader {
                    bgColor = table.headerColor
                } else if isAlternate, let altColor = table.alternateRowColor {
                    bgColor = altColor
                } else {
                    bgColor = "FFFFFF"
                }

                rowsXML += """
                            <a:tc\(mergeAttrs)>
                              <a:txBody>
                                <a:bodyPr/>
                                <a:lstStyle/>
                                <a:p>
                                  <a:r>
                                    <a:rPr lang="en-US" sz="\(Units.pointsToHundredths(table.fontSize))" b="\(isHeader ? "1" : "0")" dirty="0">
                                      <a:solidFill>\(srgbClrXML(textColor))</a:solidFill>
                                      <a:latin typeface="\(xmlEscape(table.fontFace))"/>
                                    </a:rPr>
                                    <a:t>\(xmlEscape(cellValue))</a:t>
                                  </a:r>
                                </a:p>
                              </a:txBody>
                              <a:tcPr>
                                <a:solidFill>\(srgbClrXML(bgColor))</a:solidFill>
                                <a:lnL w="12700"><a:solidFill>\(srgbClrXML(table.borderColor))</a:solidFill></a:lnL>
                                <a:lnR w="12700"><a:solidFill>\(srgbClrXML(table.borderColor))</a:solidFill></a:lnR>
                                <a:lnT w="12700"><a:solidFill>\(srgbClrXML(table.borderColor))</a:solidFill></a:lnT>
                                <a:lnB w="12700"><a:solidFill>\(srgbClrXML(table.borderColor))</a:solidFill></a:lnB>
                              </a:tcPr>
                            </a:tc>

                """
            }

            rowsXML += "          </a:tr>\n"
        }

        return """
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="\(shapeId)" name="Table \(shapeId)"/>
                  <p:cNvGraphicFramePr>
                    <a:graphicFrameLocks noGrp="1"/>
                  </p:cNvGraphicFramePr>
                  <p:nvPr/>
                </p:nvGraphicFramePr>
                <p:xfrm>
                  <a:off x="\(table.position.xEMU)" y="\(table.position.yEMU)"/>
                  <a:ext cx="\(table.position.widthEMU)" cy="\(table.position.heightEMU)"/>
                </p:xfrm>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <a:tbl>
                      <a:tblPr firstRow="\(table.hasHeader ? "1" : "0")" bandRow="1">
                        <a:tblStyle val="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>
                      </a:tblPr>
                      <a:tblGrid>
        \(gridColsXML)              </a:tblGrid>
        \(rowsXML)            </a:tbl>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>

        """
    }

    // MARK: - Chart Frame XML

    static func generateChartFrameXML(_ chart: ChartElement, shapeId: Int, rId: String) -> String {
        """
              <p:graphicFrame>
                <p:nvGraphicFramePr>
                  <p:cNvPr id="\(shapeId)" name="Chart \(shapeId)"/>
                  <p:cNvGraphicFramePr>
                    <a:graphicFrameLocks noGrp="1"/>
                  </p:cNvGraphicFramePr>
                  <p:nvPr/>
                </p:nvGraphicFramePr>
                <p:xfrm>
                  <a:off x="\(chart.position.xEMU)" y="\(chart.position.yEMU)"/>
                  <a:ext cx="\(chart.position.widthEMU)" cy="\(chart.position.heightEMU)"/>
                </p:xfrm>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                    <c:chart xmlns:c="\(OOXML.nsChart)" xmlns:r="\(OOXML.nsR)" r:id="\(rId)"/>
                  </a:graphicData>
                </a:graphic>
              </p:graphicFrame>

        """
    }
}

// MARK: - Slide Relationship

struct SlideRelationship {
    let rId: String
    let type: String
    let target: String
}

