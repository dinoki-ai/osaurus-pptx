import Foundation

// MARK: - Chart XML Generation

enum ChartXMLGenerator {

  static let defaultColors = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
    "264478", "9B4A16", "636363", "997300", "335B82", "3F6B2B",
  ]

  static func generateChartXML(chart: ChartElement) -> String {
    let titleXML: String
    if let title = chart.chartTitle {
      titleXML = """
        \(generateTitleXML(title))
            <c:autoTitleDeleted val="0"/>
        """
    } else {
      titleXML = "<c:autoTitleDeleted val=\"1\"/>"
    }

    let plotAreaXML: String
    switch chart.chartType {
    case .bar:
      plotAreaXML = generateBarChartXML(chart: chart, horizontal: true)
    case .column:
      plotAreaXML = generateBarChartXML(chart: chart, horizontal: false)
    case .line:
      plotAreaXML = generateLineChartXML(chart: chart)
    case .pie:
      plotAreaXML = generatePieChartXML(chart: chart)
    case .doughnut:
      plotAreaXML = generateDoughnutChartXML(chart: chart)
    }

    let legendXML =
      chart.showLegend
      ? """
        <c:legend>
          <c:legendPos val="b"/>
          <c:overlay val="0"/>
        </c:legend>
      """ : ""

    return """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <c:chartSpace xmlns:c="\(OOXML.nsChart)" xmlns:a="\(OOXML.nsA)" xmlns:r="\(OOXML.nsR)">
        <c:chart>
          \(titleXML)
          <c:plotArea>
            <c:layout/>
      \(plotAreaXML)
          </c:plotArea>
      \(legendXML)
          <c:plotVisOnly val="1"/>
        </c:chart>
      </c:chartSpace>
      """
  }

  // MARK: - Title

  private static func generateTitleXML(_ title: String) -> String {
    """
    <c:title>
        <c:tx>
          <c:rich>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" sz="1400" b="0"/>
                <a:t>\(xmlEscape(title))</a:t>
              </a:r>
            </a:p>
          </c:rich>
        </c:tx>
        <c:overlay val="0"/>
      </c:title>
    """
  }

  // MARK: - Category Axis Reference

  private static func generateCatRef(categories: [String]) -> String {
    var catXML =
      "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$\(categories.count + 1)</c:f><c:strCache><c:ptCount val=\"\(categories.count)\"/>"
    for (i, cat) in categories.enumerated() {
      catXML += "<c:pt idx=\"\(i)\"><c:v>\(xmlEscape(cat))</c:v></c:pt>"
    }
    catXML += "</c:strCache></c:strRef></c:cat>"
    return catXML
  }

  // MARK: - Value Reference

  private static func generateValRef(values: [Double], seriesIndex: Int, count: Int) -> String {
    let colLetter = Character(UnicodeScalar(66 + seriesIndex)!)  // B, C, D...
    var valXML =
      "<c:val><c:numRef><c:f>Sheet1!$\(colLetter)$2:$\(colLetter)$\(count + 1)</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"\(values.count)\"/>"
    for (i, val) in values.enumerated() {
      valXML += "<c:pt idx=\"\(i)\"><c:v>\(val)</c:v></c:pt>"
    }
    valXML += "</c:numCache></c:numRef></c:val>"
    return valXML
  }

  // MARK: - Data Labels

  private static func dataLabelsXML(show: Bool) -> String {
    """
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="\(show ? "1" : "0")"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
    """
  }

  // MARK: - Bar/Column Chart

  private static func generateBarChartXML(chart: ChartElement, horizontal: Bool) -> String {
    var seriesXML = ""

    for (idx, series) in chart.series.enumerated() {
      let color = series.color ?? defaultColors[idx % defaultColors.count]
      seriesXML += """
              <c:ser>
                <c:idx val="\(idx)"/>
                <c:order val="\(idx)"/>
                <c:tx><c:strRef><c:f>Sheet1!$\(Character(UnicodeScalar(66 + idx)!))$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>\(xmlEscape(series.name))</c:v></c:pt></c:strCache></c:strRef></c:tx>
                <c:spPr>
                  <a:solidFill>\(srgbClrXML(color))</a:solidFill>
                </c:spPr>
                \(generateCatRef(categories: chart.categories))
                \(generateValRef(values: series.values, seriesIndex: idx, count: chart.categories.count))
              </c:ser>

        """
    }

    return """
            <c:barChart>
              <c:barDir val="\(horizontal ? "bar" : "col")"/>
              <c:grouping val="clustered"/>
              <c:varyColors val="0"/>
      \(seriesXML)
              \(dataLabelsXML(show: chart.showDataLabels))
              <c:axId val="111111111"/>
              <c:axId val="222222222"/>
            </c:barChart>
            <c:catAx>
              <c:axId val="111111111"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/>
              <c:axPos val="\(horizontal ? "l" : "b")"/>
              <c:crossAx val="222222222"/>
            </c:catAx>
            <c:valAx>
              <c:axId val="222222222"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/>
              <c:axPos val="\(horizontal ? "b" : "l")"/>
              <c:crossAx val="111111111"/>
            </c:valAx>
      """
  }

  // MARK: - Line Chart

  private static func generateLineChartXML(chart: ChartElement) -> String {
    var seriesXML = ""

    for (idx, series) in chart.series.enumerated() {
      let color = series.color ?? defaultColors[idx % defaultColors.count]
      seriesXML += """
              <c:ser>
                <c:idx val="\(idx)"/>
                <c:order val="\(idx)"/>
                <c:tx><c:strRef><c:f>Sheet1!$\(Character(UnicodeScalar(66 + idx)!))$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>\(xmlEscape(series.name))</c:v></c:pt></c:strCache></c:strRef></c:tx>
                <c:spPr>
                  <a:ln w="28575">
                    <a:solidFill>\(srgbClrXML(color))</a:solidFill>
                  </a:ln>
                </c:spPr>
                <c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>
                \(generateCatRef(categories: chart.categories))
                \(generateValRef(values: series.values, seriesIndex: idx, count: chart.categories.count))
                <c:smooth val="0"/>
              </c:ser>

        """
    }

    return """
            <c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
      \(seriesXML)
              \(dataLabelsXML(show: chart.showDataLabels))
              <c:marker val="1"/>
              <c:axId val="111111111"/>
              <c:axId val="222222222"/>
            </c:lineChart>
            <c:catAx>
              <c:axId val="111111111"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/>
              <c:axPos val="b"/>
              <c:crossAx val="222222222"/>
            </c:catAx>
            <c:valAx>
              <c:axId val="222222222"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/>
              <c:axPos val="l"/>
              <c:crossAx val="111111111"/>
            </c:valAx>
      """
  }

  // MARK: - Pie Chart

  private static func generatePieChartXML(chart: ChartElement) -> String {
    guard let series = chart.series.first else { return "" }

    var dataPointsXML = ""
    for (idx, _) in series.values.enumerated() {
      let color = defaultColors[idx % defaultColors.count]
      dataPointsXML += """
                <c:dPt>
                  <c:idx val="\(idx)"/>
                  <c:spPr><a:solidFill>\(srgbClrXML(color))</a:solidFill></c:spPr>
                </c:dPt>

        """
    }

    let dLblsXML: String
    if chart.showDataLabels {
      dLblsXML = """
            <c:dLbls>
              <c:showLegendKey val="0"/>
              <c:showVal val="0"/>
              <c:showCatName val="1"/>
              <c:showSerName val="0"/>
              <c:showPercent val="1"/>
              <c:showBubbleSize val="0"/>
            </c:dLbls>
        """
    } else {
      dLblsXML = dataLabelsXML(show: false)
    }

    return """
            <c:pieChart>
              <c:varyColors val="1"/>
              <c:ser>
                <c:idx val="0"/>
                <c:order val="0"/>
                <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>\(xmlEscape(series.name))</c:v></c:pt></c:strCache></c:strRef></c:tx>
      \(dataPointsXML)
                \(generateCatRef(categories: chart.categories))
                \(generateValRef(values: series.values, seriesIndex: 0, count: chart.categories.count))
              </c:ser>
              \(dLblsXML)
            </c:pieChart>
      """
  }

  // MARK: - Doughnut Chart

  private static func generateDoughnutChartXML(chart: ChartElement) -> String {
    guard let series = chart.series.first else { return "" }

    var dataPointsXML = ""
    for (idx, _) in series.values.enumerated() {
      let color = defaultColors[idx % defaultColors.count]
      dataPointsXML += """
                <c:dPt>
                  <c:idx val="\(idx)"/>
                  <c:spPr><a:solidFill>\(srgbClrXML(color))</a:solidFill></c:spPr>
                </c:dPt>

        """
    }

    return """
            <c:doughnutChart>
              <c:varyColors val="1"/>
              <c:ser>
                <c:idx val="0"/>
                <c:order val="0"/>
                <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>\(xmlEscape(series.name))</c:v></c:pt></c:strCache></c:strRef></c:tx>
      \(dataPointsXML)
                \(generateCatRef(categories: chart.categories))
                \(generateValRef(values: series.values, seriesIndex: 0, count: chart.categories.count))
              </c:ser>
              \(dataLabelsXML(show: chart.showDataLabels))
              <c:holeSize val="50"/>
            </c:doughnutChart>
      """
  }

  // MARK: - Chart Relationships

  static func generateChartRels() -> String {
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="\(OOXML.nsRelationships)">
    </Relationships>
    """
  }
}
