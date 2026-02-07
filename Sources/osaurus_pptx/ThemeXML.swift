import Foundation

// MARK: - Theme XML Generation

enum ThemeXMLGenerator {

  static func generateThemeXML(theme: Theme) -> String {
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme xmlns:a="\(OOXML.nsA)" name="\(xmlEscape(theme.name))">
      <a:themeElements>
        <a:clrScheme name="\(xmlEscape(theme.name))">
          <a:dk1><a:srgbClr val="\(parseHexColor(theme.textColor))"/></a:dk1>
          <a:lt1><a:srgbClr val="\(parseHexColor(theme.backgroundColor))"/></a:lt1>
          <a:dk2><a:srgbClr val="\(parseHexColor(theme.textColor))"/></a:dk2>
          <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
          <a:accent1><a:srgbClr val="\(parseHexColor(theme.primaryColor))"/></a:accent1>
          <a:accent2><a:srgbClr val="\(parseHexColor(theme.secondaryColor))"/></a:accent2>
          <a:accent3><a:srgbClr val="\(parseHexColor(theme.accentColor1))"/></a:accent3>
          <a:accent4><a:srgbClr val="\(parseHexColor(theme.accentColor2))"/></a:accent4>
          <a:accent5><a:srgbClr val="\(parseHexColor(theme.accentColor3))"/></a:accent5>
          <a:accent6><a:srgbClr val="\(parseHexColor(theme.accentColor4))"/></a:accent6>
          <a:hlink><a:srgbClr val="\(parseHexColor(theme.primaryColor))"/></a:hlink>
          <a:folHlink><a:srgbClr val="\(parseHexColor(theme.secondaryColor))"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="\(xmlEscape(theme.name))">
          <a:majorFont>
            <a:latin typeface="\(xmlEscape(theme.fontHeading))"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="\(xmlEscape(theme.fontBody))"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="\(xmlEscape(theme.name))">
          <a:fillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
              </a:gsLst>
              <a:lin ang="16200000" scaled="1"/>
            </a:gradFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
              </a:gsLst>
              <a:lin ang="16200000" scaled="0"/>
            </a:gradFill>
          </a:fillStyleLst>
          <a:lnStyleLst>
            <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
            <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
            <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
          </a:lnStyleLst>
          <a:effectStyleLst>
            <a:effectStyle><a:effectLst/></a:effectStyle>
            <a:effectStyle><a:effectLst/></a:effectStyle>
            <a:effectStyle><a:effectLst/></a:effectStyle>
          </a:effectStyleLst>
          <a:bgFillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
              </a:gsLst>
              <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
            </a:gradFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
              </a:gsLst>
              <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
            </a:gradFill>
          </a:bgFillStyleLst>
        </a:fmtScheme>
      </a:themeElements>
      <a:objectDefaults/>
      <a:extraClrSchemeLst/>
    </a:theme>
    """
  }

  // MARK: - Slide Master XML

  static func generateSlideMasterXML(theme: Theme, layoutCount: Int = 1) -> String {
    var layoutRels = ""
    for i in 1...layoutCount {
      layoutRels += "    <p:sldLayoutId id=\"\(2_147_483_649 + i)\" r:id=\"rId\(i)\"/>\n"
    }

    return """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sldMaster xmlns:a="\(OOXML.nsA)" xmlns:r="\(OOXML.nsR)" xmlns:p="\(OOXML.nsP)">
        <p:cSld>
          <p:bg>
            <p:bgRef idx="1001">
              <a:schemeClr val="bg1"/>
            </p:bgRef>
          </p:bg>
          <p:spTree>
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
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
        <p:sldLayoutIdLst>
      \(layoutRels)  </p:sldLayoutIdLst>
      </p:sldMaster>
      """
  }

  // MARK: - Slide Master Relationships

  static func generateSlideMasterRels(layoutCount: Int = 1, themeRId: String = "rId100") -> String {
    var rels = ""
    for i in 1...layoutCount {
      rels +=
        "  <Relationship Id=\"rId\(i)\" Type=\"\(OOXML.relTypeSlideLayout)\" Target=\"../slideLayouts/slideLayout\(i).xml\"/>\n"
    }
    rels +=
      "  <Relationship Id=\"\(themeRId)\" Type=\"\(OOXML.relTypeTheme)\" Target=\"../theme/theme1.xml\"/>\n"

    return """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="\(OOXML.nsRelationships)">
      \(rels)</Relationships>
      """
  }

  // MARK: - Slide Layout XML

  static func generateSlideLayoutXML(layoutType: SlideLayoutType) -> String {
    let typeName: String
    switch layoutType {
    case .blank: typeName = "Blank"
    case .title: typeName = "Title Slide"
    case .titleContent: typeName = "Title and Content"
    case .sectionHeader: typeName = "Section Header"
    case .twoContent: typeName = "Two Content"
    case .titleOnly: typeName = "Title Only"
    }

    return """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sldLayout xmlns:a="\(OOXML.nsA)" xmlns:r="\(OOXML.nsR)" xmlns:p="\(OOXML.nsP)" type="\(ooxmlLayoutType(layoutType))" preserve="1">
        <p:cSld name="\(xmlEscape(typeName))">
          <p:spTree>
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
          </p:spTree>
        </p:cSld>
        <p:clrMapOvr>
          <a:masterClrMapping/>
        </p:clrMapOvr>
      </p:sldLayout>
      """
  }

  // MARK: - Slide Layout Relationships

  static func generateSlideLayoutRels() -> String {
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="\(OOXML.nsRelationships)">
      <Relationship Id="rId1" Type="\(OOXML.relTypeSlideMaster)" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>
    """
  }

  private static func ooxmlLayoutType(_ type: SlideLayoutType) -> String {
    switch type {
    case .blank: return "blank"
    case .title: return "title"
    case .titleContent: return "obj"
    case .sectionHeader: return "secHead"
    case .twoContent: return "twoObj"
    case .titleOnly: return "titleOnly"
    }
  }
}
