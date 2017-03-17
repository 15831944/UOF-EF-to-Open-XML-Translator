<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
                xmlns:uof="http://schemas.uof.org/cn/2003/uof"
                xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow"
                xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc"
                xmlns:图="http://schemas.uof.org/cn/2003/graph"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <xsl:import href="txBody.xsl"/>
  <xsl:import href="chartArea.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>


  <xsl:template name="prstGeom">
    <a:prstGeom>
      <xsl:if test=".//图:类别|图:名称">
        <xsl:attribute name="prst">
          <xsl:call-template name="prstName"/>
        </xsl:attribute>
        <a:avLst>
          <xsl:if test=".//图:名称='Elbow Connector' or .//图:名称='Elbow Arrow Connector' or .//图:名称='Curved Arrow Connector' or .//图:名称='Elbow Double-Arrow Connector' or .//图:名称='Curved Double-Arrow Connector' or .//图:名称='Curved Connector'">
            <a:gd name="adj1" fmla="val 50000"/>
          </xsl:if>
        </a:avLst>
      </xsl:if>
    </a:prstGeom>
  </xsl:template>
  <xsl:template name="solidFill">
    <a:solidFill>
      <a:srgbClr>
        <xsl:attribute name="val">
          <xsl:choose>
            <xsl:when test=".//图:属性/图:填充/图:颜色!='auto'">
              <xsl:variable name="RGB1">
                <xsl:value-of select=".//图:属性/图:填充/图:颜色"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($RGB1,'#')"/>
            </xsl:when>
            <xsl:otherwise>ffffff</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test=".//图:属性/图:透明度">
          <a:alpha>
            <xsl:attribute name="val">
              <xsl:value-of select="concat(100-.//图:属性/图:透明度,'%')"/>
            </xsl:attribute>
          </a:alpha>
        </xsl:if>
      </a:srgbClr>
    </a:solidFill>
  </xsl:template>
  <xsl:template name="ln">
    <a:ln>
      <xsl:if test=".//图:线粗细">
        <xsl:attribute name="w">
          <xsl:value-of select=".//图:属性/图:线粗细*12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性/图:线型='single'">
        <xsl:attribute name="cmpd">sng</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性/图:线型='double' or .//图:属性/图:线型='thick'">
        <xsl:attribute name="cmpd">dbl</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:线型='none'">
      </xsl:if>
      <xsl:if test=".//图:属性/图:线颜色">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:choose>
                <xsl:when test=".//图:属性/图:线颜色!='auto'">
                  <xsl:variable name="RGB">
                    <xsl:value-of select=".//图:属性/图:线颜色"/>
                  </xsl:variable>
                  <xsl:value-of select="substring-after($RGB,'#')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'000000'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:if>
      <xsl:choose>
        <xsl:when test=".//图:属性/图:线型='dotted'">
          <a:prstDash>
            <xsl:attribute name="val">sysDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dotted-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">sysDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash'">
          <a:prstDash>
            <xsl:attribute name="val">dash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-dot-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">dashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-long'">
          <a:prstDash>
            <xsl:attribute name="val">lgDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-long-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-dot-dot-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDotDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dot-dash'">
          <a:prstDash>
            <xsl:attribute name="val">dashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-dot-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">sysDashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-dot-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDashDotDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='double'">
          <a:prstDash>
            <xsl:attribute name="val">dbl</xsl:attribute>
          </a:prstDash>
        </xsl:when>
      </xsl:choose>
      <xsl:if test=".//图:前端箭头">
        <a:headEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:前端箭头/图:式样='normal'">triangle</xsl:when>
              <xsl:when test=".//图:前端箭头/图:式样='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:前端箭头/图:式样='open'">arrow</xsl:when>
              <xsl:when test=".//图:前端箭头/图:式样='oval'">oval</xsl:when>
              <xsl:when test=".//图:前端箭头/图:式样='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:headEnd>
      </xsl:if>
      <xsl:if test=".//图:后端箭头">
        <a:tailEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:后端箭头/图:式样='normal'">triangle</xsl:when>
              <xsl:when test=".//图:后端箭头/图:式样='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:后端箭头/图:式样='open'">arrow</xsl:when>
              <xsl:when test=".//图:后端箭头/图:式样='oval'">oval</xsl:when>
              <xsl:when test=".//图:后端箭头/图:式样='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:tailEnd>
      </xsl:if>
    </a:ln>
  </xsl:template>
  <xsl:template name="rot">
    <a:scene3d>
      <a:camera>
        <xsl:attribute name="prst">orthographicFront</xsl:attribute>
        <a:rot>
          <xsl:attribute name="lat">0</xsl:attribute>
          <xsl:attribute name="lon">0</xsl:attribute>
          <xsl:attribute name="rev">
            <xsl:choose>
              <xsl:when test="图:图形/图:预定义图形/图:属性/图:旋转角度!='0.0'">
                <xsl:value-of select="round(21600000 - 图:图形/图:预定义图形/图:属性/图:旋转角度*60000)"/>
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:rot>
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">threePt</xsl:attribute>
        <xsl:attribute name="dir">t</xsl:attribute>
      </a:lightRig>
    </a:scene3d>
  </xsl:template>
  <xsl:template name="style_zwf">
    <xdr:style>
      <a:lnRef idx="1">
        <a:schemeClr val="accent1"/>
      </a:lnRef>
      <a:fillRef idx="0">
        <a:schemeClr val="accent1"/>
      </a:fillRef>
      <a:effectRef idx="0">
        <a:schemeClr val="accent1"/>
      </a:effectRef>
      <a:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </a:fontRef>
    </xdr:style>
  </xsl:template>
  <xsl:template name="prstName">
    <xsl:choose>
      <xsl:when test=".//图:名称='Cloud Callout' or .//图:类别='54'">cloudCallout</xsl:when>
      <xsl:when test=".//图:名称='Rectangle' or .//图:类别='11'">rect</xsl:when>
      <xsl:when test=".//图:名称='Parallelogram' or .//图:类别='12'">parallelogram</xsl:when>
      <xsl:when test=".//图:名称='Trapezoid' or .//图:类别='13'">trapezoid</xsl:when>
      <xsl:when test=".//图:名称='diamond' or .//图:类别='14'">diamond</xsl:when>
      <xsl:when test=".//图:名称='Rounded Rectangle' or .//图:类别='15'">roundRect</xsl:when>
      <xsl:when test=".//图:名称='Octagon'or .//图:类别='16'">octagon</xsl:when>
      <xsl:when test=".//图:名称='Isosceles Triangle' or .//图:类别='17'">triangle</xsl:when>
      <xsl:when test=".//图:名称='Right Triangle' or .//图:类别='18'">rtTriangle</xsl:when>
      <xsl:when test=".//图:名称='Oval' or .//图:类别='19'">ellipse</xsl:when>
      <xsl:when test=".//图:名称='Right Arrow' or .//图:类别='21'">rightArrow</xsl:when>
      <xsl:when test=".//图:名称='Left Arrow' or .//图:类别='22'">leftArrow</xsl:when>
      <xsl:when test=".//图:名称='Up Arrow' or .//图:类别='23'">upArrow</xsl:when>
      <xsl:when test=".//图:名称='Down Arrow' or .//图:类别='24'">downArrow</xsl:when>
      <xsl:when test=".//图:名称='Left-Right Arrow' or .//图:类别='25'">leftRightArrow</xsl:when>
      <xsl:when test=".//图:名称='Up-Down Arrow' or .//图:类别='26'">upDownArrow</xsl:when>
      <xsl:when test=".//图:名称='Quad Arrow' or .//图:类别='27'">quadArrow</xsl:when>
      <xsl:when test=".//图:名称='Left-Right-Up Arrow' or .//图:类别='28'">leftRightUpArrow</xsl:when>
      <xsl:when test=".//图:名称='Bent Arrow' or .//图:类别='29'">bentArrow</xsl:when>
      <xsl:when test=".//图:名称='Process' or .//图:类别='31'">flowChartProcess</xsl:when>
      <xsl:when test=".//图:名称='Alternate Process' or .//图:类别='32'">flowChartAlternateProcess</xsl:when>
      <xsl:when test=".//图:名称='Decision' or .//图:类别='33'">flowChartDecision</xsl:when>
      <xsl:when test=".//图:名称='Data' or .//图:类别='34'">flowChartInputOutput</xsl:when>
      <xsl:when test=".//图:名称='Predefined Process' or .//图:类别='35'">flowChartPredefinedProcess</xsl:when>
      <xsl:when test=".//图:名称='Internal Storage' or .//图:类别='36'">flowChartInternalStorage</xsl:when>
      <xsl:when test=".//图:名称='Document' or .//图:类别='37'">flowChartDocument</xsl:when>
      <xsl:when test=".//图:名称='Multidocument' or .//图:类别='38'">flowChartMultidocument</xsl:when>
      <xsl:when test=".//图:名称='Terminator' or .//图:类别='39'">flowChartTerminator</xsl:when>
      <xsl:when test=".//图:名称='Explosion 1' or .//图:类别='41'">irregularSeal1</xsl:when>
      <xsl:when test=".//图:名称='Explosion 2' or .//图:类别='42'">irregularSeal2</xsl:when>
      <xsl:when test=".//图:名称='4-Point Star' or .//图:类别='43'">star4</xsl:when>
      <xsl:when test=".//图:名称='5-Point Star' or .//图:类别='44'">star5</xsl:when>
      <xsl:when test=".//图:名称='8-Point Star' or .//图:类别='45'">star8</xsl:when>
      <xsl:when test=".//图:名称='16-Point Star' or .//图:类别='46'">star12</xsl:when>
      <xsl:when test=".//图:名称='24-Point Star' or .//图:类别='47'">star24</xsl:when>
      <xsl:when test=".//图:名称='32-Point Star' or .//图:类别='48'">star32</xsl:when>
      <xsl:when test=".//图:名称='Up Ribbon' or .//图:类别='49'">ribbon2</xsl:when>
      <xsl:when test=".//图:名称='Rectangular Callout' or .//图:类别='51'">wedgeRectCallout</xsl:when>
      <xsl:when test=".//图:名称='Rounded Rectangular Callout' or .//图:类别='52'">wedgeRoundRectCallout</xsl:when>
      <xsl:when test=".//图:名称='Oval Callout' or .//图:类别='53'">wedgeEllipseCallout</xsl:when>
      <xsl:when test=".//图:名称='Parallelogram' or .//图:类别='54'">cloudCallout</xsl:when>
      <xsl:when test=".//图:名称='Line Callout1(No Border)' or .//图:类别='513'">borderCallout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout2' or .//图:类别='56'">borderCallout2</xsl:when>
      <xsl:when test=".//图:名称='Line Callout3' or .//图:类别='57'">borderCallout2</xsl:when>
      <xsl:when test=".//图:名称='Line Callout4' or .//图:类别='58'">borderCallout3</xsl:when>
      <xsl:when test=".//图:名称='Line Callout1(Accent Bar)' or .//图:类别='59'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout2(Accent Bar)' or .//图:类别='510'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout3(Accent Bar)' or .//图:类别='511'">accentCallout2</xsl:when>
      <xsl:when test=".//图:名称='Line Callout4(Accent Bar)' or .//图:类别='512'">accentCallout3</xsl:when>
      <xsl:when test=".//图:名称='Line Callout1' or .//图:类别='55'">callout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout2(No Border)' or .//图:类别='514'">callout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout3(No Border)' or .//图:类别='515'">callout2</xsl:when>
      <xsl:when test=".//图:名称='Line Callout4(No Border)' or .//图:类别='516'">callout3</xsl:when>
      <xsl:when test=".//图:名称='Line Callout1(Border and Accent Bar)' or .//图:类别='517'">accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout2(Border and Accent Bar)' or .//图:类别='518'">accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称='Line Callout3(Border and Accent Bar)' or .//图:类别='519'">accentBorderCallout2</xsl:when>
      <xsl:when test=".//图:名称='Line Callout4(Border and Accent Bar)' or .//图:类别='520'">accentBorderCallout3</xsl:when>
      <xsl:when test=".//图:名称='Line' or .//图:类别='61'">line</xsl:when>
      <xsl:when test=".//图:名称='Arrow' or .//图:类别='62'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称='Double Arrow' or .//图:类别='63'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称='Curve' or .//图:类别='64'">curvedConnector3</xsl:when>
      <xsl:when test=".//图:名称='Straight Connector' or .//图:类别='71'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称='Elbow Connector' or .//图:类别='74'">bentConnector3</xsl:when>
      <xsl:when test=".//图:名称='Curved Connector' or .//图:类别='77'">curvedConnector3</xsl:when>
      <xsl:when test=".//图:名称='Hexagon' or .//图:类别='110'">hexagon</xsl:when>
      <xsl:when test=".//图:名称='Cross' or .//图:类别='111'">plus</xsl:when>
      <xsl:when test=".//图:名称='Regual Pentagon' or .//图:类别='112'">pentagon</xsl:when>
      <xsl:when test=".//图:名称='Can' or .//图:类别='113'">can</xsl:when>
      <xsl:when test=".//图:名称='Cube' or .//图:类别='114'">cube</xsl:when>
      <xsl:when test=".//图:名称='Bevel' or .//图:类别='115'">bevel</xsl:when>
      <xsl:when test=".//图:名称='Folded Corner' or .//图:类别='116'">foldedCorner</xsl:when>
      <xsl:when test=".//图:名称='Smiley Face' or .//图:类别='117'">smileyFace</xsl:when>
      <xsl:when test=".//图:名称='Donut' or .//图:类别='118'">donut</xsl:when>
      <xsl:when test=".//图:名称='No Symbol' or .//图:类别='119'">noSmoking</xsl:when>
      <xsl:when test=".//图:名称='Block Arc' or .//图:类别='120'">blockArc</xsl:when>
      <xsl:when test=".//图:名称='Heart' or .//图:类别='121'">heart</xsl:when>
      <xsl:when test=".//图:名称='Lightning' or .//图:类别='122'">lightningBolt</xsl:when>
      <xsl:when test=".//图:名称='Sun' or .//图:类别='123'">sun</xsl:when>
      <xsl:when test=".//图:名称='Moon' or .//图:类别='124'">moon</xsl:when>
      <xsl:when test=".//图:名称='Arc' or .//图:类别='125'">arc</xsl:when>
      <xsl:when test=".//图:名称='Double Bracket' or .//图:类别='126'">bracketPair</xsl:when>
      <xsl:when test=".//图:名称='Double Brace' or .//图:类别='127'">bracePair</xsl:when>
      <xsl:when test=".//图:名称='Plaque' or .//图:类别='128'">plaque</xsl:when>
      <xsl:when test=".//图:名称='Left Bracket' or .//图:类别='129'">leftBracket</xsl:when>
      <xsl:when test=".//图:名称='Right Bracket' or .//图:类别='130'">rightBracket</xsl:when>
      <xsl:when test=".//图:名称='Left Brace' or .//图:类别='131'">leftBrace</xsl:when>
      <xsl:when test=".//图:名称='Right Brace' or .//图:类别='132'">rightBrace</xsl:when>
      <xsl:when test=".//图:名称='U-Turn Arrow' or .//图:类别='210'">uturnArrow</xsl:when>
      <xsl:when test=".//图:名称='Left-Up Arrow' or .//图:类别='211'">leftUpArrow</xsl:when>
      <xsl:when test=".//图:名称='Bent-Up Arrow' or .//图:类别='212'">bentUpArrow</xsl:when>
      <xsl:when test=".//图:名称='Curved Right Arrow' or .//图:类别='213'">curvedRightArrow</xsl:when>
      <xsl:when test=".//图:名称='Curved Left Arrow' or .//图:类别='214'">curvedLeftArrow</xsl:when>
      <xsl:when test=".//图:名称='Curved Up Arrow' or .//图:类别='215'">curvedUpArrow</xsl:when>
      <xsl:when test=".//图:名称='Curved Down Arrow' or .//图:类别='216'">curvedDownArrow</xsl:when>
      <xsl:when test=".//图:名称='Striped Right Arrow' or .//图:类别='217'">stripedRightArrow</xsl:when>
      <xsl:when test=".//图:名称='Notched Right Arrow' or .//图:类别='218'">notchedRightArrow</xsl:when>
      <xsl:when test=".//图:名称='Pentagon Arrow' or .//图:类别='219'">homePlate</xsl:when>
      <xsl:when test=".//图:名称='Chevron Arrow' or .//图:类别='220'">chevron</xsl:when>
      <xsl:when test=".//图:名称='Right Arrow Callout' or .//图:类别='221'">rightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Left Arrow Callout' or .//图:类别='222'">leftArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Up Arrow Callout' or .//图:类别='223'">upArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Down Arrow Callout' or .//图:类别='224'">downArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Left-Right Arrow Callout' or .//图:类别='225'">leftRightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Up-Down Arrow Callout' or .//图:类别='226'">upDownArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Quad Arrow Callout' or .//图:类别='227'">quadArrowCallout</xsl:when>
      <xsl:when test=".//图:名称='Circular Arrow' or .//图:类别='228'">circularArrow</xsl:when>
      <xsl:when test=".//图:名称='Preparation' or .//图:类别='310'">flowChartPreparation</xsl:when>
      <xsl:when test=".//图:名称='Manual Input' or .//图:类别='311'">flowChartManualInput</xsl:when>
      <xsl:when test=".//图:名称='Manual Operation' or .//图:类别='312'">flowChartManualOperation</xsl:when>
      <xsl:when test=".//图:名称='Connector' or .//图:类别='313'">flowChartConnector</xsl:when>
      <xsl:when test=".//图:名称='Off-page Connector' or .//图:类别='314'">flowChartOffpageConnector</xsl:when>
      <xsl:when test=".//图:名称='Card' or .//图:类别='315'">flowChartPunchedCard</xsl:when>
      <xsl:when test=".//图:名称='Punched Tape' or .//图:类别='316'">flowChartPunchedTape</xsl:when>
      <xsl:when test=".//图:名称='Summing Junction' or .//图:类别='317' ">flowChartSummingJunction</xsl:when>
      <xsl:when test=".//图:名称='Or' or .//图:类别='318'">flowChartOr</xsl:when>
      <xsl:when test=".//图:名称='Collate' or .//图:类别='319'">flowChartCollate</xsl:when>
      <xsl:when test=".//图:名称='Sort' or .//图:类别='320'">flowChartSort</xsl:when>
      <xsl:when test=".//图:名称='Extract' or .//图:类别='321'">flowChartExtract</xsl:when>
      <xsl:when test=".//图:名称='Merge' or .//图:类别='322'">flowChartMerge</xsl:when>
      <xsl:when test=".//图:名称='Stored Data' or .//图:类别='323'">flowChartOnlineStorage</xsl:when>
      <xsl:when test=".//图:名称='Delay' or .//图:类别='324'">flowChartDelay</xsl:when>
      <xsl:when test=".//图:名称='Sequential Access Storage' or .//图:类别='325'">flowChartMagneticTape</xsl:when>
      <xsl:when test=".//图:名称='Magnetic Disk' or .//图:类别='326'">flowChartMagneticDisk</xsl:when>
      <xsl:when test=".//图:名称='Direct Access Storage' or .//图:类别='327'">flowChartMagneticDrum</xsl:when>
      <xsl:when test=".//图:名称='Display' or .//图:类别='328'">flowChartDisplay</xsl:when>
      <xsl:when test=".//图:名称='Down Ribbon' or .//图:类别='410'">ribbon</xsl:when>
      <xsl:when test=".//图:名称='Curved Up Ribbon' or .//图:类别='411'">ellipseRibbon2</xsl:when>
      <xsl:when test=".//图:名称='Curved Down Ribbon' or .//图:类别='412'">ellipseRibbon</xsl:when>
      <xsl:when test=".//图:名称='Vertical Scroll' or .//图:类别='413'">verticalScroll</xsl:when>
      <xsl:when test=".//图:名称='Horizontal Scroll' or .//图:类别='414'">horizontalScroll</xsl:when>
      <xsl:when test=".//图:名称='Wave' or .//图:类别='415'">wave</xsl:when>
      <xsl:when test=".//图:名称='Double Wave' or .//图:类别='416'">doubleWave</xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!--2014-3-17，update by Qihy， 属性前面不应该加前缀， start-->
  <xsl:template name="fill">
    <xsl:if test=".//图:颜色 and substring(.//图:颜色,2,6)!='ffffff'">
      <a:solidFill>
        <a:srgbClr>
          <xsl:choose>
            <xsl:when test=".//图:属性/图:填充/图:颜色='auto'">
              <xsl:attribute name="val">ffffff</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="val">
                <xsl:value-of select="substring(.//图:颜色,2,6)"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test=".//图:属性/图:透明度">
            <a:alpha>
              <xsl:attribute name="val">
                <xsl:value-of select="concat(100-.//图:属性/图:透明度,'%')"/>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <xsl:if test="substring(.//图:颜色,2,6)='ffffff'">
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </xsl:if>
    <xsl:if test=".//图:渐变">
      <xsl:apply-templates select=".//图:渐变"/>
    </xsl:if>
    
    <xsl:if test=".//图:图片">
      <xsl:variable name="tttpp">
        <xsl:value-of select="translate(@uof:图形引用,'OBJ','')"/>
      </xsl:variable>
      <a:blipFill>
        <a:blip xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
          <xsl:attribute name="r:embed">
            <xsl:value-of select="concat('rId',$tttpp)"/>
          </xsl:attribute>
        </a:blip>
        <a:srcRect/>
        <xsl:if test=".//图:图片/@位置='stretch'">
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </xsl:if>
      </a:blipFill>
    </xsl:if>
    <xsl:if test=".//图:图案">
      <a:pattFill>
        <xsl:attribute name="prst">
          <xsl:call-template name="fillName"/>
        </xsl:attribute>
        <a:fgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring(.//图:图案/@前景色,2,6)"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:fgClr>
        <a:bgClr>
          
          <!--<a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring(.//图:图案/@图:背景色,2,6)"/>
            </xsl:attribute>
          </a:srgbClr>-->
          <xsl:choose>
            <xsl:when test ="contains(.//图:图案/@背景色, '#')">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring(.//图:图案/@背景色,2,6)"/>
                </xsl:attribute>
              </a:srgbClr>
            </xsl:when>
            <xsl:otherwise>
              <a:schemeClr>
                <xsl:attribute name="val">
                  <xsl:value-of select=".//图:图案/@背景色"/>
                </xsl:attribute>
              </a:schemeClr>
            </xsl:otherwise>
          </xsl:choose>
        </a:bgClr>
      </a:pattFill>
    </xsl:if>
  </xsl:template>
  <xsl:template name="fillName">
    <xsl:choose>
      <xsl:when test=".//图:图案/@类型='ptn001'">pct5</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn002'">pct50</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn003'">ltDnDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn004'">ltVert</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn005'">wave</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn006'">zigZag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn007'">divot</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn008'">smGrid</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn009'">pct10</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn010'">pct60</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn011'">ltUpDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn012'">ltHorz</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn013'">upDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn014'">wave</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn015'">dotGrid</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn016'">lgGrid</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn017'">pct20</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn018'">pct70</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn019'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn020'">narVert</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn021'">horz</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn022'">diagBrick</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn023'">dotDmnd</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn024'">smCheck</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn025'">pct25</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn026'">pct75</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn027'">dkUpDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn028'">narHorz</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn029'">vert</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn030'">horzBrick</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn031'">shingle</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn032'">lgCheck</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn033'">pct30</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn034'">pct80</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn035'">wdDnDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn036'">dkVert</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn037'">smConfetti</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn038'">weave</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn039'">trellis</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn040'">openDmnd</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn041'">pct40</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn042'">pct90</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn043'">wdUpDiag</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn044'">dkHorz</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn045'">lgConfetti</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn046'">plaid</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn047'">sphere</xsl:when>
      <xsl:when test=".//图:图案/@类型='ptn048'">solidDmnd</xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--翻转-->
  <xsl:template name="flip">
    <xsl:choose>
      <xsl:when test=".//图:翻转/@方向='y'">
        <xsl:attribute name="flipH">1</xsl:attribute>
      </xsl:when>
      <xsl:when test=".//图:翻转/@方向='x'">
        <xsl:attribute name="flipV">1</xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
<!--end-->
