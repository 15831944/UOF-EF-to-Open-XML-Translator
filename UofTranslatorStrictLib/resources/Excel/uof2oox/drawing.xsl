<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                 xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties" 
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                xmlns:dc="http://purl.org/dc/elements/1.1/" 
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" 
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:图表="http://schemas.uof.org/cn/2009/chart"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:公式="http://schemas.uof.org/cn/2009/equations"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:xsd="http://www.ord.com">
  <xsl:import href="hyperlink.xsl"/>
  <xsl:import href ="txBody.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <!--Modified by LDM in 2011/01/06-->
  <!--坐标值与起终单元格转换模板-->
  <xsl:template name="AxisToCellLocation">
    <xsl:param name="xCoordinate"/>
    <xsl:param name="yCoordinate"/>
    <xsl:param name="width"/>
    <xsl:param name="height"/>
    <xsl:param name="xUnit"/>
    <xsl:param name="yUnit"/>

    <!--如何更好的确定起终单元格的位置-->
    <!--Not Finished-->
    <xsl:variable name="fromCol" select="floor($xCoordinate div $xUnit)"/>
    <xsl:variable name="fromColOff">
      <!--xsl:value-of select="floor(($xCoordinate - $fromCol * $xUnit) * 360000 div 28.3)"/-->
      <xsl:value-of select="$xCoordinate * 1.065"/>
    </xsl:variable>

    <!-- 2014058 update by lingfeng 锚点位置不准 start -->
    <xsl:variable name="fromRow" select="floor($yCoordinate div $yUnit)"/>
    <xsl:variable name="fromRowOff">
      <!--xsl:value-of select="floor(($yCoordinate - $fromRow * $yUnit) * 360000 div 28.3)"/-->
      <xsl:value-of select="$yCoordinate"/>
    </xsl:variable>

    <xsl:variable name="toCol" select="floor(($xCoordinate + $width) div $xUnit)"/>
    <xsl:variable name="toColOff">
      <!--xsl:value-of select="floor((($width - ($toCol - $fromCol) * $xUnit)) * 360000 div 28.3 + $fromColOff)"/-->
      <xsl:value-of select="($xCoordinate + $width) * 1.065"/>
    </xsl:variable>
    <xsl:variable name="toRow" select="floor(($yCoordinate + $height) div $yUnit)"/>
    <xsl:variable name="toRowOff">
      <!--xsl:value-of select="floor((($height - ($toRow - $fromRow) * $yUnit)) * 360000 div 28.3 + $fromRowOff)"/-->
      <xsl:value-of select="$yCoordinate + $height div 1.065"/>
    </xsl:variable>
    <!--end-->
    
    <xdr:from>
      <xdr:col>
        <xsl:choose >
          <xsl:when test ="contains($fromCol,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromCol"/>
          </xsl:otherwise>
        </xsl:choose>
        
      </xdr:col>
      <xdr:colOff>
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        <xsl:value-of select="$fromColOff"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        <xsl:choose >
          <xsl:when test ="contains($fromRow,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromRow"/>
          </xsl:otherwise>
        </xsl:choose>
        
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$fromRowOff"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>
        <xsl:value-of select="$toCol"/>
      </xdr:col>
      <xdr:colOff>
        <!-- 20130506 update by xuzhenwei BUG_2882 回归 互操作oo-uof（编辑）-oo 文件010图表.xlsx转换后的.xlsx文件需要修复 start
        <xsl:choose>
          <xsl:when test="contains($toColOff,'-')">
            <xsl:value-of select ="substring-after($toColOff,'-')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$toColOff"/>
          </xsl:otherwise>
        </xsl:choose>
        end -->
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        
        <!--2014-3-29, update by Qihy, 修复bug3159，文本框大小不一致， start-->
        <!--<xsl:value-of select="$xCoordinate + $width"/>-->
        
        <!--2014-4-24, update by Qihy, 图表转换时元素取值不能为小数， start-->
        <xsl:value-of select="$toColOff"/>
        <!--2014-4-24 end-->
        
        <!--2014-3-29 end-->
        
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        
        <!--2014-5-5, update by Qihy, 修复bug3302，打开需要恢复问题， start-->
        <!--<xsl:value-of select="$toRow"/>-->
        <xsl:choose >
          <xsl:when test ="contains($toRow,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$toRow"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--2014-5-5 end-->
        
      </xdr:row>
      <xdr:rowOff>

        <!--2014-4-24, update by Qihy, 图表转换时元素取值不能为负数， start-->
        <!--<xsl:value-of select="$toRowOff"/>-->
        <xsl:choose >
          <xsl:when test ="contains($toRowOff,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$toRowOff"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--2014-4-24 end-->
        
      </xdr:rowOff>
    </xdr:to>
  </xsl:template>

  <!--只需要xdr:from-->
  <xsl:template name="AxisToCellLocationOneCellAnchor">
    <xsl:param name="xCoordinate"/>
    <xsl:param name="yCoordinate"/>
    <xsl:param name="width"/>
    <xsl:param name="height"/>
    <xsl:param name="xUnit"/>
    <xsl:param name="yUnit"/>

    <!--如何更好的确定起终单元格的位置-->
    <!--Not Finished-->
    <xsl:variable name="fromCol" select="floor($xCoordinate div $xUnit)"/>
    <xsl:variable name="fromColOff">
      <!--xsl:value-of select="floor(($xCoordinate - $fromCol * $xUnit) * 360000 div 28.3)"/-->
      <xsl:value-of select="floor($xCoordinate * 1.065)"/>
    </xsl:variable>

    <!-- 2014058 update by lingfeng 锚点位置不准 start -->
    <xsl:variable name="fromRow" select="floor($yCoordinate div $yUnit)"/>
    <xsl:variable name="fromRowOff">
      <!--xsl:value-of select="floor(($yCoordinate - $fromRow * $yUnit) * 360000 div 28.3)"/-->
      <xsl:value-of select="floor($yCoordinate)"/>
    </xsl:variable>

    <xsl:variable name="toCol" select="floor(($xCoordinate + $width) div $xUnit)"/>
    <xsl:variable name="toColOff">
      <!--xsl:value-of select="floor((($width - ($toCol - $fromCol) * $xUnit)) * 360000 div 28.3 + $fromColOff)"/-->
      <xsl:value-of select="floor(($xCoordinate + $width) * 1.065)"/>
    </xsl:variable>
    <xsl:variable name="toRow" select="floor(($yCoordinate + $height) div $yUnit)"/>
    <xsl:variable name="toRowOff">
      <!--xsl:value-of select="floor((($height - ($toRow - $fromRow) * $yUnit)) * 360000 div 28.3 + $fromRowOff)"/-->
      <xsl:value-of select="floor($yCoordinate + $height div 1.065)"/>
    </xsl:variable>
    <!--end-->

    <xdr:from>
      <xdr:col>
        <xsl:choose >
          <xsl:when test ="contains($fromCol,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromCol"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:col>
      <xdr:colOff>
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        <xsl:value-of select="$fromColOff"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        <xsl:choose >
          <xsl:when test ="contains($fromRow,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromRow"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$fromRowOff"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:ext>
      <xsl:attribute name="cx">
        <xsl:value-of select="floor($height * 12700)"/>
      </xsl:attribute>
      <xsl:attribute name="cy">
        <xsl:value-of select="floor($width * 12700)"/>
      </xsl:attribute>
    </xdr:ext>
  </xsl:template>

  <xsl:template name="equDrawing" >
    <xsl:param name="ref"/>
    <xsl:variable name="xCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="yCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="width">
      <xsl:value-of select="./uof:大小_C621/@宽_C605"/>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:value-of select="./uof:大小_C621/@长_C604"/>
    </xsl:variable>
    <xsl:variable name="sheetNo">
      <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
    </xsl:variable>
    <!--随动方式-->
    <xsl:variable name="moveType">
      <xsl:choose>
        <xsl:when test="./@表:随动方式">
          <xsl:value-of select="./@表:随动方式"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'move and re-size'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xdr:oneCellAnchor>
      <!--xdr:from-->
      <xsl:call-template name="AxisToCellLocationOneCellAnchor">
        <xsl:with-param name="xCoordinate" select="$xCoordinate"/>
        <xsl:with-param name="yCoordinate" select="$yCoordinate"/>
        <xsl:with-param name="width" select="$width"/>
        <xsl:with-param name="height" select="$height"/>
        <xsl:with-param name="xUnit" select="'54'"/>
        <!--72-->
        <xsl:with-param name="yUnit" select="'14'"/>
        <!--19-->
      </xsl:call-template>
      <xsl:variable name="mc">
        &lt;mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"&gt;
      </xsl:variable>
      <mc:AlternateContent>
        <!--<xsl:value-of select ="$mc" disable-output-escaping="yes"/>-->
        <xsl:variable name="choice">&lt;mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"&gt;</xsl:variable>
        <xsl:value-of select="$choice" disable-output-escaping="yes"/>
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="2" name="文本框 1"/>
            <xdr:cNvSpPr txBox="1"/>
          </xdr:nvSpPr>
          <xdr:spPr >
            <a:xfrm>
              <a:off>
                <xsl:attribute name="x">
                  <xsl:value-of select="floor($xCoordinate * 12700)"/>
                </xsl:attribute>
                <xsl:attribute name="y">
                  <xsl:value-of select="floor($yCoordinate * 12700)"/>
                </xsl:attribute>
              </a:off>
              <a:ext>
                <xsl:attribute name="cx">
                  <xsl:value-of select="floor($height * 12700)"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="floor($width * 12700)"/>
                </xsl:attribute>
              </a:ext>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
          </xdr:spPr>
          <xdr:style>
            <a:lnRef idx="0">
              <a:scrgbClr r="0" g="0" b="0"/>
            </a:lnRef>
            <a:fillRef idx="0">
              <a:scrgbClr r="0" g="0" b="0"/>
            </a:fillRef>
            <a:effectRef idx="0">
              <a:scrgbClr r="0" g="0" b="0"/>
            </a:effectRef>
            <a:fontRef idx="minor">
              <a:schemeClr val="tx1"/>
            </a:fontRef>
          </xdr:style>
          <xdr:txBody>
            <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t">
              <a:spAutoFit/>
            </a:bodyPr>
            <a:lstStyle/>
            <a:p>
              <a:pPr/>
              <a14:m >
                <xsl:copy-of select ="./ancestor::uof:UOF/公式:公式集_C200/公式:数学公式_C201[@标识符_C202=$ref]/m:oMathPara"/>
              </a14:m>
            </a:p>
          </xdr:txBody>
        </xdr:sp>
        <xsl:variable name="endChoice">&lt;/mc:Choice&gt;</xsl:variable>
        <xsl:value-of select="$endChoice" disable-output-escaping="yes"/>
        <xsl:variable name="endmc">&lt;/mc:AlternateContent&gt;</xsl:variable>
        <!--<xsl:value-of select="$endmc" disable-output-escaping="yes"/>-->
      </mc:AlternateContent>
      <xdr:clientData />
    </xdr:oneCellAnchor>
  </xsl:template>


  <!--处理填充-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="Fill_Pre">
    <!--处理除白色外的纯色填充-->
    <xsl:if test=".//图:颜色_8004 and substring(.//图:颜色_8004,2,6)!='ffffff'">
      <a:solidFill>
        <a:srgbClr>
          <xsl:choose>
            <xsl:when test=".//图:属性_801D/图:填充_804C/图:颜色_8004='auto'">
              <xsl:attribute name="val">ffffff</xsl:attribute>
              <!--3065bc-->
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="val">
                <xsl:value-of select="substring(.//图:颜色_8004,2,6)"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test=".//图:属性_801D/图:透明度_8050">
            <a:alpha>
              <xsl:attribute name="val">
                
                <!--<xsl:value-of select="(100 - .//图:属性_801D/图:透明度_8050)*1000"/>-->
                <xsl:value-of select="concat(number(100 - number(.//图:属性_801D/图:透明度_8050)),'%')"/>
                
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <xsl:if test="substring(.//图:颜色_8004,2,6)='ffffff'">
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </xsl:if>
    <xsl:if test=".//图:渐变_800D">
      <!--渐变模板在chartArea.xsl文件中，待修改。 李杨2012-2-17-->
      <xsl:apply-templates select=".//图:渐变_800D"/>
    </xsl:if>
    <xsl:if test=".//图:图片_8005">
      <xsl:variable name="tttpp">
        <xsl:value-of select="translate(.//图:图片_8005/@图形引用_8007,'Obj','')"/>
      </xsl:variable>
      <a:blipFill>
        <a:blip xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
          <xsl:attribute name="r:embed">
            <xsl:value-of select="concat('rId_imageFill_',$tttpp)"/>
          </xsl:attribute>
        </a:blip>
        <a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>
        <xsl:if test=".//图:图片_8005/@位置_8006='stretch'">
          <!--<a:stretch>
            <a:fillRect/>
          </a:stretch>-->
        </xsl:if>
      </a:blipFill>
    </xsl:if>
    <xsl:if test=".//图:图案_800A">
      <a:pattFill>
        <xsl:attribute name="prst">
          <xsl:call-template name="FillName"/>
        </xsl:attribute>
        <a:fgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:choose >
                <xsl:when test =".//图:图案_800A/@前景色_800B = 'auto'">
                  <xsl:value-of select ="'FFFFFF'"/>
                </xsl:when>
                <xsl:otherwise >
                  <xsl:value-of select ="substring(.//图:图案_800A/@前景色_800B,2,6)"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--<xsl:value-of select="substring(.//图:图案_800A/@前景色_800B,2,6)"/>-->
            </xsl:attribute>
          </a:srgbClr>
        </a:fgClr>
        <!--2014-3-13, update by Qihy, 图案填充打开需要恢复， start-->
        <!--<a:bgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:choose >
                <xsl:when test =".//图:图案_800A/@背景色_800C = 'auto'">
                  <xsl:value-of select ="'FFFFFF'"/>
                </xsl:when>
                <xsl:otherwise >
                  <xsl:value-of select="substring(.//图:图案_800A/@背景色_800C,2,6)"/>
                </xsl:otherwise>
              </xsl:choose>
              --><!--<xsl:value-of select="substring(.//图:图案_800A/@背景色_800C,2,6)"/>--><!--
            </xsl:attribute>
          </a:srgbClr>
        </a:bgClr>-->
        <a:bgClr>
          <xsl:choose>
            <xsl:when test =".//图:图案_800A/@背景色_800C = 'auto'">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select ="'FFFFFF'"/>
                </xsl:attribute>
              </a:srgbClr>
            </xsl:when>
            <xsl:when test ="contains(.//图:图案_800A/@背景色_800C, '#')">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring(.//图:图案_800A/@背景色_800C,2,6)"/>
                </xsl:attribute>
              </a:srgbClr>
            </xsl:when>
            <xsl:otherwise>
              <a:schemeClr>
                <xsl:attribute name="val">
                  <xsl:value-of select=".//图:图案_800A/@背景色_800C"/>
                </xsl:attribute>
              </a:schemeClr>
            </xsl:otherwise>
          </xsl:choose>
        </a:bgClr>
        <!--end-->
      </a:pattFill>
    </xsl:if>
  </xsl:template>
  
  <xsl:template name="FillName">
    <xsl:choose>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn002'">pct10</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn003'">pct20</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn004'">pct25</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn005'">pct30</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn006'">pct40</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn007'">pct50</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn008'">pct60</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn009'">pct70</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn010'">pct75</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn011'">pct80</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn012'">pct90</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn013'">ltDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn014'">ltUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn015'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn016'">dkUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn017'">wdDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn018'">wdUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn019'">ltVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn020'">ltHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn021'">narVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn022'">narHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn023'">dkVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn024'">dkHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn025'">dashDnDiag</xsl:when>
      
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn026'">dashUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn027'">dashHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn028'">dashVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn029'">smConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn030'">lgConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn031'">zigZag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn032'">wave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn033'">diagBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn034'">horzBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn035'">weave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn036'">plaid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn037'">divot</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn038'">dotGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn039'">dotDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn040'">shingle</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn041'">trellis</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn042'">sphere</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn043'">smGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn044'">lgGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn045'">smCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn046'">lgCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn047'">openDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn048'">solidDmnd</xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--drawing文件对应-->
  <!--李杨修改 2012-2-20-->
  <xsl:template name="drawing">
    <xsl:variable name="xCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="yCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="width">
      <xsl:value-of select="./uof:大小_C621/@宽_C605"/>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:value-of select="./uof:大小_C621/@长_C604"/>
    </xsl:variable>
    <xsl:variable name="sheetNo">
      <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
    </xsl:variable>
    <!--随动方式-->
    <xsl:variable name="moveType">
      <xsl:choose>
        <xsl:when test="./@表:随动方式">
          <xsl:value-of select="./@表:随动方式"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'move and re-size'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="chartNo" select="position()"/>
    <!--系列名-->
    <xsl:variable name="seriesName">
      <xsl:choose>
        <xsl:when test="./表:数据源/表:系列/@表:系列名">
          <xsl:value-of select="./表:数据源/表:系列/@表:系列名"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'系列'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!--系列值-->
    <xsl:variable name="seriesValue">
      <xsl:value-of select="./表:数据源/表:系列/@表:系列值"/>
    </xsl:variable>
    <xdr:twoCellAnchor>
      
      <xsl:call-template name="AxisToCellLocation">
        <xsl:with-param name="xCoordinate" select="$xCoordinate"/>
        <xsl:with-param name="yCoordinate" select="$yCoordinate"/>
        <xsl:with-param name="width" select="$width"/>
        <xsl:with-param name="height" select="$height"/>
        <xsl:with-param name="xUnit" select="'54'"/><!--72-->
        <xsl:with-param name="yUnit" select="'14'"/><!--19-->
      </xsl:call-template>
      
      <xdr:graphicFrame macro="">
        <xdr:nvGraphicFramePr>
          <xdr:cNvPr>
            <xsl:attribute name="id">
              <xsl:value-of select="$chartNo"/>
            </xsl:attribute>
            <!--默认为‘图表’UOF中无法设置图表的名称-->
            <xsl:attribute name="name">
              <xsl:value-of select="concat('图表',$chartNo)"/>
            </xsl:attribute>
          </xdr:cNvPr>
          <xdr:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"/>
          </xdr:cNvGraphicFramePr>
        </xdr:nvGraphicFramePr>
        <!--图表默认的此部分扩展和偏移无效？-->
        <xdr:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </xdr:xfrm>
        <a:graphic>
          <a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart">
            <c:chart xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('rId_chart',$chartNo)"/>
              </xsl:attribute>
            </c:chart>
          </a:graphicData>
        </a:graphic>
      </xdr:graphicFrame>
      <xdr:clientData/>
    </xdr:twoCellAnchor>
  </xsl:template>
  <!--图表对应的relationship-->
  <xsl:template name="ChartRelationship">
    <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:variable name="sheetNo">
        <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
      </xsl:variable>
      <xsl:variable name="chartID">
        <!--
            <xsl:value-of select="concat('chart',$sheetNo,'_',@uof:chartNo)"/>
            -->
        <xsl:value-of select="concat('chart',$sheetNo,'_',position())"/>
      </xsl:variable>
      <xsl:variable name="rId">
        <xsl:value-of select="concat('rId_chart',position())"/>
      </xsl:variable>
      <xsl:variable name="target">
        <xsl:value-of select="concat('../charts/',$chartID,'.xml')"/>
      </xsl:variable>
      <xsl:attribute name="Id">
        <xsl:value-of select="$rId"/>
      </xsl:attribute>
      <xsl:attribute name="Type">
        <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chart'"/>
      </xsl:attribute>
      <xsl:attribute name="Target">
        <xsl:value-of select="$target"/>
      </xsl:attribute>
    </Relationship>
  </xsl:template>
  <xsl:template name="ChartSheetRelationship">
    <xsl:param name="seq"/>
    <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:variable name="sheetNo">
        <xsl:value-of select="$seq"/>
      </xsl:variable>
      <xsl:variable name="chartID">
        <xsl:value-of select="concat('chartsheet',$sheetNo)"/>
      </xsl:variable>
      <xsl:variable name="rId">
        <xsl:value-of select="concat('rIdChartSheet',$seq)"/>
      </xsl:variable>
      <xsl:variable name="target">
        <xsl:value-of select="concat('../charts/',$chartID,'.xml')"/>
      </xsl:variable>
      <xsl:attribute name="Id">
        <xsl:value-of select="$rId"/>
      </xsl:attribute>
      <xsl:attribute name="Type">
        <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chart'"/>
      </xsl:attribute>
      <xsl:attribute name="Target">
        <xsl:value-of select="$target"/>
      </xsl:attribute>
    </Relationship>
  </xsl:template>
  <!--锚点对应的relationship-->
  <!--图片引用修改。单元格填充中的图片填充，由于OO不支持，所以不转换。 李杨2012-2-20-->
  <xsl:template name="AnchorRelationship">
    <!--<xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=./@图形引用_C62E]/图:图片数据引用_8037">-->
    <!--添加 李杨2012-2-20-->  
	  
    <xsl:variable name ="graphRef">
          <xsl:value-of select ="./@图形引用_C62E"/>
      </xsl:variable>

	  <!--<xsl:value-of select ="./@图形引用_C62E"/>-->

	  <xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRef]/图:图片数据引用_8037 ">
      <xsl:variable name ="picref">
        <xsl:value-of select ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRef]/图:图片数据引用_8037"/>
      </xsl:variable>
      <xsl:variable name ="picName1">
        <xsl:value-of select ="./ancestor::uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]/对象:路径_D703"/>
      </xsl:variable>
      <xsl:variable name ="picName2">
        <!-- update by xuzhenwei BUG_2674:图片内容丢失 2013-01-10 -->
        <xsl:choose>
          <xsl:when test="contains($picName1,'.\data\')">
            <xsl:value-of select="substring-after($picName1,'.\data\')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="substring-after($picName1,'/data/')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!-- end -->
      </xsl:variable>
      <xsl:variable name ="picName3">
        <xsl:value-of select ="substring-before($picName2,'.')"/>
      </xsl:variable>
    
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:variable name="sheetNo">
          <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="$picName3"/>
        </xsl:variable>
        <xsl:variable name="imageID">
          <xsl:value-of select="concat('image',$sheetNo,'_',$imageName)"/>
        </xsl:variable>
        <xsl:variable name="rId">
          <!--
          <xsl:value-of select="concat('rId_',$imageName)"/>
          -->
          <xsl:value-of select="concat('rID_image',position())"/>
        </xsl:variable>
        <xsl:attribute name="Id">
          <xsl:value-of select="$rId"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>

        <xsl:variable name="target">
          <xsl:value-of select="concat('../media/',$picName2)"/>
        </xsl:variable>
        <xsl:attribute name="Target">
          <xsl:value-of select="$target"/>
        </xsl:attribute>
      </Relationship>
    </xsl:if>
    <xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRef]//图:填充_804C/图:图片_8005">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:variable name="sheetNo">
          <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRef]//图:填充_804C/图:图片_8005/@图形引用_8007"/>
        </xsl:variable>
        <xsl:variable name="imageID">
          <xsl:value-of select="./ancestor::uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$imageName]/对象:路径_D703"/>
        </xsl:variable>
        <xsl:variable name ="imageName2">
        
          <!-- 20130207 gaoyuwei BUG_2672_1 转换后文档需要修复(其中预定义图形图片填充出错) start -->
			<xsl:choose>
				<xsl:when test="contains($imageID,'.\data\')">
					<xsl:value-of select="substring-after($imageID,'.\data\')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="substring-after($imageID,'/data/')"/>
				</xsl:otherwise>
			</xsl:choose>
			<!-- end -->
          
        </xsl:variable>
        <xsl:variable name="rId">
          <xsl:value-of select="concat('rId_imageFill_',translate($imageName,'Obj',''))"/>
        </xsl:variable>
        <xsl:attribute name="Id">
          <xsl:value-of select="$rId"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>

        <!--<xsl:variable name="postfix_current">
          <xsl:value-of select="/uof:UOF/uof:对象集/uof:其他对象[@uof:标识符=$imageID]/@uof:公共类型"/>
        </xsl:variable>
        <xsl:variable name="postfix">
          <xsl:choose>
            <xsl:when test="$postfix_current='jpg' or $postfix_current='jpeg'">
              <xsl:value-of select="'jpeg'"/>
            </xsl:when>
            <xsl:when test="$postfix_current='bmp'">
              <xsl:value-of select="'png'"/>
            </xsl:when>
            <xsl:when test="$postfix_current='png'">
              <xsl:value-of select="'png'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$postfix_current"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>-->
        <xsl:variable name="target">
          <xsl:value-of select="concat('../media/',$imageName2)"/>
        </xsl:variable>
        <xsl:attribute name="Target">
          <xsl:value-of select="$target"/>
        </xsl:attribute>
      </Relationship>
    </xsl:if>
  </xsl:template>
  <!-- 20130319 add by xuzhenwei 组合图形图片丢失问题-->
  <xsl:template name="ZuheAnchorRelationship">
    <!--<xsl:param name ="graphRefzuhe"/>-->

	  <!-- 20130425 update by gaoyuwei 2875_1 uof-oo功能测试 组合图片图片丢失 start-->
	  <xsl:param name ="groupPosition"/>
 	  <xsl:variable name ="graphRefzuhe" select=" ./@标识符_804B"/>
	  <!--end-->
    <xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRefzuhe]/图:图片数据引用_8037">
      <xsl:variable name ="picref">
        <xsl:value-of select ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRefzuhe]/图:图片数据引用_8037"/>
      </xsl:variable>
		<xsl:variable name ="picName1">
        <xsl:value-of select ="./ancestor::uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]/对象:路径_D703"/>
      </xsl:variable>
      <xsl:variable name ="picName2">
        <!-- update by xuzhenwei BUG_2674:图片内容丢失 2013-01-10 -->
        <xsl:choose>
          <xsl:when test="contains($picName1,'.\data\')">
            <xsl:value-of select="substring-after($picName1,'.\data\')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="substring-after($picName1,'/data/')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!-- end -->
      </xsl:variable>
      <xsl:variable name ="picName3">
        <xsl:value-of select ="substring-before($picName2,'.')"/>
      </xsl:variable>

      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:variable name="sheetNo">
          <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="$picName3"/>
        </xsl:variable>
        <xsl:variable name="imageID">
          <xsl:value-of select="concat('image',$sheetNo,'_',$imageName)"/>
        </xsl:variable>
		  <!-- 20130425 update by gaoyuwei 2875_2 uof-oo功能测试 组合图片图片丢失 start-->
        <xsl:variable name="rId">
          <xsl:value-of select="concat('rID_image',$groupPosition)"/>
        </xsl:variable>
		  <!--end-->
        <xsl:attribute name="Id">
          <xsl:value-of select="$rId"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>
        <xsl:variable name="target">
          <xsl:value-of select="concat('../media/',$picName2)"/>
        </xsl:variable>
        <xsl:attribute name="Target">
          <xsl:value-of select="$target"/>
        </xsl:attribute>
      </Relationship>
    </xsl:if>
    <xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRefzuhe]//图:填充_804C/图:图片_8005">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:variable name="sheetNo">
          <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphRefzuhe]//图:填充_804C/图:图片_8005/@图形引用_8007"/>
        </xsl:variable>
        <xsl:variable name="imageID">
          <xsl:value-of select="./ancestor::uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$imageName]/对象:路径_D703"/>
        </xsl:variable>
        <xsl:variable name ="imageName2">
          <!-- 20130207 gaoyuwei BUG_2672_1 转换后文档需要修复(其中预定义图形图片填充出错) start -->
          <xsl:choose>
            <xsl:when test="contains($imageID,'.\data\')">
              <xsl:value-of select="substring-after($imageID,'.\data\')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after($imageID,'/data/')"/>
            </xsl:otherwise>
          </xsl:choose>
          <!-- end -->

        </xsl:variable>
        <xsl:variable name="rId">
          <xsl:value-of select="concat('rId_imageFill_',translate($imageName,'Obj',''))"/>
        </xsl:variable>
        <xsl:attribute name="Id">
          <xsl:value-of select="$rId"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>
        <xsl:variable name="target">
          <xsl:value-of select="concat('../media/',$imageName2)"/>
        </xsl:variable>
        <xsl:attribute name="Target">
          <xsl:value-of select="$target"/>
        </xsl:attribute>
      </Relationship>
    </xsl:if>
  </xsl:template>
  <!-- end -->
  <!--预定义图形图片填充对应的realtionship-->
  <xsl:template name="PicInPreGraphic">
    <xsl:if test="./图:图形[@图:其他对象]">
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:variable name="sheetNo">
          <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="./图:图形/@图:其他对象"/>
        </xsl:variable>
        <xsl:variable name="imageID">
          <xsl:value-of select="concat('image',$sheetNo,'_',$imageName)"/>
        </xsl:variable>
        <xsl:variable name="rId">
          <!--
          <xsl:value-of select="concat('rId_',$imageName)"/>
          -->
          <xsl:value-of select="concat('rID_image',position())"/>
        </xsl:variable>
        <xsl:attribute name="Id">
          <xsl:value-of select="$rId"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>

        <xsl:variable name="postfix_current">
          <xsl:value-of select="/uof:UOF/uof:对象集/uof:其他对象[@uof:标识符=$imageName]/@uof:公共类型"/>
        </xsl:variable>
        <xsl:variable name="postfix">
          <xsl:choose>
            <xsl:when test="$postfix_current='jpg' or $postfix_current='jpeg'">
              <xsl:value-of select="'jpeg'"/>
            </xsl:when>
            <xsl:when test="$postfix_current='bmp'">
              <xsl:value-of select="'png'"/>
            </xsl:when>
            <xsl:when test="$postfix_current='png'">
              <xsl:value-of select="'png'"/>
            </xsl:when>
            <!--
            <xsl:when test="$emf!=''">
              <xsl:value-of select="$emf"/>
            </xsl:when>
            -->
            <xsl:otherwise>
              <xsl:value-of select="$postfix_current"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="target">
          <xsl:value-of select="concat('../media/',$imageName,'.',$postfix)"/>
        </xsl:variable>
        <xsl:attribute name="Target">
          <xsl:value-of select="$target"/>
        </xsl:attribute>
      </Relationship>
    </xsl:if>
  </xsl:template>
  <xsl:template name="drawpicture">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select="uof:锚点[图:图形/@图:其他对象]">
        <xsl:for-each select="图:图形[@图:其他对象]">
          <xsl:variable name="uxux" select="@图:标识符"/>
          <xsl:variable name="tttt">
            <xsl:value-of select="@图:其他对象"/>
          </xsl:variable>
          <xsl:variable name="tttt1">
            <xsl:choose>
              <xsl:when test="contains($uxux,'OBJ0000')">
                <xsl:value-of select="substring-after($uxux,'OBJ0000')"/>
              </xsl:when>
              <xsl:when test="contains($uxux,'OBJ000')">
                <xsl:value-of select="substring-after($uxux,'OBJ000')"/>
              </xsl:when>
              <xsl:when test="contains($uxux,'OBJ00')">
                <xsl:value-of select="substring-after($uxux,'OBJ00')"/>
              </xsl:when>
              <xsl:when test="contains($uxux,'OBJ0')">
                <xsl:value-of select="substring-after($uxux,'OBJ0')"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="ttttt">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:variable name="uxux1">
            <xsl:choose>
              <xsl:when test="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:公共类型">
                <xsl:value-of select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:公共类型"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="emf" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:私有类型"/>
          <xsl:comment>
            <xsl:value-of select="$emf"/>
          </xsl:comment>
          <Relationship>
            <xsl:attribute name="Id">
              <xsl:value-of select="concat('rId',$tttt1)"/>
            </xsl:attribute>
            <xsl:attribute name="Type">
              <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
            </xsl:attribute>
            <xsl:variable name="uxuxt2">
              <xsl:choose>
                <xsl:when test="$uxux1='jpg' or $uxux1='jpeg'">
                  <xsl:value-of select="'jpeg'"/>
                </xsl:when>
                <xsl:when test="$uxux1='bmp'">
                  <xsl:value-of select="'png'"/>
                </xsl:when>
                <xsl:when test="$uxux1='png'">
                  <xsl:value-of select="'png'"/>
                </xsl:when>
                <xsl:when test="$emf!=''">
                  <xsl:value-of select="$emf"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$uxux1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:attribute name="Target">
              <xsl:value-of select="concat('../media/',$tttt,'.',$uxuxt2)"/>
            </xsl:attribute>
          </Relationship>
        </xsl:for-each>
      </xsl:for-each>
      <xsl:for-each select="uof:锚点/图:图形/图:预定义图形[..//图:图片]">
        <xsl:variable  name="pixxt"    select="parent::图:图形/@图:标识符"/>
        <xsl:variable name="ttttp">
          <xsl:choose>
            <xsl:when test="contains($pixxt,'OBJ0000')">
              <xsl:value-of select="substring-after($pixxt,'OBJ0000')"/>
            </xsl:when>
            <xsl:when test="contains($pixxt,'OBJ000')">
              <xsl:value-of select="substring-after($pixxt,'OBJ000')"/>
            </xsl:when>
            <xsl:when test="contains($pixxt,'OBJ00')">
              <xsl:value-of select="substring-after($pixxt,'OBJ00')"/>
            </xsl:when>
            <xsl:when test="contains($pixxt,'OBJ0')">
              <xsl:value-of select="substring-after($pixxt,'OBJ0')"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="tuyinyong">
          <xsl:value-of select="图:属性/图:填充/图:图片/@图:图形引用"/>
        </xsl:variable>
        <xsl:variable name="uxux1t" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tuyinyong]/@uof:公共类型"/>
        <xsl:variable name="uxux2t" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/图:图形[//图:图片]/@图:标识符"/>
        <xsl:variable name="private" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tuyinyong]/@uof:私有类型"/>
        <xsl:variable name="biaozhitt">
          <xsl:choose>
            <xsl:when test="$uxux1t='jpg'">
              <xsl:value-of select="'jpeg'"/>
            </xsl:when>
            <xsl:when test="$uxux1t='bmp'">
              <xsl:value-of select="png"/>
            </xsl:when>
            <xsl:when test="$private!=''">
              <xsl:value-of select="$private"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$uxux1t"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select="concat('rId',$ttttp)"/>
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat('../media/',$tuyinyong,'.',$biaozhitt)"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
    </Relationships>
  </xsl:template>

  <!--图表中填充图片的chart.xml.rel模板-->
  <!--Modified by LDM in 2010/12/18-->
  
 <!--20130116 gaoyuwei 2642 标题集图片填充的填充内容丢失 start-->
  <xsl:template name="PhotosInChartsRels">
    <xsl:variable name="picRef">
      <!--<xsl:value-of select="./@图:图形引用_8007"/>-->
		<xsl:value-of select="./@图形引用_8007"/>
	</xsl:variable>
	  <xsl:variable name ="picName1">
		  <xsl:value-of select ="/uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picRef]/对象:路径_D703"/>
	  </xsl:variable>
	  <xsl:variable name ="picName2">
		  <xsl:choose>
			  <xsl:when test="contains($picName1,'.\data\')">
				  <xsl:value-of select="substring-after($picName1,'.\data\')"/>
			  </xsl:when>
			  <xsl:otherwise>
				  <xsl:value-of select="substring-after($picName1,'/data/')"/>
			  </xsl:otherwise>
		  </xsl:choose>
		  <!-- end -->
	  </xsl:variable>
	  <Relationship>
		  <xsl:attribute name="Id">
			  <xsl:value-of select="concat('rId',$picRef)"/>
		  </xsl:attribute>
		  <xsl:attribute name="Type">
			  <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
		  </xsl:attribute>
		  <xsl:attribute name="Target">
			  <xsl:value-of select ="concat('../media/',$picName2)"/>
		  </xsl:attribute>
      </Relationship>
  </xsl:template>
	
	<!--end-->


  <xsl:template name ="Axis">
    <xsl:param name="xCoordinate"/>
    <xsl:param name="yCoordinate"/>
    <xsl:param name="width"/>
    <xsl:param name="height"/>
    <xsl:param name="xUnit"/>
    <xsl:param name="yUnit"/>

    <!--如何更好的确定起终单元格的位置-->
    <!--Not Finished-->
    <xsl:variable name="fromCol" select="floor($xCoordinate div $xUnit)"/>
    <xsl:variable name="fromColOff">
      <xsl:value-of select="$xCoordinate"/>
    </xsl:variable>
    <xsl:variable name="fromRow" select="floor($yCoordinate div $yUnit)"/>
    <xsl:variable name="fromRowOff">
      <xsl:value-of select="$yCoordinate * 14.25 div 13.5"/>
    </xsl:variable>
    <xsl:variable name="toCol" select="floor(($xCoordinate + $width) div $xUnit)"/>
    <xsl:variable name="toColOff">
      <xsl:value-of select="($xCoordinate + $width) *1.062"/>
    </xsl:variable>
    <xsl:variable name="toRow" select="floor(($yCoordinate + $height) div $yUnit)"/>
    <xsl:variable name="toRowOff">
      <xsl:value-of select="($yCoordinate + $height)"/>
    </xsl:variable>

    <xdr:from>
      <xdr:col>
        <xsl:choose >
          <xsl:when test ="contains($fromCol,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromCol"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:col>
      <xdr:colOff>
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        <xsl:value-of select="$xCoordinate"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        <xsl:choose >
          <xsl:when test ="contains($fromRow,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromRow"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$fromRowOff"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>
        <xsl:value-of select="$toCol "/>
      </xdr:col>
      <xdr:colOff>
        <!-- 20130506 update by xuzhenwei BUG_2882 回归 互操作oo-uof（编辑）-oo 文件010图表.xlsx转换后的.xlsx文件需要修复 start
        <xsl:choose>
          <xsl:when test="contains($toColOff,'-')">
            <xsl:value-of select ="substring-after($toColOff,'-')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$toColOff"/>
          </xsl:otherwise>
        </xsl:choose>
        end -->
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromRowOff"/>-->
        <xsl:value-of select="$xCoordinate + $width"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        <xsl:choose >
          <xsl:when test ="($toRow - 3) &lt;= $fromRow">
            <xsl:value-of select ="$fromRow"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$toRow"/>
          </xsl:otherwise>
        </xsl:choose>
        
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$toRowOff"/>
      </xdr:rowOff>
    </xdr:to>
  </xsl:template>
  <!--Modified by LDM in 2011/01/06-->
  <!--Picture-->
  <!--修改 李杨2012-2-20-->
  <xsl:template name="PictureDrawing">
    <xsl:param name ="id"/>
    <xsl:variable name="xCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="yCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="width">
      <xsl:value-of select="./uof:大小_C621/@宽_C605"/>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:value-of select="./uof:大小_C621/@长_C604"/>
    </xsl:variable>

    <xdr:twoCellAnchor editAs="oneCell">

      <!--修改模板Axis 李杨 2012-5-28-->
      
      <!--2014-4-17, update by Qihy, 修复bug3147，图变形， start-->
      <!--<xsl:call-template name="Axis">-->
      <xsl:call-template name="AxisToCellLocation">
      <!--2014-4-17 end-->
        
        <xsl:with-param name="xCoordinate" select="$xCoordinate"/>
        <xsl:with-param name="yCoordinate" select="$yCoordinate"/>
        <xsl:with-param name="width" select="$width"/>
        <xsl:with-param name="height" select="$height"/>
        <xsl:with-param name="xUnit" select="'54'"/>
        <xsl:with-param name="yUnit" select="'14.25'"/>
      </xsl:call-template>

      <xdr:pic>
        <xdr:nvPicPr>
          <xdr:cNvPr>
            <xsl:variable name="ID">
              <!--默认最大只能有100个图表，100以后为图片的ID-->
              <!--此处仅是权宜之策-->
              <xsl:value-of select="number(100 + position())"/>
            </xsl:variable>
            <xsl:attribute name="id">
              <xsl:value-of select="$ID"/>
            </xsl:attribute>

            <xsl:attribute name="name">
              <xsl:value-of select ="./@图形引用_C62E"/>
            </xsl:attribute>
          </xdr:cNvPr>
          <xdr:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </xdr:cNvPicPr>
        </xdr:nvPicPr>
        <xdr:blipFill>
          <a:blip xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <xsl:variable name="rID">
              <xsl:value-of select="concat('rID_image',position())"/>
            </xsl:variable>
            <xsl:attribute name="r:embed">
              <xsl:value-of select="$rID"/>
            </xsl:attribute>
            <a:extLst>
              <a:ext uri="">
                <a14:useLocalDpi
                    xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                    val="0"/>
              </a:ext>
            </a:extLst>
          </a:blip>
          <!-- 20130309 add by xuzhenwei 图片裁剪 start -->
          <xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022">
            <a:srcRect>
              <xsl:attribute name="t">
                <xsl:value-of select="concat(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023,'%') "/>
              </xsl:attribute>
              <xsl:attribute name="b">
                <xsl:value-of select="concat(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024,'%')"/>
              </xsl:attribute>
              <xsl:attribute name="l">
                <xsl:value-of select="concat(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025,'%')"/>
              </xsl:attribute>
              <xsl:attribute name="r">
                <xsl:value-of select="concat(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026,'%')"/>
              </xsl:attribute>
            </a:srcRect>
          </xsl:if>
          <!-- end -->
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </xdr:blipFill>
        <xdr:spPr>
          <!--
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="0" cy="0"/>
          </a:xfrm>
          -->
          <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060">
			 <a:xfrm>
         <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:翻转_803A = 'x'">
           <xsl:attribute name="flipH">
             <xsl:value-of select="1"/>
           </xsl:attribute>
         </xsl:if>
         
         <!--2014-6-10, add by Qihy, 增加图片旋转时的转换， start-->
         <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D!='0.0'">
           <xsl:variable name="rotAngle" select="number(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D)"/>
           <xsl:attribute name="rot">
             <xsl:choose >
               <xsl:when test ="./图:预定义图形_8018/图:名称_801A='Extract'">
                 <xsl:value-of select="round(21600000 - (180 - $rotAngle)*60000)"/>
               </xsl:when>
               <xsl:when test ="./图:预定义图形_8018/图:名称_801A='Process'">
                 <xsl:value-of select="round(($rotAngle)*60000*(-1))"/>
               </xsl:when>
               <xsl:otherwise >
                 <xsl:value-of select="round(21600000 - (360 - $rotAngle)*60000)"/>
               </xsl:otherwise>
             </xsl:choose>
           </xsl:attribute>
         </xsl:if>
         <!--2014-6-10 end-->
         
              <a:ext>
                <xsl:attribute name ="cx">
                  <xsl:value-of select ="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*12700)"/>
                </xsl:attribute>
                <xsl:attribute name ="cy">
                  <xsl:value-of select ="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604*12700)"/>
                </xsl:attribute>
              </a:ext>
            </a:xfrm>
          </xsl:if>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    </xdr:twoCellAnchor>
  </xsl:template>

  <xsl:template name ="Axis2">
    <xsl:param name="xCoordinate"/>
    <xsl:param name="yCoordinate"/>
    <xsl:param name="width"/>
    <xsl:param name="height"/>
    <xsl:param name="xUnit"/>
    <xsl:param name="yUnit"/>

    <!--如何更好的确定起终单元格的位置-->
    <!--Not Finished-->
    <xsl:variable name="fromCol" select="floor($xCoordinate div $xUnit)"/>
    <xsl:variable name="fromColOff">
      <xsl:value-of select="$xCoordinate"/>
    </xsl:variable>
    <xsl:variable name="fromRow" select="floor($yCoordinate div $yUnit)"/>
    <xsl:variable name="fromRowOff">
      <xsl:value-of select="$yCoordinate * 14.25 div 13.5"/>
    </xsl:variable>
    <xsl:variable name="toCol" select="floor(($xCoordinate + $width) div $xUnit -1.5)"/>
    <xsl:variable name="toColOff">
      <xsl:value-of select="($xCoordinate + $width) *1.062"/>
    </xsl:variable>
    <xsl:variable name="toRow" select="floor(($yCoordinate + $height) div $yUnit)"/>
    <xsl:variable name="toRowOff">
      <xsl:value-of select="($yCoordinate + $height)"/>
    </xsl:variable>

    <xdr:from>
      <xdr:col>
        <xsl:choose >
          <xsl:when test ="contains($fromCol,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromCol"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:col>
      <xdr:colOff>
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        <xsl:value-of select="$xCoordinate"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
        <xsl:choose >
          <xsl:when test ="contains($fromRow,'-')">
            <xsl:value-of select ="'0'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select="$fromRow"/>
          </xsl:otherwise>
        </xsl:choose>

      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$fromRowOff"/>
      </xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>
        <xsl:value-of select="$toCol"/>
      </xdr:col>
      <xdr:colOff>
        <!-- 20130506 update by xuzhenwei BUG_2882 回归 互操作oo-uof（编辑）-oo 文件010图表.xlsx转换后的.xlsx文件需要修复 start
        <xsl:choose>
          <xsl:when test="contains($toColOff,'-')">
            <xsl:value-of select ="substring-after($toColOff,'-')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$toColOff"/>
          </xsl:otherwise>
        </xsl:choose>
        end -->
        <!-- update by xuzhenwei 修正从UOF到OOXML方向，单元格拉伸问题 start -->
        <!--<xsl:value-of select="$fromColOff"/>-->
        <xsl:value-of select="$xCoordinate + $width"/>
        <!-- end -->
      </xdr:colOff>
      <xdr:row>
            <xsl:value-of select="$toRow"/>
      </xdr:row>
      <xdr:rowOff>
        <xsl:value-of select="$toRowOff"/>
      </xdr:rowOff>
    </xdr:to>
  </xsl:template>
  <!--Modified by LDM in 2011/01/06-->
  <!--Predefined Picture-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="PredefinedPicDrawing">
    <xsl:param name="seq"/>
    
    <xsl:variable name="xCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="yCoordinate">
      <xsl:value-of select="./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="width">
      <xsl:value-of select="./uof:大小_C621/@宽_C605"/>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:value-of select="./uof:大小_C621/@长_C604"/>
    </xsl:variable>

    <xdr:twoCellAnchor editAs="oneCell">
      <!--Axis2模板（预定义图形大小正常）和AxisToCellLocation（批注无修复）模板有冲突 李杨2012-5-28-->
      <!--20130304 gaoyuwei Bug_2672_2 转换后文档需要修复 start-->
		<xsl:variable name ="grapref2">
			<xsl:value-of select ="./@图形引用_C62E"/>
		</xsl:variable>
		<xsl:if test="not(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2 and @组合列表_8064])">
			<xsl:call-template name="AxisToCellLocation">
				<xsl:with-param name="xCoordinate" select="$xCoordinate"/>
				<xsl:with-param name="yCoordinate" select="$yCoordinate"/>
				<xsl:with-param name="width" select="$width"/>
				<xsl:with-param name="height" select="$height"/>
				<xsl:with-param name="xUnit" select="'54'"/>
				<xsl:with-param name="yUnit" select="'14.25'"/>
			</xsl:call-template>
		</xsl:if>

		<!--end-->
		
      <!--修改组合图形相关调用。李杨2012-2-28-->
      <xsl:choose>
        <xsl:when test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2 and @组合列表_8064]">
        
        <!--20130304 gaoyuwei Bug_2672_3 转换后文档需要修复 start-->
			<xsl:call-template name="AxisToCellLocation">
				<xsl:with-param name="xCoordinate" select="$xCoordinate"/>
				<xsl:with-param name="yCoordinate" select="$yCoordinate"/>
				<xsl:with-param name="width" select="$width"/>
				<xsl:with-param name="height" select="$height"/>
				<xsl:with-param name="xUnit" select="'54'"/>
				<xsl:with-param name="yUnit" select="'14.25'"/>
			</xsl:call-template>
			<!--end-->
			
          <xdr:grpSp>
            <xdr:nvGrpSpPr>
              <xdr:cNvPr>
                <xsl:variable name="ID">
                  <!--默认最多只能有100个图表，最多只能添加100张外图片，以后为预定义图片的ID-->
                  <!--此处仅是权宜之策-->
                  <!--
                <xsl:value-of select="number(200 + position())"/>
                -->
                  <xsl:value-of select="$seq"/>
                </xsl:variable>
                <xsl:attribute name="id">
                  <xsl:value-of select="$ID"/>
                </xsl:attribute>

                <xsl:attribute name="name">
                  <xsl:value-of select="concat('组合',./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062/@标识符_804B)"/>
                </xsl:attribute>
              </xdr:cNvPr>
              <xdr:cNvGrpSpPr/>
            </xdr:nvGrpSpPr>
            <xdr:grpSpPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="0" cy="0"/>
              </a:xfrm>
            </xdr:grpSpPr>

            <xsl:variable name ="groupList">
              <xsl:value-of select ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2]/@组合列表_8064"/>
            </xsl:variable>
            <!-- 20130309 add by xuzhenwei BUG_2705:图片内容丢失 start -->
            <xsl:choose>
                          <xsl:when test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2]/@组合列表_8064">
                            
                            <xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062">
                              <xsl:variable name="id" select="@标识符_804B"/>
                              <xsl:if test ="contains($groupList,$id)">
                                <xsl:choose>
                                  <xsl:when test ="./图:图片数据引用_8037">
                                    <xdr:pic>
                                      <xdr:nvPicPr>
                                        <xdr:cNvPr>
                                          <xsl:variable name="ID">
                                            <!--默认最大只能有100个图表，100以后为图片的ID-->
                                            <!--此处仅是权宜之策-->
                                            <xsl:value-of select="number(100 + position())"/>
                                          </xsl:variable>
                                          <xsl:attribute name="id">
                                            <xsl:value-of select="$ID"/>
                                          </xsl:attribute>

                                          <xsl:attribute name="name">
                                            <xsl:value-of select ="./@图形引用_C62E"/>
                                          </xsl:attribute>
                                        </xdr:cNvPr>
                                        <xdr:cNvPicPr>
                                          <a:picLocks noChangeAspect="1"/>
                                        </xdr:cNvPicPr>
                                      </xdr:nvPicPr>
                                      <xdr:blipFill>
                                        <a:blip xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
											<!-- 20130426 update by gaoyuwei 2875_3 uof-oo功能测试 组合图片图片丢失 start--> 
										<xsl:variable name="rID">
                                            <xsl:value-of select="concat('rID_image',100 + position())"/>
                                          </xsl:variable>
											<!--end-->
                                          <xsl:attribute name="r:embed">
                                            <xsl:value-of select="$rID"/>
                                          </xsl:attribute>
                                          <a:extLst>
                                            <a:ext uri="">
                                              <a14:useLocalDpi
                                                  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                                  val="0"/>
                                            </a:ext>
                                          </a:extLst>
                                        </a:blip>
                                        <xsl:if test="./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022">
                                          <a:srcRect>
                                            <xsl:attribute name="t">
                                              <xsl:value-of select="concat(./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023,'%')"/>
                                            </xsl:attribute>
                                            <xsl:attribute name="b">
                                              <xsl:value-of select="concat(./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024,'%')"/>
                                            </xsl:attribute>
                                            <xsl:attribute name="l">
                                              <xsl:value-of select="concat(./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025,'%')"/>
                                            </xsl:attribute>
                                            <xsl:attribute name="r">
                                              <xsl:value-of select="concat(./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026,'%')"/>
                                            </xsl:attribute>
                                          </a:srcRect>
                                        </xsl:if>
                                        <a:stretch>
                                          <a:fillRect/>
                                        </a:stretch>
                                      </xdr:blipFill>
                                      <xdr:spPr>
                                        
                                          <a:xfrm>
                                       <!-- 20130423 add by gaoyuwei 2875_4 uof-oo功能测试 组合图片图片丢失 start-->
											  <xsl:choose>
												  <!--图形图片组合 和 图片图片组合是不同的 前者的位置都取组合位置 后者位置分别取对应每个图形的“组合位置”-->
												  <xsl:when test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B = $id]/图:图片数据引用_8037">
													  <a:off>
														  <xsl:attribute name="x">

															  <xsl:value-of select="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:组合位置_803B/@x_C606 * 12700)"/>
														  </xsl:attribute>
														  <xsl:attribute name="y">
															  <xsl:value-of select="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:组合位置_803B/@y_C607 * 12700)"/>
														  </xsl:attribute>

													  </a:off>
												  </xsl:when>
												  <xsl:when test ="not(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B = $id]/图:图片数据引用_8037)">
													  <xsl:for-each select ="./ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825//uof:锚点_C644[@图形引用_C62E = $grapref2]">
														  <a:off>
														  <xsl:attribute name="x">
														  <xsl:value-of select="round(./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 * 12700)"/>
													      </xsl:attribute>
													      <xsl:attribute name="y">
														  <xsl:value-of select="round(./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108 * 12700)"/>
													      </xsl:attribute>
												          </a:off>
											          </xsl:for-each>
												  </xsl:when>
											  </xsl:choose>	
											  
											  <!--end-->
											  
											 <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060">
                                            <a:ext>
                                              <xsl:attribute name ="cx">
                                                <xsl:value-of select ="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*12700)"/>
                                              </xsl:attribute>
                                              <xsl:attribute name ="cy">
                                                <xsl:value-of select ="round(./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$id]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604*12700)"/>
                                              </xsl:attribute>
                                            </a:ext>
                                            </xsl:if>
                                          </a:xfrm>
                                        
                                        <a:prstGeom prst="rect">
                                          <a:avLst/>
                                        </a:prstGeom>
                                      </xdr:spPr>
                                    </xdr:pic>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    <xsl:call-template name="SpTransfer">
                                      <xsl:with-param name="seq" select="number($seq) - 1"/>
                                    </xsl:call-template>
                                    
                                    </xsl:otherwise>
                                </xsl:choose>

                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:otherwise>
                          </xsl:otherwise>
                        </xsl:choose>
            <!-- end -->
            <!-- <xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062">
              <xsl:if test ="contains($groupList,@标识符_804B)">
                <xsl:call-template name="SpTransfer">
                  <xsl:with-param name="seq" select="number($seq) - 1"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:for-each>-->
          </xdr:grpSp>
        </xsl:when>
        <xsl:otherwise>
        
        <!--20130304 gaoyuwei Bug_2672_4 转换后文档需要修复 start-->
			<xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2]">
				<!--end-->
				
            <xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref2 and not(@组合列表_8064)]">
              <xsl:call-template name="SpTransfer">
                <xsl:with-param name="seq" select="$seq"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
      <!--
      <xsl:if test="./图:图形[not(@图:组合列表)]">
        
      <xsl:if test="">
        
      </xsl:if>
      </xsl:if>
      -->
      <xdr:clientData/>
    </xdr:twoCellAnchor>
  </xsl:template>
  
  <!--李杨修改 2012-2-17-->
  <xsl:template name="SpTransfer">
    <xsl:param name="seq"/>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr>
          <xsl:variable name="ID">
            <!--默认最多只能有100个图表，最多只能添加100张外图片，以后为预定义图片的ID number(200 + position())-->
            <!--此处仅是权宜之策-->
            <xsl:value-of select="$seq"/>
          </xsl:variable>
          <xsl:attribute name="id">
            <xsl:value-of select="$ID"/>
          </xsl:attribute>
          <xsl:attribute name="name">
            <xsl:if test ="not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
              <xsl:value-of select="./@标识符_804B"/>
            </xsl:if>
            <xsl:if test ="./图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
              <xsl:value-of select ="'矩形 2'"/>
            </xsl:if>
          </xsl:attribute>
          <!-- 20130515 add by xuzhenwei 解决 Alt Text功能丢失 start -->
          <xsl:if test ="./图:预定义图形_8018/图:属性_801D/图:Web文字_804F">
            <xsl:attribute name="descr">
              <xsl:value-of select ="./图:预定义图形_8018/图:属性_801D/图:Web文字_804F"/>
            </xsl:attribute>
            <xsl:attribute name="title">
              <xsl:value-of select ="./图:预定义图形_8018/图:属性_801D/图:Web文字_804F"/>
            </xsl:attribute>
          </xsl:if>
          <!-- end -->
        </xdr:cNvPr>
        
        <!--2014-3-13, update by Qihy, 修复图案填充需要修复打开， start-->
        <!--<xdr:cNvSpPr />-->
        <xdr:cNvSpPr txBox="1"/>
        <!--end-->
        
      </xdr:nvSpPr>
      <xdr:spPr>
        <xsl:call-template name="PicLocation"/>
        <!--关键点坐标 待修改。李杨 2012-2-28-->
        <xsl:if test=".//图:关键点坐标">
          <xsl:call-template name="custGeom"/>
        </xsl:if>

        <xsl:if test="not(.//图:关键点坐标) and not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
          <a:prstGeom>
            <xsl:attribute name="prst">
              <xsl:call-template name="PrstName"/>
            </xsl:attribute>
            <a:avLst/>
          </a:prstGeom>
        </xsl:if>
        <!--添加艺术字相关的a:prstGeom prst="rect"值。李杨2012-2-28-->
        <xsl:if test ="not(.//图:关键点坐标) and (./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
          <a:prstGeom>
            <xsl:attribute name ="prst">
              <xsl:value-of select ="'rect'"/>
            </xsl:attribute>
            <a:avLst/>
          </a:prstGeom>
        </xsl:if>
        <!--修改艺术字所在预定义图形的填充，线条为“无线条”。李杨 2012-2-27-->
        <xsl:if test=".//图:填充_804C and not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
          <xsl:call-template name="Fill_Pre"/>
        </xsl:if>
        <xsl:if test =".//图:填充_804C/图:渐变_800D[@起始色_800E and @终止色_800F] and ./图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
          <xsl:call-template name="Fill_Pre"/>
        </xsl:if>
        <xsl:if test =".//图:填充_804C/图:渐变_800D[not(@起始色_800E)] and ./图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
          <a:noFill/>
        </xsl:if>

        <!--修改艺术字所在预定义图形的填充，线条为“无线条”。李杨 2012-2-27-->
        <xsl:if test="(.//图:线_8057|.//图:箭头_805D) and not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
          <xsl:call-template name="Line_Drawing"/>
        </xsl:if>
        <xsl:if test ="(.//图:线_8057|.//图:箭头_805D) and (./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
          <a:ln>
            <a:noFill/>
          </a:ln>
        </xsl:if>

        <!--添加预定义图形的阴影效果。李杨2012-3-1-->
        <xsl:for-each select ="./图:预定义图形_8018/图:属性_801D/图:阴影_8051[@类型_C61D='single']">
            <xsl:call-template name ="阴影single"/>
        </xsl:for-each>
        <xsl:for-each select ="./图:预定义图形_8018/图:属性_801D/图:阴影_8051[@类型_C61D='shaperelative']">
          <xsl:call-template name ="阴影ShapePer"/>
        </xsl:for-each>
        <xsl:for-each select ="./图:预定义图形_8018/图:属性_801D/图:阴影_8051[@类型_C61D='perspective']">
          <xsl:call-template name ="阴影ShapePer"/>
        </xsl:for-each>
        
        <!--添加三维 李杨2012-2-24-->
        <!--20130318 gaoyuwei Bug_2747_1 uof-oo 功能测试 预定义图形4，5样式不正确-->
		  <xsl:if test="not(./图:预定义图形_8018/图:名称_801A='Or') and not(./图:预定义图形_8018/图:名称_801A='Collate') ">
		  <xsl:for-each select ="./图:预定义图形_8018/图:属性_801D/图:三维效果_8061 ">
          <xsl:call-template name ="图:三维效果_8061"/>
        </xsl:for-each>
		  </xsl:if>
		  <!--end-->
      </xdr:spPr>
      
      <!--添加艺术字 李杨2012-2-27-->
      <xsl:if test ="./图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
        <xsl:apply-templates select="./图:预定义图形_8018/图:属性_801D/图:艺术字_802D"/>
      </xsl:if>
      <xsl:for-each select="./图:文本_803C[*]">
        <xsl:call-template name="TxBody"/>
      </xsl:for-each>
    </xdr:sp>
  </xsl:template>
  <xsl:template name="custGeom">
    <a:custGeom>
      <a:avLst/>
      <a:rect l="l" t="t" r="r" b="b"/>
      <xsl:if test=".//图:关键点坐标">
        <a:pathLst>
          <a:path>
            <xsl:if test=".//图:宽度">
              <xsl:attribute name="w">
                <xsl:value-of select="round(.//图:宽度 * 12700)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//图:高度">
              <xsl:attribute name="h">
                <xsl:value-of select="round(.//图:高度 * 12700)"/>
              </xsl:attribute>
            </xsl:if>
            <!--<xsl:for-each select=".//图:关键点坐标/*">-->
            <xsl:for-each select ="./图:控制点_8039/*">
              <xsl:if test="name(.)='M'">
                <a:moveTo>
                  <a:pt>
                    <xsl:attribute name="x">
                      <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                    </xsl:attribute>
                    <xsl:attribute name="y">
                      <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                    </xsl:attribute>
                  </a:pt>
                </a:moveTo>
              </xsl:if>
              <xsl:if test="name(.)='C'">
                <a:cubicBezTo>
                  <a:pt>
                    <xsl:attribute name="x">
                      <xsl:value-of select="round(pt[1]/@x * 12700 div 1.34)"/>
                    </xsl:attribute>
                    <xsl:attribute name="y">
                      <xsl:value-of select="round(pt[1]/@y * 12700 div 1.34)"/>
                    </xsl:attribute>
                  </a:pt>
                  <a:pt>
                    <xsl:attribute name="x">
                      <xsl:value-of select="round(pt[2]/@x * 12700 div 1.34)"/>
                    </xsl:attribute>
                    <xsl:attribute name="y">
                      <xsl:value-of select="round(pt[2]/@y * 12700 div 1.34)"/>
                    </xsl:attribute>
                  </a:pt>
                  <a:pt>
                    <xsl:attribute name="x">
                      <xsl:value-of select="round(pt[3]/@x * 12700 div 1.34)"/>
                    </xsl:attribute>
                    <xsl:attribute name="y">
                      <xsl:value-of select="round(pt[3]/@y * 12700 div 1.34)"/>
                    </xsl:attribute>
                  </a:pt>
                </a:cubicBezTo>
              </xsl:if>
              <xsl:if test="name(.)='L'">
                <a:lnTo>
                  <a:pt>
                    <xsl:attribute name="x">
                      <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                    </xsl:attribute>
                    <xsl:attribute name="y">
                      <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                    </xsl:attribute>
                  </a:pt>
                </a:lnTo>
              </xsl:if>
            </xsl:for-each>
          </a:path>
        </a:pathLst>
      </xsl:if>
    </a:custGeom>
  </xsl:template>
  <!--预定义图形对应列表-->
  <!--PrstName-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="PrstName">
    <xsl:choose>
      <xsl:when test=".//图:名称_801A='Rectangle' or .//图:类别_8019='11'">rect</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='12'">parallelogram</xsl:when>
      <xsl:when test=".//图:名称_801A='Trapezoid' or .//图:类别_8019='13'">trapezoid</xsl:when>
      <xsl:when test=".//图:名称_801A='diamond' or .//图:类别_8019='14'">diamond</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangle' or .//图:类别_8019='15'">roundRect</xsl:when>
      <xsl:when test=".//图:名称_801A='Octagon'or .//图:类别_8019='16'">octagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Isosceles Triangle' or .//图:类别_8019='17'">triangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Triangle' or .//图:类别_8019='18'">rtTriangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval' or .//图:类别_8019='19'">ellipse</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Arrow' or .//图:类别_8019='21'">rightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow' or .//图:类别_8019='22'">leftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow' or .//图:类别_8019='23'">upArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow' or .//图:类别_8019='24'">downArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow' or .//图:类别_8019='25'">leftRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up-Down Arrow' or .//图:类别_8019='26'">upDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Quad Arrow' or .//图:类别_8019='27'">quadArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right-Up Arrow' or .//图:类别_8019='28'">leftRightUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent Arrow' or .//图:类别_8019='29'">bentArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Process' or .//图:类别_8019='31'">flowChartProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Alternate Process' or .//图:类别_8019='32'">flowChartAlternateProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Decision' or .//图:类别_8019='33'">flowChartDecision</xsl:when>
      <xsl:when test=".//图:名称_801A='Data' or .//图:类别_8019='34'">flowChartInputOutput</xsl:when>
      <xsl:when test=".//图:名称_801A='Predefined Process' or .//图:类别_8019='35'">flowChartPredefinedProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Internal Storage' or .//图:类别_8019='36'">flowChartInternalStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Document' or .//图:类别_8019='37'">flowChartDocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Multidocument' or .//图:类别_8019='38'">flowChartMultidocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Terminator' or .//图:类别_8019='39'">flowChartTerminator</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 1' or .//图:类别_8019='41'">irregularSeal1</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 2' or .//图:类别_8019='42'">irregularSeal2</xsl:when>
      <xsl:when test=".//图:名称_801A='4-Point Star' or .//图:类别_8019='43'">star4</xsl:when>
      <xsl:when test=".//图:名称_801A='5-Point Star' or .//图:类别_8019='44'">star5</xsl:when>
      <xsl:when test=".//图:名称_801A='8-Point Star' or .//图:类别_8019='45'">star8</xsl:when>
      <xsl:when test=".//图:名称_801A='16-Point Star' or .//图:类别_8019='46'">star12</xsl:when>
      <xsl:when test=".//图:名称_801A='24-Point Star' or .//图:类别_8019='47'">star24</xsl:when>
      <xsl:when test=".//图:名称_801A='32-Point Star' or .//图:类别_8019='48'">star32</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Ribbon' or .//图:类别_8019='49'">ribbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Rectangular Callout' or .//图:类别_8019='51'">wedgeRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangular Callout' or .//图:类别_8019='52'">wedgeRoundRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval Callout' or .//图:类别_8019='53'">wedgeEllipseCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='54'">cloudCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(No Border)' or .//图:类别_8019='513'">borderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2' or .//图:类别_8019='56'">borderCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3' or .//图:类别_8019='57'">borderCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4' or .//图:类别_8019='58'">borderCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Accent Bar)' or .//图:类别_8019='59'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(Accent Bar)' or .//图:类别_8019='510'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Accent Bar)' or .//图:类别_8019='511'">accentCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Accent Bar)' or .//图:类别_8019='512'">accentCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1' or .//图:类别_8019='55'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(No Border)' or .//图:类别_8019='514'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(No Border)' or .//图:类别_8019='515'">callout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(No Border)' or .//图:类别_8019='516'">callout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Border and Accent Bar)' or .//图:类别_8019='517'">accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(Border and Accent Bar)' or .//图:类别_8019='518'">accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Border and Accent Bar)' or .//图:类别_8019='519'">accentBorderCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Border and Accent Bar)' or .//图:类别_8019='520'">accentBorderCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line' or .//图:类别_8019='61'">line</xsl:when>
      <xsl:when test=".//图:名称_801A='Arrow' or .//图:类别_8019='62'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Arrow' or .//图:类别_8019='63'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Curve' or .//图:类别_8019='64'">curvedConnector3</xsl:when>
      <xsl:when test=".//图:名称_801A='Straight Connector' or .//图:类别_8019='71'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Elbow Connector' or .//图:类别_8019='74'">bentConnector3</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Connector' or .//图:类别_8019='77'">curvedConnector3</xsl:when>
      <xsl:when test=".//图:名称_801A='Hexagon' or .//图:类别_8019='110'">hexagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Cross' or .//图:类别_8019='111'">plus</xsl:when>
      <xsl:when test=".//图:名称_801A='Regual Pentagon' or .//图:类别_8019='112'">pentagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Can' or .//图:类别_8019='113'">can</xsl:when>
      <xsl:when test=".//图:名称_801A='Cube' or .//图:类别_8019='114'">cube</xsl:when>
      <xsl:when test=".//图:名称_801A='Bevel' or .//图:类别_8019='115'">bevel</xsl:when>
      <xsl:when test=".//图:名称_801A='Folded Corner' or .//图:类别_8019='116'">foldedCorner</xsl:when>
      <xsl:when test=".//图:名称_801A='Smiley Face' or .//图:类别_8019='117'">smileyFace</xsl:when>
      <xsl:when test=".//图:名称_801A='Donut' or .//图:类别_8019='118'">donut</xsl:when>
      <xsl:when test=".//图:名称_801A='No Symbol' or .//图:类别_8019='119'">noSmoking</xsl:when>
      <xsl:when test=".//图:名称_801A='Block Arc' or .//图:类别_8019='120'">blockArc</xsl:when>
      <xsl:when test=".//图:名称_801A='Heart' or .//图:类别_8019='121'">heart</xsl:when>
      <xsl:when test=".//图:名称_801A='Lightning' or .//图:类别_8019='122'">lightningBolt</xsl:when>
      <xsl:when test=".//图:名称_801A='Sun' or .//图:类别_8019='123'">sun</xsl:when>
      <xsl:when test=".//图:名称_801A='Moon' or .//图:类别_8019='124'">moon</xsl:when>
      <xsl:when test=".//图:名称_801A='Arc' or .//图:类别_8019='125'">arc</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Bracket' or .//图:类别_8019='126'">bracketPair</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Brace' or .//图:类别_8019='127'">bracePair</xsl:when>
      <xsl:when test=".//图:名称_801A='Plaque' or .//图:类别_8019='128'">plaque</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Bracket' or .//图:类别_8019='129'">leftBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Bracket' or .//图:类别_8019='130'">rightBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Brace' or .//图:类别_8019='131'">leftBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Brace' or .//图:类别_8019='132'">rightBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='U-Turn Arrow' or .//图:类别_8019='210'">uturnArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Up Arrow' or .//图:类别_8019='211'">leftUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent-Up Arrow' or .//图:类别_8019='212'">bentUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Right Arrow' or .//图:类别_8019='213'">curvedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Left Arrow' or .//图:类别_8019='214'">curvedLeftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Arrow' or .//图:类别_8019='215'">curvedUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Arrow' or .//图:类别_8019='216'">curvedDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Striped Right Arrow' or .//图:类别_8019='217'">stripedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Notched Right Arrow' or .//图:类别_8019='218'">notchedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Pentagon Arrow' or .//图:类别_8019='219'">homePlate</xsl:when>
      <xsl:when test=".//图:名称_801A='Chevron Arrow' or .//图:类别_8019='220'">chevron</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Arrow Callout' or .//图:类别_8019='221'">rightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow Callout' or .//图:类别_8019='222'">leftArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow Callout' or .//图:类别_8019='223'">upArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow Callout' or .//图:类别_8019='224'">downArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow Callout' or .//图:类别_8019='225'">leftRightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Up-Down Arrow Callout' or .//图:类别_8019='226'">upDownArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Quad Arrow Callout' or .//图:类别_8019='227'">quadArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Circular Arrow' or .//图:类别_8019='228'">circularArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Preparation' or .//图:类别_8019='310'">flowChartPreparation</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Input' or .//图:类别_8019='311'">flowChartManualInput</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Operation' or .//图:类别_8019='312'">flowChartManualOperation</xsl:when>
      <xsl:when test=".//图:名称_801A='Connector' or .//图:类别_8019='313'">flowChartConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Off-page Connector' or .//图:类别_8019='314'">flowChartOffpageConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Card' or .//图:类别_8019='315'">flowChartPunchedCard</xsl:when>
      <xsl:when test=".//图:名称_801A='Punched Tape' or .//图:类别_8019='316'">flowChartPunchedTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Summing Junction' or .//图:类别_8019='317' ">flowChartSummingJunction</xsl:when>
      <xsl:when test=".//图:名称_801A='Or' or .//图:类别_8019='318'">flowChartOr</xsl:when>
      <xsl:when test=".//图:名称_801A='Collate' or .//图:类别_8019='319'">flowChartCollate</xsl:when>
      <xsl:when test=".//图:名称_801A='Sort' or .//图:类别_8019='320'">flowChartSort</xsl:when>
      <xsl:when test=".//图:类别_8019='321'">flowChartExtract</xsl:when>
      <xsl:when test=".//图:名称_801A='Merge'">flowChartMerge</xsl:when>
      <xsl:when test=".//图:名称_801A='Extract' or .//图:类别_8019='322'">flowChartMerge</xsl:when>
      <xsl:when test=".//图:名称_801A='Stored Data' or .//图:类别_8019='323'">flowChartOnlineStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Delay' or .//图:类别_8019='324'">flowChartDelay</xsl:when>
      <xsl:when test=".//图:名称_801A='Sequential Access Storage' or .//图:类别_8019='325'">flowChartMagneticTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Magnetic Disk' or .//图:类别_8019='326'">flowChartMagneticDisk</xsl:when>
      <xsl:when test=".//图:名称_801A='Direct Access Storage' or .//图:类别_8019='327'">flowChartMagneticDrum</xsl:when>
      <xsl:when test=".//图:名称_801A='Display' or .//图:类别_8019='328'">flowChartDisplay</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Ribbon' or .//图:类别_8019='410'">ribbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Ribbon' or .//图:类别_8019='411'">ellipseRibbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Ribbon' or .//图:类别_8019='412'">ellipseRibbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Vertical Scroll' or .//图:类别_8019='413'">verticalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Horizontal Scroll' or .//图:类别_8019='414'">horizontalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Wave' or .//图:类别_8019='415'">wave</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Wave' or .//图:类别_8019='416'">doubleWave</xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--Modified by LDM in 2011/01/07-->
  <!--图形的位置信息模板-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="PicLocation">
    <a:xfrm>
      <xsl:if test="./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D!='0.0'">
        <xsl:attribute name="rot">
          <xsl:choose >
            <xsl:when test ="./图:预定义图形_8018/图:名称_801A='Extract'">
        <!--yanghaojie-start-->       
			  <!--20130304 gaoyuwei Bug_2672_5 转换后文档需要修复 start-->
				<xsl:value-of select="round(21600000 - (180 - number(./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D))*60000-10800180)"/>	
				<!--end-->
        <!--yanghaojie-end-->
				
            </xsl:when>
            <xsl:when test ="./图:预定义图形_8018/图:名称_801A='Process'">
              <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D*60000*(-1))"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select="round(21600000 - (360 - ./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D)*60000)"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--<xsl:if test ="./图:预定义图形_8018/图:名称_801A='Process'">
            <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D*60000*(-1))"/>
          </xsl:if>
          <xsl:if test ="./图:预定义图形_8018/图:名称_801A='Extract'">
            <xsl:value-of select="round(21600000 - (360 - ./图:预定义图形_8018/图:属性_801D/图:旋转角度_804D)*60000)"/>
          </xsl:if>-->
        </xsl:attribute>
      </xsl:if>

          <xsl:if test="./图:翻转_803A">
            <xsl:call-template name="Flip"/>
          </xsl:if>
          <!---->
      <a:off>

        <xsl:if test="./图:组合位置_803B">
          <xsl:attribute name="x">
            <xsl:value-of select="round(./图:组合位置_803B/@x_C606 * 12700)"/>
          </xsl:attribute> 
          <xsl:attribute name="y">
            <xsl:value-of select="round(./图:组合位置_803B/@y_C607 * 12700)"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:variable name ="graRef">
          <xsl:value-of select ="./@标识符_804B"/>
        </xsl:variable>
        <xsl:if test="not(./图:组合位置_803B)">
          <xsl:if test ="./ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825//uof:锚点_C644[@图形引用_C62E = $graRef]">
            <xsl:for-each select ="./ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825//uof:锚点_C644[@图形引用_C62E = $graRef]">
            <xsl:attribute name="x">
              <xsl:value-of select="round(./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 * 12700)"/>
            </xsl:attribute>
            <xsl:attribute name="y">
              <xsl:value-of select="round(./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108 * 12700)"/>
            </xsl:attribute>
            </xsl:for-each>
          </xsl:if>
        </xsl:if>
        
      </a:off>
      <a:ext>
        <xsl:attribute name="cx">
          <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*12700)"/>
        </xsl:attribute>
        <xsl:attribute name="cy">
          <xsl:value-of select="round(.//图:属性_801D/图:大小_8060/@长_C604*12700)"/>
        </xsl:attribute>
      </a:ext>
    </a:xfrm>
    <!--
      <xsl:call-template name="prstGeom"/>
    --> 
  </xsl:template>

  <!--Need consideration-->
  <!--Marked by LDM in 2011/01/07-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="Flip">
    <xsl:if test="./图:翻转_803A='x'">
      <xsl:attribute name="flipH">1</xsl:attribute>
    </xsl:if>
    <xsl:if test="./图:翻转_803A='y'">
      <xsl:attribute name="flipV">1</xsl:attribute>
    </xsl:if>
  </xsl:template> 
  
  <!--Modified by LDM in 2011/01/07-->
  <!--预定义图形填充模板-->
  <xsl:template name="Fill_Drawing">
    <xsl:if test=".//图:填充/图:颜色 and not(.//图:填充/图:颜色 = 'auto')">
      <a:solidFill>
      <a:srgbClr>
        <xsl:attribute name="val">
          <xsl:value-of select="substring-after(.//图:填充/图:颜色,'#')"/>
        </xsl:attribute>
      </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <xsl:if test=".//图:填充/图:图案 and not(.//图:填充/图:图案 = 'auto')">
      <patternFill>
        <xsl:attribute name="patternType">
          <xsl:choose>
            <xsl:when test="图:填充/图:颜色">
              <xsl:value-of select="'solid'"/>
            </xsl:when>
            <xsl:when test="图:填充/图:图案">
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'solid'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="图:填充/图:颜色">
          <xsl:variable name="color" select="图:填充/图:颜色"/>
          <bgColor>
            <xsl:attribute name="rgb">
              <xsl:value-of select="concat('FF',$color)"/>
            </xsl:attribute>
          </bgColor>
        </xsl:if>
        <xsl:if test="表:填充/图:图案">
          <xsl:variable name="color2" select="表:填充/图:图案/@图:前景色"/>
          <xsl:variable name="color3" select="表:填充/图:图案/@图:背景色"/>
          <fgColor>
            <xsl:attribute name="rgb">
              <xsl:value-of select="concat('FF',$color2)"/>
            </xsl:attribute>
          </fgColor>
          <bgColor>
            <xsl:attribute name="rgb">
              <xsl:value-of select="concat('FF',$color3)"/>
            </xsl:attribute>
          </bgColor>
        </xsl:if>
      </patternFill>
    </xsl:if>
  </xsl:template>

  <!--Modified by LDM in 2011/01/07-->
  <!--预定义线型模板-->
  <!--李杨修改 2012-2-17-->
  <xsl:template name="Line_Drawing">
    <a:ln>
      <xsl:if test=".//图:线粗细_805C">
        <xsl:attribute name="w">
          <xsl:value-of select=".//图:线粗细_805C * 12700"/>
        </xsl:attribute>
        
        <!--20130318 gaoyuwei Bug_2748 uof-oo 功能测试 “虚线+三线”边框线不正确-->
		  <xsl:if test=".//图:线类型_8059/@线型_805A">
		  <xsl:attribute name="cmpd">
			  <xsl:choose>
				  <!--单线-->
			  <xsl:when test=".//图:线类型_8059/@线型_805A ='single'">sng</xsl:when>
				  <!--双线-->
			  <xsl:when test=".//图:线类型_8059/@线型_805A ='double'">dbl</xsl:when>
				  <!--三线-->
			  <xsl:when test=".//图:线类型_8059/@线型_805A ='thick-between-thin'">tri</xsl:when>
				  <!--上粗下细-->
			  <xsl:when test=".//图:线类型_8059/@线型_805A ='thin-thick'">thickThin</xsl:when>
				  <!--上细下粗-->		  
			  <xsl:when test=".//图:线类型_8059/@线型_805A ='thick-thin'">thinThick</xsl:when>
			  </xsl:choose>
		  </xsl:attribute>
		</xsl:if>
		  <!--end-->
        
      </xsl:if>
      <xsl:if test=".//图:线颜色_8058">
        <a:solidFill>
          <a:srgbClr>
            <xsl:choose >
              <xsl:when test =".//图:线颜色_8058='auto'">
                <xsl:attribute name ="val">
                  <xsl:value-of select ="'000000'"/>
                </xsl:attribute>
              </xsl:when>
              <!--添加颜色，图案选取。 李杨2012-2-25-->
              
             	<!--20130318 gaoyuwei Bug_2747_2 uof-oo 功能测试 预定义图形4，5样式不正确-->
              <xsl:when test ="./图:预定义图形_8018/图:属性_801D/图:三维效果_8061 and ./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:颜色_8004 and not(./图:预定义图形_8018/图:名称_801A='Or')" >
				  <!--end-->
				  
                <xsl:attribute name ="val">
                  <xsl:value-of select="substring-after(./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:颜色_8004,'#')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test ="./图:预定义图形_8018/图:属性_801D/图:三维效果_8061 and ./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图案_800A">
                <xsl:attribute name ="val">
                  <xsl:value-of select ="substring-after(./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图案_800A/@前景色_800B,'#')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise >
                <xsl:attribute name ="val">
                  <xsl:value-of select="substring-after(.//图:线颜色_8058,'#')"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </a:srgbClr>
      </a:solidFill>
      </xsl:if>
        <!--边框线型-->
      <!--线类型中的 虚实属性没有none值，线型属性有none值。 待考虑 李杨2012-2-17-->
        <xsl:if test=".//图:线类型_8059[@虚实_805B] and not(.//图:线类型_8059/@虚实_805B = 'none')">
        <xsl:variable name="dashType">
          <xsl:value-of select=".//图:线类型_8059/@虚实_805B"/>
        </xsl:variable>
          <xsl:call-template name="LineDashMapping_Border_2">
            <xsl:with-param name="dashType" select="$dashType"/>
          </xsl:call-template>
        </xsl:if>
      <xsl:if test=".//图:前端箭头_805E">
        <a:headEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:headEnd>
      </xsl:if>
      <xsl:if test=".//图:后端箭头_805F">
        <a:tailEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:tailEnd>
      </xsl:if>
    </a:ln>
  </xsl:template>

  <!--Modified by LDM in 2011/01/07-->
  <!--线型类型对应模板-->
  <!--LineTypeMapping_Border-->
  <xsl:template name="LineTypeMapping_Border">
    <xsl:param name ="lineType"/>
    <xsl:choose>
      <xsl:when test="$lineType = 'single'">
        <xsl:value-of select="'sng'"/>
      </xsl:when>
      <xsl:when test="$lineType = 'double'">
        <xsl:value-of select="'dbl'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!--Modified by LDM in 2010/12/29-->
  <!--线形模板-->
  <!--LineTypeMapping_Border-->
  <xsl:template name="LineDashMapping_Border_2">
    <xsl:param name="dashType"/>
    <xsl:choose>
      <!--单线-->
      <xsl:when test ="not($dashType = 'none')">
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:choose>
              <!--实线-->
              <xsl:when test="$dashType = 'noDash' or $dashType = 'solid'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <!--虚线-->
              <xsl:when test="$dashType = 'square-dot'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <!--虚线-->
              <xsl:when test="$dashType = 'round-dot'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <!--划线-->
              <xsl:when test="$dashType = 'dash'">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <!--长划线-->
              <xsl:when test="$dashType = 'long-dash'">
                <xsl:value-of select="'lgDash'"/>
              </xsl:when>
              <!--点划线-->
              <xsl:when test="$dashType = 'dash-dot'">
                <xsl:value-of select="'lgDashDot'"/>
              </xsl:when>
              <!--双点划线-->
              <xsl:when test="$dashType = 'dash-dot-dot'">
                <xsl:value-of select="'lgDashDotDot'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'solid'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:prstDash>
      </xsl:when>
      <xsl:otherwise>
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:value-of select="'solid'"/>
          </xsl:attribute>
        </a:prstDash>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!--Need Consideration====================================================================================================-->
  <!--Modified by LDM in 2010/12/16-->
  <!--这是啥意思？-->
  <!--
      <xsl:choose>
        <xsl:when test="($sername!='' and not(contains($sername,$relsheet))) or ($servalue!='' and not(contains($servalue,$relsheet)))">
          <xdr:absoluteAnchor xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
            <xdr:pos>
              <xsl:attribute name="x">
                <xsl:value-of select="$x * 12700"/>
              </xsl:attribute>
              <xsl:attribute name="y">
                <xsl:value-of select="$y * 12700"/>
              </xsl:attribute>
            </xdr:pos>
            <xdr:ext>
              <xsl:attribute name="cx">
                <xsl:value-of select="$w * 12700"/>
              </xsl:attribute>
              <xsl:attribute name="cy">
                <xsl:value-of select="$h * 12700"/>
              </xsl:attribute>
            </xdr:ext>
            <xdr:graphicFrame macro="">
              <xdr:nvGraphicFramePr>
                <xdr:cNvPr>
                  <xsl:attribute name="id">
                    <xsl:value-of select="@uof:chartid"/>
                  </xsl:attribute>
                  <xsl:attribute name="name">
                    <xsl:value-of select="'图表'"/>
                  </xsl:attribute>
                </xdr:cNvPr>
                <xdr:cNvGraphicFramePr/>
              </xdr:nvGraphicFramePr>
              <xdr:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="0" cy="0"/>
              </xdr:xfrm>
              <a:graphic>
                <a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart">
                  <c:chart xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                    <xsl:variable name="rid">
                      <xsl:value-of select="@uof:chartid"/>
                    </xsl:variable>
                    <xsl:attribute name="r:id">
                      <xsl:value-of select="concat('rId',$rid,'chart')"/>
                    </xsl:attribute>
                  </c:chart>
                </a:graphicData>
              </a:graphic>
            </xdr:graphicFrame>
            <xdr:clientData/>
          </xdr:absoluteAnchor>
        </xsl:when>
        <xsl:otherwise>
        -->

<!--添加三维效果。李杨2012-2-24-->
  <xsl:template name="图:三维效果_8061">
    <a:scene3d>
      <a:camera>
        <xsl:attribute name="prst">
          <xsl:if test="uof:方向_C63C">
            <xsl:call-template name="scene3dprst"/>
          </xsl:if>
          <xsl:if test="not(uof:方向_C63C)">
            <xsl:value-of select="'orthographicFront'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:if test="not(uof:角度_C635/uof:x方向_C636=0 and uof:角度_C635/uof:y方向_C637=0)">
          <a:rot>
            <xsl:attribute name="lon">
            
             <!--20130319 gaoyuwei bug_2752_1 uof-oo 功能测试 艺术字文档需要修复-->
			<xsl:choose>
				<xsl:when test ="../../图:名称_801A='Fade Right'">
				  <xsl:value-of select="'0'"/>
				</xsl:when>
				<xsl:otherwise>
				  <xsl:value-of select="uof:角度_C635/uof:x方向_C636 * 60000"/>
				</xsl:otherwise>
			</xsl:choose>
			 <!--end-->
			 
            </xsl:attribute>
            <xsl:attribute name="lat">
              <xsl:value-of select ="'0'"/>
              <!--<xsl:value-of select="uof:角度_C635/uof:y方向_C637 * 60000"/>-->
            </xsl:attribute>
            <xsl:attribute name="rev">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </a:rot>
        </xsl:if>
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">
          <xsl:if test="uof:照明_C638/uof:强度_C63A='bright'">
            <xsl:value-of select="'threePt'"/>
          </xsl:if>
          <xsl:if test="uof:照明_C638/uof:强度_C63A='dim'">
            <xsl:value-of select="'harsh'"/>
          </xsl:if>
          <xsl:if test="uof:照明_C638/uof:强度_C63A='normal'">
            <xsl:value-of select="'threePt'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:attribute name="dir">
          <xsl:value-of select="'t'"/>
        </xsl:attribute>
        <xsl:if test="not(round(uof:照明_C638/uof:角度_C639)=0)">
          <a:rot>
            <xsl:attribute name="lat">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
            <xsl:attribute name="lon">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
            <xsl:attribute name ="rev">
              <xsl:value-of select="uof:照明_C638/uof:角度_C639"/>
            </xsl:attribute>
          </a:rot>
        </xsl:if>

      </a:lightRig>
    </a:scene3d>
    <a:sp3d>
      <xsl:if test="uof:深度_C63B">
        <xsl:attribute name="extrusionH">
          <xsl:value-of select="uof:深度_C63B * 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="uof:表面效果_C63E">
        <xsl:attribute name="prstMaterial">
          <xsl:if test="uof:表面效果_C63E='metal'">
            <xsl:value-of select="'metal'"/>
          </xsl:if>
          <xsl:if test="uof:表面效果_C63E='plastic'">
            <xsl:value-of select="'plastic'"/>
          </xsl:if>
          <xsl:if test="uof:表面效果_C63E='matte'">
            <xsl:value-of select="'matte'"/>
          </xsl:if>
          <xsl:if test="uof:表面效果_C63E='wire-frame'">
            <xsl:value-of select="'legacyWireframe'"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="uof:颜色_C63F">
        <a:extrusionClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <!--<xsl:value-of select ="uof:颜色_C63F"/>-->
              <xsl:value-of select="substring-after(uof:颜色_C63F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:extrusionClr>
      </xsl:if>
      <xsl:if test ="not(uof:颜色_C63F)">

        <xsl:choose >
          <xsl:when test ="../图:填充_804C/图:颜色_8004">
            <a:extrusionClr>
              <a:srgbClr>
                <xsl:attribute name ="val">
                  <xsl:value-of select ="substring-after(../图:填充_804C/图:颜色_8004,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:extrusionClr>
          </xsl:when>
          <xsl:when test ="../图:填充_804C/图:图案_800A">
            <a:extrusionClr>
              <a:srgbClr>
                <xsl:attribute name ="val">
                  <xsl:value-of select ="substring-after(../图:填充_804C/图:图案_800A/@前景色_800B,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:extrusionClr>
          </xsl:when>
        </xsl:choose>

      </xsl:if>
    </a:sp3d>
  </xsl:template>
  <xsl:template name="scene3dprst">
    <xsl:if test="not(uof:样式_C634)">
      <xsl:if test="uof:方向_C63C/uof:角度_C639='none' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'orthographicFront'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'obliqueTopLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'perspectiveBelow'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'perspectiveRight'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'perspectiveLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'obliqueBottomLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'isometricLeftDown'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom' and uof:方向_C63C/uof:方式_C63D='parallel'">
        <xsl:value-of select="'perspectiveAbove'"/>
      </xsl:if>

      <xsl:if test="uof:方向_C63C/uof:角度_C639='none' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'perspectiveFront'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'obliqueTopLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'perspectiveBelow'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-top-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'perspectiveRight'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'perspectiveLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'obliqueBottomLeft'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'perspectiveAbove'"/>
      </xsl:if>
      <xsl:if test="uof:方向_C63C/uof:角度_C639='to-bottom-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
        <xsl:value-of select="'obliqueBottomRight'"/>
      </xsl:if>
    </xsl:if>
    <xsl:if test="uof:样式_C634">
      <xsl:if test="uof:样式_C634='0'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='1'">
        <xsl:value-of select="'obliqueTopLeft'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='2'">
        <xsl:value-of select="'obliqueBottomLeft'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='3'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='4'">
        <xsl:value-of select="'perspectiveContrastingLeftFacing'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='5'">
        <xsl:value-of select="'perspectiveContrastingRightFacing'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='6'">
        <xsl:value-of select="'orthographicFront'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='7'">
        <xsl:value-of select="'obliqueBottomLeft'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='8'">
        <xsl:value-of select="'perspectiveHeroicExtremeLeftFacing'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='9'">
        <xsl:value-of select="'perspectiveHeroicExtremeRightFacing'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='10'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='11'">
        <xsl:value-of select="'perspectiveBelow'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='12'">
        <xsl:value-of select="'perspectiveRelaxed'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='13'">
        <xsl:value-of select="'perspectiveRelaxed'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='14'">
        <xsl:value-of select="'perspectiveAbove'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='15'">
        <xsl:value-of select="'obliqueBottomLeft'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='16'">
        <xsl:value-of select="'isometricOffAxis2Left'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='17'">
        <xsl:value-of select="'isometricOffAxis1Right'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='18'">
        <xsl:value-of select="'obliqueTopRight'"/>
      </xsl:if>
      <xsl:if test="uof:样式_C634='19'">
        <xsl:value-of select="'perspectiveBelow'"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  
  <!--添加艺术字。李杨 2012-2-28-->
  <xsl:template match ="图:艺术字_802D">
    <xdr:txBody>
      <a:bodyPr wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720">
        <a:spAutoFit/>
      </a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr>
            <xsl:attribute name ="lang">zh-CN</xsl:attribute>
            <xsl:attribute name ="altLang">en-US</xsl:attribute>
            <xsl:for-each select ="图:字体_802E/@字号_412D">
              <xsl:attribute name ="sz">
                <xsl:value-of select="round(number(current()) * 100)"/>
              </xsl:attribute>
            </xsl:for-each>
            <xsl:attribute name ="b">1</xsl:attribute>
            <xsl:attribute name ="cap">none</xsl:attribute>
            <xsl:attribute name ="spc">0</xsl:attribute>
            <a:ln>
              <a:prstDash val="solid"/>
            </a:ln>
            <!--字体颜色 李杨2012-2-29-->
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name ="val">
                  <!--20130319 gaoyuwei bug_2752_3 uof-oo 功能测试 艺术字文档需要修复-->
					<xsl:choose>
						<xsl:when test ="not(../图:填充_804C/图:渐变_800D/@类型_8008) and ../图:填充_804C/图:颜色_8004">
							<xsl:value-of select ="substring-after(../图:填充_804C/图:颜色_8004,'#')"/>
						</xsl:when>
						<xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008 and not(../图:填充_804C/图:颜色_8004)">
							<xsl:call-template name="frontClr"/>
						</xsl:when>
						<xsl:when test ="not(../图:填充_804C/图:渐变_800D/@类型_8008) and not(../图:填充_804C/图:颜色_8004) and ../图:填充_804C/图:图片_8005/@名称_8009">
							<xsl:call-template name="frontClr"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:if test ="../../图:名称_801A='Plafinalintext'">
								<!--金黄-->
								<xsl:value-of select ="'FFD700'"/>
							</xsl:if>
							<xsl:if test ="../../图:名称_801A='Fade Right'">
								<!--金黄-->
								<xsl:value-of select ="'FFD700'"/>
							</xsl:if>
							<xsl:if test ="../../图:名称_801A='Inflate Top'">
								<!--黄-->
								<xsl:value-of select ="'FF7D40'"/>
							</xsl:if>
							<xsl:if test ="../../图:类别_8019='549'">
								<!--红-->
								<xsl:value-of select ="FF6347"/>
							</xsl:if>
							<xsl:if test ="../../图:类别_8019='547'">
								<!--粉色-->
								<xsl:value-of select ="FFC0CB"/>
							</xsl:if>
						</xsl:otherwise>
					</xsl:choose>
					<!--end-->
					
                </xsl:attribute>
              </a:srgbClr>
              
            </a:solidFill>
            <a:effectLst>
              <a:outerShdw blurRad="88000" dist="50800" dir="5040000" algn="tl">
                <!--<a:schemeClr>
                  <xsl:attribute name ="val">
                    <xsl:value-of select ="substring-after(../图:阴影_8051/@颜色_C61E,'#')"/>
                  </xsl:attribute>
                </a:schemeClr>-->
				  <!--20130319 gaoyuwei bug_2752_2 uof-oo 功能测试 艺术字文档需要修复-->
				  <a:scrgbClr r="0" g="0" b="0"/>
				  <!--end-->
              </a:outerShdw>
            </a:effectLst>
          </a:rPr>
          <a:t>
            <xsl:value-of select ="图:艺术字文本_8036"/>
          </a:t>
        </a:r>
      </a:p>
    </xdr:txBody>
  </xsl:template>

  <!--添加艺术字字体颜色模板。李杨2012-2-29-->
  <xsl:template name ="frontClr">
    <xsl:choose>
      <!--黄-->
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='gold'">FFE384</xsl:when>
      <xsl:when test ="../图:填充_804C/图:图片_8005/@名称_8009='木片'">FF8000</xsl:when>
      <!--绿-->
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='grassland'">228B22</xsl:when>
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='rainbow'">228B22</xsl:when>
      <xsl:when test ="../图:填充_804C/图:图片_8005/@名称_8009='月影'">8B864E</xsl:when>
      <!--灰-->
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='misty'">CCCCCC</xsl:when>
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='aluminum'">CCCCCC</xsl:when>
      <!--蓝-->
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='sea-green'">0000FF</xsl:when>
      <xsl:when test ="../图:填充_804C/图:图片_8005/@名称_8009='Wave2'">0000FF</xsl:when>
      <xsl:when test ="../图:填充_804C/图:图片_8005/@名称_8009='水波'">0000FF</xsl:when>
      <xsl:when test ="../图:填充_804C/图:图片_8005/@名称_8009='方格'">0000FF</xsl:when>
      <!--红-->
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='brown-coffee'">990033</xsl:when>
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='dark-red'">990033</xsl:when>
      <xsl:when test ="../图:填充_804C/图:渐变_800D/@类型_8008='polar-lights'">E3170D</xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--添加预定义图形的阴影。李杨2012-3-1-->
  <xsl:template name ="阴影single">
    <a:effectLst>
      <xsl:choose >
        <!--左下-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606='-6.0' and @y_C607='6.0']">
          <a:outerShdw blurRad="50800" dist="127000" dir="8100000" algn="tr" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <!--左上-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606='-6.0' and @y_C607='-6.0']">
          <a:outerShdw blurRad="50800" dist="127000" dir="13500000" algn="br" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <!--右上-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606='6.0' and @y_C607='-6.0']">
          <a:outerShdw blurRad="50800" dist="127000" dir="18900000" algn="bl" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <!--右下-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606='6.0' and @y_C607='6.0']">
          <a:outerShdw blurRad="50800" dist="127000" dir="2700000" algn="tl" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <!--其他类型左上-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606 &lt;='-25.0' and @x_C606 &gt;='-50.0' and @y_C607&lt;='-25.0' and @y_C607 &gt;='-50.0']">
          <a:outerShdw blurRad="50800" dist="190500" dir="13200000" sx="122000" sy="122000" algn="br" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <xsl:when test ="uof:偏移量_C61B[@x_C606='-3.0' and @y_C607='-3.0']">
          <a:outerShdw blurRad="50800" dist="63500" dir="13500000" algn="br" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:when>
        <!--其他类型右下-->
        <xsl:when test ="uof:偏移量_C61B[@x_C606='2.0' and @y_C607='2.0']">
          <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
          
       		<!--20130318 gaoyuwei bug_2747_3 uof-oo 功能测试 预定义图形4，5样式不正确-->
			<!--默认左 特殊无三维图Or-->
        </xsl:when>
		  <xsl:when test ="uof:偏移量_C61B[@x_C606='-3.0' and @y_C607='0.0']">
			  <a:outerShdw blurRad="50800" dist="38100" dir="10800000" algn="r" rotWithShape="0">
				  <a:srgbClr val="00B050">
					  <a:alpha val="40%"/>
				  </a:srgbClr>
			  </a:outerShdw>
		  </xsl:when>
		  <!--默认左 特殊无三维图Collate-->
		  <xsl:when test ="uof:偏移量_C61B[@x_C606='0.0' and @y_C607='2.0']">
			  <a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0">
				  <a:prstClr val="black">
					  <a:alpha val="40%"/>
				  </a:prstClr>
			  </a:outerShdw>
		  </xsl:when>
	  
            <!--end-->
        <xsl:otherwise >
          <a:outerShdw>
            <xsl:attribute name ="blurRad">50800</xsl:attribute>
            <xsl:attribute name ="dist">381000</xsl:attribute>
            <xsl:choose >
              <xsl:when test ="uof:偏移量_C61B[@x_C606 &gt;= '0.0' and @y_C607 &gt;= '0.0']">
                <xsl:attribute name ="dir">2700000</xsl:attribute>
                <xsl:attribute name ="algn">tl</xsl:attribute>
              </xsl:when>
              <xsl:when test ="uof:偏移量_C61B[@x_C606 &gt;= '0.0' and @y_C607 &lt;= '0.0']">
                <xsl:attribute name ="dir">18900000</xsl:attribute>
                <xsl:attribute name ="algn">bl</xsl:attribute>
              </xsl:when>
              <xsl:when test ="uof:偏移量_C61B[@x_C606 &lt;= '0.0' and @y_C607 &gt;= '0.0']">
                <xsl:attribute name ="dir">8100000</xsl:attribute>
                <xsl:attribute name ="algn">tr</xsl:attribute>
              </xsl:when>
              <xsl:when test ="uof:偏移量_C61B[@x_C606 &lt;= '0.0' and @y_C607 &lt;= '0.0']">
                <xsl:attribute name ="dir">13500000</xsl:attribute>
                <xsl:attribute name ="algn">br</xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:attribute name ="rotWithShape">0</xsl:attribute>
            <a:prstClr val="black">
              <a:alpha val="40%"/>
            </a:prstClr>
          </a:outerShdw>
        </xsl:otherwise>
      </xsl:choose>
    </a:effectLst>
  </xsl:template>
  <!--添加预定义图形的阴影。李杨2012-3-1-->
  <xsl:template name ="阴影ShapePer">
    <a:effectLst>
      <a:outerShdw>
        <xsl:attribute name ="blurRad">76200</xsl:attribute>
        <xsl:choose >
          <xsl:when test ="uof:偏移量_C61B[@x_C606 &lt;= '0.0' and @y_C607 &gt;= '0.0' and @y_C607 &lt;= '90.0']">
            <xsl:attribute name ="dir">13500000</xsl:attribute>
            <xsl:attribute name ="sy">23000</xsl:attribute>
            <xsl:attribute name ="kx">1200000</xsl:attribute>
            <xsl:attribute name ="algn">br</xsl:attribute>
          </xsl:when>
          <xsl:when test ="uof:偏移量_C61B[@x_C606 &gt;= '0.0' and @y_C607 &gt;= '0.0' and @y_C607 &lt;= '90.0']">
            <xsl:attribute name ="dir">18900000</xsl:attribute>
            <xsl:attribute name ="sy">23000</xsl:attribute>
            <xsl:attribute name ="kx">-1200000</xsl:attribute>
            <xsl:attribute name ="algn">bl</xsl:attribute>
          </xsl:when>
          <xsl:when test ="uof:偏移量_C61B[@x_C606 &lt;= '0.0' and @y_C607 &gt;= '90.0']">
            <xsl:attribute name ="dist">12700</xsl:attribute>
            <xsl:attribute name ="dir">8100000</xsl:attribute>
            <xsl:attribute name ="sy">-23000</xsl:attribute>
            <xsl:attribute name ="kx">800400</xsl:attribute>
            <xsl:attribute name ="algn">br</xsl:attribute>
          </xsl:when>
          <xsl:when test ="uof:偏移量_C61B[@x_C606 &gt;= '0.0' and @y_C607 &gt;= '90.0']">
            <xsl:attribute name ="dist">12700</xsl:attribute>
            <xsl:attribute name ="dir">2700000</xsl:attribute>
            <xsl:attribute name ="sy">-23000</xsl:attribute>
            <xsl:attribute name ="kx">-800400</xsl:attribute>
            <xsl:attribute name ="algn">bl</xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:attribute name ="rotWithShape">0</xsl:attribute>
        <a:prstClr val="black">
          <a:alpha val="20%"/>
        </a:prstClr>
      </a:outerShdw>
    </a:effectLst>
  </xsl:template>
</xsl:stylesheet>
