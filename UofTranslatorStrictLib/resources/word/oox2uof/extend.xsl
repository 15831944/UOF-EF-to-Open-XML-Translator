<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="2.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
  xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:a15="http://schemas.microsoft.com/office/drawing/2012/main"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special">
  <xsl:import href="paragraph.xsl"/>
    <xsl:output method="xml" indent="yes"/>
    
  <xsl:template name="shdExtend">
    <扩展:扩展_B201>
      <扩展:软件名称_B202>Yozo Office</扩展:软件名称_B202>
      <扩展:软件版本_B203>2011</扩展:软件版本_B203>
      <扩展:扩展内容_B204>
        <扩展:路径_B205>图形集/图形/预定义图形/属性/阴影</扩展:路径_B205>
        <扩展:内容_B206>
          <xsl:for-each select="//a:effectLst">
            <xsl:call-template name="shdExtendItem"/>
          </xsl:for-each>
        </扩展:内容_B206>
      </扩展:扩展内容_B204>
    </扩展:扩展_B201>
  </xsl:template>

  <xsl:template name ="shdExtendItem">
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="//v:rect | //wp:anchor | //wp:inline"/>
    </xsl:variable>
    <扩展:阴影_F02A>
      <xsl:attribute name="对象标识符引用_F02C">
        <xsl:if test="//wpg:grpSp">
          <!--组合图形，未做案例，有待验证-->
          <xsl:value-of select="concat('grpspObj',./wps:cNvPr/@id)"/>
        </xsl:if>
        <xsl:if test="not(//wpg:grpSp)">
          <xsl:value-of select="concat('documentObj',$number * 2 +1)"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="阴影类型_F02B">1</xsl:attribute>
    </扩展:阴影_F02A>
  </xsl:template>
  
  <!--cxl,2012.3.8新增首字下沉转换（扩展区）-->
  <xsl:template name="framePrExtend">
    <扩展:扩展_B201>
      <扩展:软件名称_B202>WPS Office</扩展:软件名称_B202>
      <扩展:软件版本_B203>2009</扩展:软件版本_B203>
      <扩展:扩展内容_B204>
        <xsl:for-each select="//w:p">
          <xsl:variable name="pLocation" >
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:for-each select="node()">
            <xsl:variable name="pPrLocation">
              <xsl:value-of select="position()"/>
            </xsl:variable>
            <xsl:if test="./w:framePr">
              <扩展:路径_B205>
                <xsl:value-of select="concat(concat(concat(concat('字:文字处理文档_4225/字:段落_416B[',$pLocation) , ']/字:段落属性_419B[' ),$pPrLocation),']')"/>
              </扩展:路径_B205>
              <扩展:内容_B206>
                <扩展:首字扩展属性_B43E>
                  <xsl:apply-templates select="self::w:pPr[position()=1]">
                    <xsl:with-param name="framePrLocation" select="'1'"></xsl:with-param>
                  </xsl:apply-templates>
                </扩展:首字扩展属性_B43E>
              </扩展:内容_B206>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </扩展:扩展内容_B204>
    </扩展:扩展_B201>
  </xsl:template>
</xsl:stylesheet>
