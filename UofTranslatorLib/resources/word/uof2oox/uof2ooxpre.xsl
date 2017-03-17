<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <xsl:import href="preprocessing1.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
    <xsl:output method="xml" indent="yes"/>
  <xsl:template match="/">
    <xsl:variable name="graphics">
      <xsl:if test="document('content.xml')//uof:锚点_C644">
        <xsl:value-of select="'graphicsexist'"/>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="objectdata">
      <xsl:if test="$graphics='graphicsexist'"><!--2013-04-23，wudi，增加自动编号图片符号引用-->
        <xsl:if test="document('graphics.xml')//图:图片数据引用_8037 or document('graphics.xml')//图:填充_804C/图:图片_8005 or document('styles.xml')//式样:自动编号集_990E//字:图片符号_411B">
          <xsl:value-of select="'objectexist'"/>
        </xsl:if>
      </xsl:if>
      <xsl:if test="$graphics!='graphicsexist'"><!--cxl,2012.3.20增加自动编号图片符号引用-->
        <xsl:if test="document('content.xml')/字:文字处理文档_4225/字:分节_416A/字:节属性_421B/字:填充_4134/图:图片_8005 or document('styles.xml')//式样:自动编号集_990E//字:图片符号_411B">
          <xsl:value-of select="'objectexist'"/>
        </xsl:if>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="hyperlinks">
      <xsl:if test="document('content.xml')//字:区域开始_4165[@类型_413B='hyperlink']">
        <xsl:value-of select="document('hyperlinks.xml')"/>
      </xsl:if>
    </xsl:variable>
    <uof:UOF>
      <uof:元数据 id="_meta/meta.xml">
        <xsl:copy-of select="document('_meta/meta.xml')"/>
      </uof:元数据>
      <xsl:if test="$hyperlinks!=''">
        <uof:链接集 id="hyperlinks.xml">
          <xsl:copy-of select="document('hyperlinks.xml')//超链:超级链接_AA0C"/>
        </uof:链接集>
      </xsl:if>
      <uof:式样集 id="styles.xml">
        <xsl:copy-of select="document('styles.xml')/式样:式样集_990B/*"/>
      </uof:式样集>
      <xsl:if test="$graphics='graphicsexist' or $objectdata='objectexist'">
        <uof:对象集 id="object">
          <xsl:if test="$objectdata='objectexist'">
            <xsl:copy-of select="document('objectdata.xml')//对象:对象数据_D701"/>
          </xsl:if>
          <xsl:if test="$graphics='graphicsexist'">
            <xsl:copy-of select="document('graphics.xml')//图:图形_8062"/>
          </xsl:if>          
        </uof:对象集>
      </xsl:if>
      <uof:文字处理 id="word">
        <xsl:apply-templates select="document('rules.xml')/规则:公用处理规则_B665" mode="preRev"/>
        <字:文字处理文档_4225>
          <xsl:apply-templates select="document('tblVertical.xml')/字:文字处理文档_4225" mode="preRev"/>
          <!--<xsl:apply-templates select="document('content.xml')/字:文字处理文档_4225/字:分节_416A[1]" mode="preRev"/>-->      
        </字:文字处理文档_4225>   
      </uof:文字处理>
      <uof:扩展区 id="extend">
        <xsl:copy-of select="document('extend.xml')/扩展:扩展区_B200"/>
      </uof:扩展区>
    </uof:UOF>
  </xsl:template>
 
</xsl:stylesheet>
