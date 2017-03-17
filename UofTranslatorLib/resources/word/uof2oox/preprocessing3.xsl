<?xml version="1.0" encoding="UTF-8"?>
<!--cxl,2012.2.29增加预处理第三步，将最后一个分节之前的所有分节放到段落属性标签内-->
<xsl:stylesheet version="1.0"  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  <xsl:template match="//uof:UOF">
    <uof:UOF>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='uof:元数据'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='uof:链接集'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='uof:式样集'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='uof:对象集'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='w:media'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='uof:文字处理'">            
            <uof:文字处理>
              <xsl:attribute name="id">
                <xsl:value-of select="'word'"/>
              </xsl:attribute>
              <xsl:copy-of select="./规则:公用处理规则_B665"/>
              <xsl:apply-templates select="字:文字处理文档_4225" mode="preRev3"/>
            </uof:文字处理>
          </xsl:when>
          <xsl:when test="name(.)='uof:扩展区'">
            <xsl:copy-of select="."/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </uof:UOF>
  </xsl:template>
  <xsl:template match="字:文字处理文档_4225" mode="preRev3">
      <xsl:variable name="sectNO" select="count(//字:分节_416A)"/>
      <xsl:if test="$sectNO=1">
        <xsl:copy-of select="."/>
      </xsl:if>
      <xsl:if test="$sectNO!=1">
        <字:文字处理文档_4225>
          <xsl:for-each select="node()">
            <xsl:choose>
              <xsl:when test="name(.)='字:分节_416A'">
                <xsl:variable name="position" select="count(preceding-sibling::字:分节_416A)+1"/>
                <xsl:comment>
                  <xsl:value-of select="$position"/>
                </xsl:comment>
                <xsl:if test="$position!= $sectNO">
                  <字:段落_416B>
                    <字:段落属性_419B>
                      <xsl:copy-of select="."/>
                    </字:段落属性_419B>
                  </字:段落_416B>
                </xsl:if>
                <xsl:if test="$position=$sectNO">
                  <xsl:copy-of select="."/>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:copy-of select="."/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:for-each>
        </字:文字处理文档_4225>        
      </xsl:if>       
  </xsl:template>
 
</xsl:stylesheet>

