<?xml version="1.0" encoding="UTF-8"?>
<!--cxl,2012.2.29增加预处理第二步，将分节与之后的元素互换位置-->
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
              <xsl:apply-templates select="字:文字处理文档_4225" mode="preRev2"/>
            </uof:文字处理>
          </xsl:when>
          <xsl:when test="name(.)='uof:扩展区'">
            <xsl:copy-of select="."/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </uof:UOF>   
  </xsl:template>
  <xsl:template match="字:文字处理文档_4225" mode="preRev2">
    <字:文字处理文档_4225>
      <xsl:for-each select="字:分节_416A">
        <xsl:variable name="id1" select="generate-id(.)"/>
        <xsl:variable name="id" select="generate-id(following-sibling::字:分节_416A[1])"/>
        <xsl:if test="following-sibling::node()[generate-id(preceding-sibling::字:分节_416A[1]) = $id1 and generate-id(.) != $id and generate-id(.) != $id1 and name(.)!='字:分节_416A']">
          <xsl:apply-templates select="following-sibling::node()[generate-id(preceding-sibling::字:分节_416A[1]) = $id1 and generate-id(.) != $id and generate-id(.) != $id1 and name(.)!='字:分节_416A']" mode="preRev2"/>
        </xsl:if>
        <xsl:copy-of select="."/>     
      </xsl:for-each>
    </字:文字处理文档_4225>
  </xsl:template>
  <xsl:template match="node()" mode="preRev2">
    <xsl:copy-of select="."/>
  </xsl:template>
</xsl:stylesheet>

