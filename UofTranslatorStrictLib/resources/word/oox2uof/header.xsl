<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:fo="http://www.w3.org/1999/XSL/Format" 
xmlns:xs="http://www.w3.org/2001/XMLSchema" 
xmlns:fn="http://www.w3.org/2005/xpath-functions" 
xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" 
xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <!--
  <xsl:import href="table.xsl"/>
  <xsl:import href="paragraph.xsl"/>-->

  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <!--cxl修改于2011/11/11-->
  <xsl:template name="header">
    <字:页眉_41F3>
      <xsl:for-each select="./w:headerReference">
        <xsl:variable name="temp" select="@r:id"/>
        <xsl:variable name="target" select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$temp]/@Target"/>
     
        <xsl:choose>
          <!--偶页页眉-->
          <xsl:when test="@w:type='even'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:hdr" mode="evenHeader"/>
          </xsl:when>
          <!--奇页页眉-->
          <xsl:when test="@w:type='default'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:hdr" mode="oddHeader"/>
          </xsl:when>
          <!--首页页眉-->
          <xsl:when test="@w:type='first'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:hdr" mode="coverHeader"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:页眉_41F3>
  </xsl:template>

  <xsl:template match="w:hdr" mode="coverHeader">
    <字:首页页眉_41F6>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:首页页眉_41F6>
  </xsl:template>

  <xsl:template match="w:hdr" mode="oddHeader">
    <字:奇数页页眉_41F4>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:奇数页页眉_41F4>
  </xsl:template>

  <xsl:template match="w:hdr" mode="evenHeader">
    <字:偶数页页眉_41F5>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:偶数页页眉_41F5>
  </xsl:template>

</xsl:stylesheet>
