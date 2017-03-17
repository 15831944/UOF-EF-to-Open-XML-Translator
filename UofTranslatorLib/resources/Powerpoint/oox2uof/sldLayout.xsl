<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
			  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
			  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
			  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
			  xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:规则="http://schemas.uof.org/cn/2009/rules">
  <xsl:import href="ph.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<!--子元素 UOF 锚点 没有扩充 李娟 11.11.11-->
  <xsl:template match="p:sldLayout">

	  <规则:页面版式_B652>
      <xsl:attribute name="标识符_6B0D">
        <xsl:value-of select="substring-before(@id,'.xml')"/>
      </xsl:attribute>
      <xsl:if test="p:cSld/@name">
        <xsl:attribute name="名称_6BE3">
          <xsl:value-of select="p:cSld/@name"/>
        </xsl:attribute>
      </xsl:if>
			  <演:布局类型_6BE2>
        <!--<xsl:attribute name="演:类型">-->
          <xsl:choose>
            <xsl:when test="@type='title'">title-subtitle</xsl:when>
            <xsl:when test="@type='titleOnly'">title</xsl:when>
            <xsl:when test="@type='obj'">title-media</xsl:when>
            <xsl:when test="@type='twoObj'">title-media-text</xsl:when>
            <xsl:when test="@type='vertTx'">title-text</xsl:when>
            <xsl:when test="@type='vertTitleAndTx'">vtext-vtitle</xsl:when>
            <xsl:when test="@type='picTx'">title-text</xsl:when>
            <xsl:otherwise>title-text</xsl:otherwise>
          </xsl:choose>
        <!--</xsl:attribute>-->
      </演:布局类型_6BE2>
		  <xsl:for-each select="p:cSld/p:spTree/p:sp|p:cSld/p:spTree/p:cxnSp|p:cSld/p:spTree/p:grpSp|p:cSld/p:spTree/p:pic|p:cSld/p:spTree/p:graphicFrame">
			  <xsl:apply-templates select="." mode="ph"/>
		  </xsl:for-each>
	   
			</规则:页面版式_B652>
      </xsl:template>
</xsl:stylesheet>