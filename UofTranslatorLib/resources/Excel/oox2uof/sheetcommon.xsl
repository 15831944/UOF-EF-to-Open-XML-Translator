<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	<!--根节点匹配workbook.xml-->
	<xsl:template name="common" match="/">
	
		<!--表:公用处理规则 uof:locID="s0000"-->
		<表:度量单位 uof:locID="s0001">pt</表:度量单位>
		<xsl:if test=".//ws:workbookPr[@date1904 and @date1904 = '1' ]">
			<表:日期系统-1904 uof:locID="s0003" uof:attrList="值" 表:值="true"/>
		</xsl:if>
    <xsl:value-of select="name()"/>
    <xsl:if test=".//ws:workbookView[@showHorizontalScroll = '0']">
      <表:是否显示水平滚动条 uof:locID="s0133" uof:attrList="值" 表:值="false"/>
    </xsl:if>

    <xsl:if test=".//ws:workbookView[@showVerticalScroll = '0']">
      <表:是否显示垂直滚动条 uof:locID="s0134" uof:attrList="值" 表:值="false"/>
    </xsl:if>
		<!--计算设置-->
		<xsl:apply-templates select="ws:workbook/ws:calcPr"/>
		
		<!--数据有效性-->
		<!--OOXML中没有‘是否RC引用’，但UOF中是必须的，位置不科学！-->
		<表:是否RC引用 uof:locID="s0124" uof:attrList="值" 表:值="false"/>
		<!--/表:公用处理规则-->
	</xsl:template>
	<xsl:template match="ws:calcPr">
		
		<xsl:if test="@iterate = 1 or @iterate = 'true'">
			<表:计算设置 uof:locID="s0004" uof:attrList="迭代次数 偏差值">
				<xsl:attribute name="表:迭代次数"><xsl:if test="@iterateCount"><xsl:value-of select="@iterateCount"/></xsl:if><xsl:if test="not(@iterateCount)"><xsl:value-of select="100"/></xsl:if></xsl:attribute>
				<xsl:attribute name="表:偏差值"><xsl:if test="@iterateDelta"><xsl:value-of select="@iterateDelta"/></xsl:if><xsl:if test="not(@iterateDelta)"><xsl:value-of select="0.001"/></xsl:if></xsl:attribute>
			</表:计算设置>
		</xsl:if>
		<xsl:if test="@fullPrecision">
			
			<表:精确度以显示值为准 uof:locID="s0002" uof:attrList="值" 表:值="true"/>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
