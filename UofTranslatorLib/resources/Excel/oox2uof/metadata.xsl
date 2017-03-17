<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:fo="http://www.w3.org/1999/XSL/Format" 
                xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" 
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                xmlns:dc="http://purl.org/dc/elements/1.1/" 
                xmlns:dcterms="http://purl.org/dc/terms/" 
                xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
                xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" 
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" 
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:元="http://schemas.uof.org/cn/2009/metadata"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

	<xsl:template name="metadata">
    <元:元数据_5200>
      <xsl:if test="/ws:spreadsheets/app:Properties">
        <xsl:apply-templates select="/ws:spreadsheets/app:Properties"/>
      </xsl:if>
      <xsl:if test="/ws:spreadsheets/cp:coreProperties">
        <xsl:apply-templates select="/ws:spreadsheets/cp:coreProperties"/>
      </xsl:if>
      <xsl:if test="/ws:spreadsheets/cus:Properties">
        <xsl:apply-templates select="/ws:spreadsheets/cus:Properties"/>
      </xsl:if>
    </元:元数据_5200>
	</xsl:template>
	
	<xsl:template match="app:Properties">
		<xsl:if test="app:TotalTime">
      <xsl:variable name ="Mintime">
        <xsl:value-of select="./app:TotalTime"/>
      </xsl:variable>
      <xsl:variable name ="pretime">
        <xsl:value-of select ="concat('P0Y0M0DT0H',$Mintime)"/>
      </xsl:variable>
      <xsl:variable name ="totaltime">
        <xsl:value-of select ="concat($pretime,'M0S')"/>
      </xsl:variable>
      <元:编辑时间_5209>
        <xsl:value-of select="$totaltime"/>
      </元:编辑时间_5209>
		</xsl:if>
		<xsl:if test="app:Application">
      <元:创建应用程序_520A>
        <xsl:value-of select="./app:Application"/>
      </元:创建应用程序_520A>
		</xsl:if>
		<xsl:if test="app:Template">
      <元:文档模板_520C>
        <xsl:value-of select="./app:Template"/>
      </元:文档模板_520C>
		</xsl:if>
		<xsl:if test="app:Company">
      <元:公司名称_5213>
        <xsl:value-of select="./app:Company"/>
      </元:公司名称_5213>
		</xsl:if>
		<xsl:if test="app:Manager">
      <元:经理名称_5214>
        <xsl:value-of select="./app:Manager"/>
      </元:经理名称_5214>
		</xsl:if>
		<xsl:if test="app:Pages">
      <元:页数_5215>
        <xsl:value-of select="./app:Pages"/>
      </元:页数_5215>
		</xsl:if>
		<xsl:if test="app:Words">
      <元:字数_5216>
        <xsl:value-of select="./app:Words"/>
      </元:字数_5216>
		</xsl:if>
		<xsl:if test="app:Lines">
      <元:行数_5219>
        <xsl:value-of select="./app:Lines"/>
      </元:行数_5219>
		</xsl:if>
		<xsl:if test="app:Paragraphs">
      <元:段落数_521A>
        <xsl:value-of select="./app:Paragraphs"/>
      </元:段落数_521A>
		</xsl:if>
	</xsl:template>
	<xsl:template match="cp:coreProperties">
		<xsl:if test="dc:title">
      <元:标题_5201>
        <xsl:value-of select="./dc:title"/>
      </元:标题_5201>
		</xsl:if>
		<xsl:if test="dc:subject">
      <元:主题_5202>
        <xsl:value-of select="./dc:subject"/>
      </元:主题_5202>
		</xsl:if>
		<xsl:if test="dc:creator">
      <元:创建者_5203>
        <xsl:value-of select="./dc:creator"/>
      </元:创建者_5203>
      <元:作者_5204>
        <xsl:value-of select="./dc:creator"/>
      </元:作者_5204>
		</xsl:if>
		<xsl:if test="cp:lastModifiedBy">
      <元:最后作者_5205>
        <xsl:value-of select="./cp:lastModifiedBy"/>
      </元:最后作者_5205>
		</xsl:if>
		<xsl:if test="dc:description">
      <元:摘要_5206>
        <xsl:value-of select="./dc:description"/>
      </元:摘要_5206>
		</xsl:if>
		<xsl:if test="dcterms:created">
      <元:创建日期_5207>
        <xsl:value-of select="substring(dcterms:created,1,string-length(dcterms:created)-1)"/>
      </元:创建日期_5207>
		</xsl:if>
		<xsl:if test="cp:revision">
      <元:编辑次数_5208>
        <xsl:value-of select="./cp:revision"/>
      </元:编辑次数_5208>
		</xsl:if>
    <xsl:if test ="not(cp:revision) and /ws:spreadsheets/workbook/fileVersion/@lastEdited">
      <元:编辑次数_5208>
        <xsl:value-of select ="/ws:spreadsheets/workbook/fileVersion/@lastEdited"/>
      </元:编辑次数_5208>
    </xsl:if>
		<xsl:if test="cp:category">
      <元:分类_520B>
        <xsl:value-of select="./cp:category"/>
      </元:分类_520B>
		</xsl:if>
		<xsl:if test="cp:keywords">
      <元:关键字集_520D>
        <xsl:for-each select="./cp:keywords">
          <元:关键字_520E>
            <xsl:value-of select="."/>
          </元:关键字_520E>
        </xsl:for-each>
      </元:关键字集_520D>
		</xsl:if>
	</xsl:template>
	<xsl:template match="cus:Properties">
    <元:用户自定义元数据集_520F>
      <xsl:for-each select="./cus:property">
        <元:用户自定义元数据_5210>
          <xsl:attribute name="名称_5211">
            <xsl:value-of select="./@name"/>
          </xsl:attribute>
          
          <!--2014-3-11 update Qihy 修复BUG2937-自定义属性值丢失 start-->
          <xsl:choose>
            <xsl:when test="name(./child::*[position() = 1]) = 'vt:lpwstr'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'string'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="name(./child::*[position() = 1]) = 'vt:filetime'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'datetime'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="name(./child::*[position() = 1]) = 'vt:i4'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'float'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="name(./child::*[position() = 1]) = 'vt:bool'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'boolean'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
          <xsl:value-of select="./child::*[position() = 1]"/>
          <!--2014-3-11 end-->
          
        </元:用户自定义元数据_5210>
      </xsl:for-each>
    </元:用户自定义元数据集_520F>
	</xsl:template>
</xsl:stylesheet>
