<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties" 
                xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:cus="http://purl.oclc.org/ooxml/officeDocument/customProperties"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
                xmlns:ori="http://purl.oclc.org/ooxml/officeDocument/relationships/image"
                xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
                
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:pc="http://schemas.openxmlformats.org/package/2006/content-types"
                xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
  <!--电子表格预处理，入口为xslx/xl/workbook.xml-->
  <!--预处理，整合所有单个子文件于一个XML文件中-->
  <xsl:template name="spreadsheets" match="/">
    <ws:spreadsheets>
      <xsl:variable name="rels" select="document('xlsx/_rels/.rels')"/>
      <xsl:variable name="workboolRels" select="document('xlsx/xl/_rels/workbook.xml.rels')"/>
      <!--workbook.xml.rels-->
      <workbook.xml.rels>
        <xsl:copy-of select="document('xlsx/xl/_rels/workbook.xml.rels')"/>
      </workbook.xml.rels>
      <!--[Content_Types].xml-->
      <xsl:copy-of select="document('xlsx/[Content_Types].xml')"/>
      <!--theme1.xml-->
      <xsl:copy-of select="document('xlsx/xl/theme/theme1.xml')"/>
      <!--workbook.xml-->
      <xsl:copy-of select="document('xlsx/xl/workbook.xml')"/>
      <!--sharedStrings.xml-->
      <xsl:if test="$workboolRels/pr:Relationships/pr:Relationship[@Target='sharedStrings.xml']">
        <xsl:copy-of select="document('xlsx/xl/sharedStrings.xml')"/>
      </xsl:if>
      <!--styles.xml-->
      <xsl:copy-of select="document('xlsx/xl/styles.xml')"/>
      <!--app.xml-->
      <xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/app.xml']">
        <xsl:copy-of select="document('xlsx/docProps/app.xml')"/>
      </xsl:if>
      <!--core.xml-->
      <xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/core.xml']">
        <xsl:copy-of select="document('xlsx/docProps/core.xml')"/>
      </xsl:if>
      <!--custom.xml-->
      <xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/custom.xml']">
        <xsl:copy-of select="document('xlsx/docProps/custom.xml')"/>
      </xsl:if> 
      <!--sheet-->
      <!--入口为xslx/xl/workbool.xml-->
      <xsl:for-each select="ws:workbook/ws:sheets/ws:sheet">
        <xsl:variable name="sheetSeq" select="position()"/>
        <xsl:variable name="rId" select="@r:id"/>
        <xsl:variable name="target" select="$workboolRels/pr:Relationships/pr:Relationship[@Id=$rId]/@Target"/>
        <xsl:choose>
          <!--chartsheet-->
          <xsl:when test="contains($target,'chartsheets')">
            <xsl:call-template name="chartsheet">
              <xsl:with-param name="workbookId" select="@r:id"/>
              <xsl:with-param name="sId" select="@sheetId"/>
              <xsl:with-param name="sName" select="@name"/>
            </xsl:call-template>
          </xsl:when>
          
          <!--sheet-->
          <!--嵌入表格中的图表-->
          <xsl:when test="contains($target,'sheet')">
            <xsl:call-template name="sheet">
              <xsl:with-param name="workbookId" select="@r:id"/>
              <!--workbookId:rId1-->
              <xsl:with-param name="sheetId" select="@sheetId"/>
              <!--sId:1-->
              <xsl:with-param name="sheetName" select="@name"/>
            </xsl:call-template>
          </xsl:when>
          <!--chart-->
          <!--
          <xsl:when test="contains($style,'chart')">
            <xsl:call-template name="chart">
              <xsl:with-param name="workbookId" select="@r:id"/>
              <xsl:with-param name="sId" select="@sheetId"/>
              <xsl:with-param name="sName" select="@name"/>
            </xsl:call-template>
          </xsl:when>
          -->
        </xsl:choose>
      </xsl:for-each>
    </ws:spreadsheets>
  </xsl:template>

  <!--sheet-->
  <xsl:template name="sheet">
    <xsl:param name="workbookId"/>
    <xsl:param name="sheetId"/>
    <xsl:param name="sheetName"/>
    <!--形如worksheets/sheet1.xml-->
    <xsl:variable name="sheetLoc" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$workbookId]/@Target"/>
    <!--形如sheet1.xml-->
    <xsl:variable name="fileName" select="substring-after($sheetLoc,'/')"/>
    <!--sheetrels:sheet1.xml.rels-->
    <xsl:variable name="sheetRels" select="concat($fileName,'.rels')"/>
    <ws:spreadsheet>
      <!--sheet1.xml.rels中的内容已在c#部分添加到sheet1.xml中了-->
      <xsl:variable name="dawingLocTemp" select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/r:Relationships/r:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/drawing']/@Target"/>
      <xsl:variable name="dawingLoc" select="substring-after($dawingLocTemp,'../drawings/')"/>
      <xsl:attribute name="sheetId">
        <xsl:value-of select="$sheetId"/>
      </xsl:attribute>
      <xsl:attribute name="sheetName">
        <xsl:value-of select="$sheetName"/>
      </xsl:attribute>
      
      <xsl:copy-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet"/>
      
      <xsl:choose>
        <xsl:when test="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/descendant-or-self::*[@r:id]">
          <xsl:copy-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/pr:Relationships"/>
        </xsl:when>
        <xsl:otherwise>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
        </xsl:otherwise>
      </xsl:choose>
      
      <xsl:if test="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/pr:Relationships/pr:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/drawing']">
        <ws:Drawings>
          <xsl:variable name="drawingTarget">
            <xsl:value-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/pr:Relationships/pr:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/drawing']/@Target"/>
          </xsl:variable>
          <xsl:copy-of select="document(concat('xlsx/xl/drawings/',substring-after($drawingTarget,'../drawings/')))"/>
        </ws:Drawings>
      </xsl:if>

      <!--comments-->
      <xsl:for-each select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/pr:Relationships/pr:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/comments']">
        <xsl:copy-of select="document(concat('xlsx/xl/',substring-after(@Target,'/')))"/>
      </xsl:for-each>
      <!-- 20130327 add by xuzhenwei 批注位置修正 start -->
      <!-- comments vmlDrawing1.xml -->
      <xsl:for-each select="document(concat('xlsx/xl/',$sheetLoc))/ws:worksheet/pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing']">
        <xsl:copy-of select="document(concat('xlsx/xl/drawings/',substring-after(@Target,'../drawings/')))"/>
      </xsl:for-each>
    </ws:spreadsheet>
  </xsl:template>
  
  <!--chart-->
  <xsl:template name="chart">
    <xsl:param name="workbookId"/>
    <xsl:param name="sId"/>
    <xsl:param name="sName"/>
    <xsl:variable name="sheetLoc" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$workbookId]/@Target"/>
    <xsl:variable name="sheetname" select="substring-after($sheetLoc,'/')"/>
    <xsl:variable name="sheetrels" select="concat($sheetname,'.rels')"/>
    <ws:spreadsheet>
      <ws:pos>
        <xsl:value-of select="$sId"/>
      </ws:pos>
      <ws:sname>
        <xsl:value-of select="$sName"/>
      </ws:sname>
      <xsl:choose>
        <xsl:when test="contains($sheetLoc,'chartsheets')">
          <xsl:copy-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:chartsheet"/>
          <xsl:copy-of select="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships"/>
        </xsl:when>
      </xsl:choose>
      <xsl:variable name="pos">
        <xsl:value-of select="translate($workbookId,'rId','')"/>
      </xsl:variable>
      <xsl:variable name="d">
        <xsl:value-of select="concat('../drawings/drawing',$pos,'.xml')"/>
      </xsl:variable>
      <xsl:variable name="dd">
        <xsl:value-of select="concat('drawing',$pos,'.xml')"/>
      </xsl:variable>
      <xsl:variable name="ddd">
        <xsl:value-of select="concat('drawing',$pos,'.xml.rels')"/>
      </xsl:variable>
      <xsl:if test="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships/pr:Relationship[@Target=$d]">
        <ws:Drawings>
          <xsl:copy-of select="document(concat('xlsx/xl/drawings/',$dd))"/>
        </ws:Drawings>
      </xsl:if>
    </ws:spreadsheet>
  </xsl:template>
  
  <xsl:template name="chartsheet">
    <xsl:param name="workbookId"/>
    <xsl:param name="sId"/>
    <xsl:param name="sName"/>
    <xsl:variable name="sheetLoc" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$workbookId]/@Target"/>
    <xsl:variable name="sheetname" select="substring-after($sheetLoc,'/')"/>
    <xsl:variable name="sheetrels" select="concat($sheetname,'.rels')"/>
    <ws:spreadsheet>
      <xsl:attribute name="sheetId">
        <xsl:value-of select="$sId"/>
      </xsl:attribute>
      <xsl:attribute name="sheetName">
        <xsl:value-of select="$sName"/>
      </xsl:attribute>
      
      <xsl:copy-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:chartsheet"/>
      <xsl:copy-of select="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships"/>
      <xsl:variable name="pos">
        <xsl:value-of select="document(concat('xlsx/xl/',$sheetLoc))/ws:chartsheet/ws:drawing/@r:id"/>
      </xsl:variable>
      <xsl:if test="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships/pr:Relationship[@Id=$pos]">
        <xsl:variable name="goal" select="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships/pr:Relationship/@Target"/>
        <xsl:variable name="drawingname" select="substring-after($goal,'../drawings/')"/>
        <!--<xsl:value-of select="$drawingname"/>-->
        <ws:Drawings>
          <xsl:copy-of select="document(concat('xlsx/xl/drawings/',$drawingname))"/>
        </ws:Drawings>
      </xsl:if>
    </ws:spreadsheet>
  </xsl:template>
</xsl:stylesheet>
