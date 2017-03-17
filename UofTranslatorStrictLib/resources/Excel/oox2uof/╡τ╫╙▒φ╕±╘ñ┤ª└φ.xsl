<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:uof="http://schemas.uof.org/cn/2003/uof"
                xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow"
                xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc"
                xmlns:图="http://schemas.uof.org/cn/2003/graph" 
                
                xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                               
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"> 
  <!--电子表格预处理，入口为workbook.xml-->
  <xsl:template name="spreadsheets" match="/">
    <ws:spreadsheets>
      <!--sheet name="Sheet10 (6)" sheetId="54" r:id="rId1"/-->
      <xsl:for-each select="ws:workbook/ws:sheets/ws:sheet">
        <xsl:variable name="rId" select="@r:id"/>
        <xsl:variable name="style" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$rId]/@Target"/>

        <xsl:if test="contains($style,'sheet')">
          <xsl:call-template name="sheet">
            <xsl:with-param name="wkbkid" select="@r:id"/>
            <!--wkbkid:rId1-->
            <xsl:with-param name="sId" select="@sheetId"/>
            <!--sId:1-->
            <xsl:with-param name="sName" select="@name"/>
            <!--sName:Sheet1  Sheet10 (6)-->
          </xsl:call-template>
        </xsl:if>

        <xsl:if test="contains($style,'chart')">
          <xsl:call-template name="chart">
            <xsl:with-param name="wkbkid" select="@r:id"/>
            <!--wkbkid:rId1-->
            <xsl:with-param name="sId" select="@sheetId"/>
            <!--sId:1-->
            <xsl:with-param name="sName" select="@name"/>
            <!--sName:Sheet1-->
          </xsl:call-template>
        </xsl:if>

      </xsl:for-each>
      <ws:chartpicture>
        <xsl:copy-of select="document(concat('xlsx/xl/charts/_rels','chart1.xml.rels'))"/>
      </ws:chartpicture>
    </ws:spreadsheets>
  </xsl:template>

  <!--\xlsx\xl\worksheets\sheet10 (6).xml-->


  <xsl:template name="sheet">
    <xsl:param name="wkbkid"/>
    <xsl:param name="sId"/>
    <xsl:param name="sName"/>
    <xsl:variable name="sheetloc" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$wkbkid]/@Target"/>
    <!--sheetloc:worksheets/sheet1.xml-->
    <xsl:variable name="sheetname" select="substring-after($sheetloc,'/')"/>
    <!--sheetname:sheet1.xml-->
    <xsl:variable name="sheetrels" select="concat($sheetname,'.rels')"/>
    <!--sheetrels:sheet1.xml.rels-->
    <ws:spreadsheet>
      <ws:pos>
        <xsl:value-of select="$sId"/>
      </ws:pos>
      <ws:sname>
        <xsl:value-of select="$sName"/>
      </ws:sname>

      <xsl:choose>
        <xsl:when test="contains($sheetloc,'worksheets')">
          <xsl:copy-of select="document(concat('xlsx/xl/',$sheetloc))/ws:worksheet"/>
          <xsl:if test="document(concat('xlsx/xl/',$sheetloc))/ws:worksheet/*[@r:id]">
            <xsl:copy-of select="document(concat('xlsx/xl/worksheets/_rels/',$sheetrels))/pr:Relationships"/>
          </xsl:if>
        </xsl:when>
      </xsl:choose>

      <xsl:variable name="pos">
        <xsl:value-of select="translate($wkbkid,'rId','')"/>
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

      <xsl:if test="document(concat('xlsx/xl/',$sheetloc))/ws:worksheet/*[@r:id]">
        <xsl:if test="document(concat('xlsx/xl/worksheets/_rels/',$sheetrels))/pr:Relationships/pr:Relationship[@Target=$d]">
          <ws:Drawings>
            <xsl:for-each select="document(concat('xlsx/xl/drawings/',$dd))/xdr:wsDr">
              <xsl:if test="descendant::xdr:pic or descendant::xdr:graphicFrame">
                <xsl:copy-of select="document(concat('xlsx/xl/drawings/_rels/',$ddd))"/>
              </xsl:if>
            </xsl:for-each>
            <xsl:copy-of select="document(concat('xlsx/xl/drawings/',$dd))"/>
          </ws:Drawings>
        </xsl:if>
      </xsl:if>

    </ws:spreadsheet>
  </xsl:template>


  <xsl:template name="chart">
    <xsl:param name="wkbkid"/>
    <xsl:param name="sId"/>
    <xsl:param name="sName"/>
    <xsl:variable name="sheetloc" select="document('xlsx/xl/_rels/workbook.xml.rels')/pr:Relationships/pr:Relationship[@Id=$wkbkid]/@Target"/>
    <xsl:variable name="sheetname" select="substring-after($sheetloc,'/')"/>
    <xsl:variable name="sheetrels" select="concat($sheetname,'.rels')"/>

    <ws:spreadsheet>
      <ws:pos>
        <xsl:value-of select="$sId"/>
      </ws:pos>
      <ws:sname>
        <xsl:value-of select="$sName"/>
      </ws:sname>
      <a>
        <xsl:value-of select ="$sheetloc"/>
      </a>
      <xsl:choose>
        <xsl:when test="contains($sheetloc,'chartsheets')">
          <!--<xsl:copy-of select="document(concat('xlsx/xl/',$sheetloc))/ws:chartsheet"/>-->
          <xsl:copy-of select="document(concat('xlsx/xl/',$sheetloc))/ws:chartsheet"/>
          <xsl:copy-of select="document(concat('xlsx/xl/chartsheets/_rels/',$sheetrels))/pr:Relationships"/>
        </xsl:when>
      </xsl:choose>


      <xsl:variable name="pos">
        <xsl:value-of select="translate($wkbkid,'rId','')"/>
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
          <xsl:copy-of select="document(concat('xlsx/xl/drawings/_rels/',$ddd))"/>
          <xsl:copy-of select="document(concat('xlsx/xl/drawings/',$dd))"/>
          <ll>lllllllllll</ll>
        </ws:Drawings>

        <!--3月17huai？？？？？？？？？？？？？？-->
        <!--ws:chartpicture>
                <xsl:copy-of select="document(concat('xlsx/xl/charts/_rels','chart1.xml.rels'))"/>			
			</ws:chartpicture-->

      </xsl:if>



    </ws:spreadsheet>
  </xsl:template>









</xsl:stylesheet>
