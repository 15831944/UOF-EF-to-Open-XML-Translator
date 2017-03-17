<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <xsl:template name="LineTypeMapping_Border">
    <xsl:param name ="lineType"/>
    <xsl:choose>
      <xsl:when test="$lineType = 'single'">
        <xsl:value-of select="'sng'"/>
      </xsl:when>
      <xsl:when test="$lineType = 'double'">
        <xsl:value-of select="'dbl'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--Modified by LDM in 2010/12/29-->
  <!--线形模板-->
  <!--LineTypeMapping_Border-->
  <xsl:template name="LineDashMapping_Border">
    <xsl:param name="dashType"/>
    <xsl:choose>
      <!--单线-->
      <xsl:when test ="not($dashType = 'none')">
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:choose>
              <!--实线-->
              <xsl:when test="$dashType = 'noDash' or $dashType = 'solid'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <!--虚线-->
              <xsl:when test="$dashType = 'square-dot'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <!--划线-->
              <xsl:when test="$dashType = 'dash'">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <!--长划线-->
              <xsl:when test="$dashType = 'long-dash'">
                <xsl:value-of select="'lgDash'"/>
              </xsl:when>
              <!--点划线-->
              <xsl:when test="$dashType = 'dash-dot'">
                <xsl:value-of select="'lgDashDot'"/>
              </xsl:when>
              <!--双点划线-->
              <xsl:when test="$dashType = 'dash-dot-dot'">
                <xsl:value-of select="'lgDashDotDot'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </a:prstDash>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
