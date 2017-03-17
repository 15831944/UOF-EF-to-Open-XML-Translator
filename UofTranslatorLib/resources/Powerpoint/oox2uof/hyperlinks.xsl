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
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks">
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<!--李娟修改2011年11月02-->
  <xsl:template name="hyperlinks">
	  <超链:链接集_AA0B>
  <xsl:for-each select="//a:p">
        <xsl:variable name="pID" select="translate(@id,':','')"/>
        <xsl:for-each select="a:r">
          <xsl:variable name="position">
            <xsl:value-of select="position()"/>
          </xsl:variable>
          <xsl:if test="a:rPr/a:hlinkMouseOver">
            <xsl:for-each select="a:rPr/a:hlinkMouseOver">
              <xsl:call-template name="hyperlink">
                <xsl:with-param name="pID" select="$pID"/>
                <xsl:with-param name="position" select="$position"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test="not(a:rPr/a:hlinkMouseOver)">
            <xsl:for-each select="a:rPr/a:hlinkClick">
              <xsl:call-template name="hyperlink">
                <xsl:with-param name="pID" select="$pID"/>
                <xsl:with-param name="position" select="$position"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
        </xsl:for-each>
      </xsl:for-each>
  </超链:链接集_AA0B>
	 
  </xsl:template>

  <xsl:template name="hyperlink">
    <xsl:param name="pID"/>
    <xsl:param name="position"/>
    <xsl:variable name="sldID" select="ancestor::p:cSld/../@id"/>
    <!--xsl:if test="not(current()/@r:id = preceding::a:hlinkClick[ancestor::p:cSld/../@id=$sldID]/@r:id)"-->
    <!--09.12.22 马有旭 修改-->
    <xsl:if test="not(a:snd)">
	<超链:超级链接_AA0C >
        <xsl:attribute name="标识符_AA0A">
          <xsl:value-of select="concat('hc_',$position,$pID)"/>
        </xsl:attribute>
			
        <xsl:if test="@tooltip">
			<超链:提示_AA05>
            <xsl:value-of select="@tooltip"/>
          </超链:提示_AA05>
        </xsl:if>
        <xsl:variable name="rid" select="@r:id"/>
        <xsl:variable name="target">
          <xsl:for-each select="ancestor::p:cSld/..">
            <xsl:value-of select="following-sibling::rel:Relationships[1]/rel:Relationship[@Id=$rid]/@Target"/>
          </xsl:for-each>
        </xsl:variable>
			<超链:目标_AA01>
          <!--指定任意幻灯片 文件名对应到标识符-->
          <xsl:choose>
            <xsl:when test="@action='ppaction://hlinksldjump'">
              <xsl:value-of select="concat('Slide:','slideId',substring-after(substring-before($target,'.xml'),'slide'))"/>
            </xsl:when>
			<xsl:when test="substring-before(@action,'?')='ppaction://customshow'">
              
              <xsl:choose>
                <xsl:when test="contains(@action,'&amp;')">
                  <xsl:value-of select="concat('Custom Show:',//p:presentation/p:custShowLst/p:custShow[@id=substring-before(substring-after(current()/@action,'?id='),'&amp;')]/@name)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="concat('Custom Show:',//p:presentation/p:custShowLst/p:custShow[@id=substring-after(current()/@action,'?id=')]/@name)"/>
                </xsl:otherwise>
              </xsl:choose>
			</xsl:when>
			  <xsl:when test="substring-before(@action,'?')='ppaction://hlinkshowjump'">
				  <xsl:choose>
					  <xsl:when test="substring-after(@action,'?')='jump=firstslide'">
						  <xsl:value-of select="'First Slide'"/>
					  </xsl:when>
					  <xsl:when test="substring-after(@action,'?')='jump=previousslide'">
						  <xsl:value-of select="'Previous Slide'"/>
					  </xsl:when>
					  <xsl:when test="substring-after(@action,'?')='jump=nextslide'">
						  <xsl:value-of select="'Next Slide'"/>
					  </xsl:when>
					  <xsl:when test="substring-after(@action,'?')='jump=lastslide'">
						  <xsl:value-of select="'Last Slide'"/>
					  </xsl:when>
				  </xsl:choose>
			  </xsl:when>
            
            <xsl:otherwise>
              <xsl:if test="substring-after($target,'file:///')!=''">
                
                <xsl:choose>
                  <xsl:when test="contains($target,'%20')">
                    <xsl:value-of select="translate(substring-after($target,'file:///'),'%20',' ')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="substring-after($target,'file:///')"/>
                  </xsl:otherwise>
                </xsl:choose>

              </xsl:if>
              <xsl:if test="substring-after($target,'file:///')=''">
                <xsl:value-of select="$target"/>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </超链:目标_AA01>
		<超链:链源_AA00>
          <xsl:value-of select="concat('hcS_',$position,$pID)"/>
        </超链:链源_AA00>

        <!--<xsl:if test="contains(@action,'&amp;')">
          <xsl:attribute name="return">
            <xsl:value-of select="substring-after(@action,'return=')"/>
          </xsl:attribute>
        </xsl:if>-->
    </超链:超级链接_AA0C>
    </xsl:if>

  </xsl:template>
</xsl:stylesheet>
