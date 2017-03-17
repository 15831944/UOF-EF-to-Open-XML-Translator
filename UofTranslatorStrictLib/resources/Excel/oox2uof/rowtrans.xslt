<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="row26" select="'ABC'"/>

	
	
	<!--xsl:template match="/">

		</xsl:variable>
		<xsl:call-template name="Ts2Dec">
			<xsl:with-param name="tsSource" select="$RowID">
			</xsl:with-param>
		</xsl:call-template>
		<xsl:variable name="ss">
		<xsl:call-template name="Ts2Dec">
			<xsl:with-param name="tsSource" select="$RowID">
			</xsl:with-param>
		</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select="$ss"/>
	</xsl:template-->
	
	
	<xsl:template name="Ts2Dec">
		<xsl:param name="tsSource"/>
		
		<xsl:variable name="z" select="'000000'"/>
		<xsl:variable name="RowID" select="concat(substring($z,1,(string-length($z)-string-length($tsSource))),$tsSource)"/>
		<xsl:value-of select="$tsSource"/>
		<xsl:text>======</xsl:text>
		<xsl:variable name="Char26" select="'0ABCDEFGHIJKLMNOPQRSTUVWXYZ'"/>
		<xsl:variable name="dec1" select="string-length(substring-before($Char26,substring($RowID,1,1)))">
		</xsl:variable>
		<xsl:variable name="dec2" select="$dec1*26+string-length(substring-before($Char26,substring($RowID,2,1)))">
		</xsl:variable>
		<xsl:variable name="dec3" select="$dec2*26+string-length(substring-before($Char26,substring($RowID,3,1)))">
		</xsl:variable>
		<xsl:variable name="dec4" select="$dec3*26+string-length(substring-before($Char26,substring($RowID,4,1)))">
		</xsl:variable>
		<xsl:variable name="dec5" select="$dec4*26+string-length(substring-before($Char26,substring($RowID,5,1)))">
		</xsl:variable>
		<xsl:variable name="Result" select="$dec5*26+string-length(substring-before($Char26,substring($RowID,6,1)))">
		</xsl:variable>
		<xsl:value-of select="$Result"/>
		
	</xsl:template>
	
	
	
</xsl:stylesheet>
