<?xml version="1.0" encoding="UTF-8"?>
<!--
* Copyright (c) 2006, Beihang University, China
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of Clever Age, nor the names of its contributors may
*       be used to endorse or promote products derived from this software
*       without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
-->
<!--
<Author>Fang Chunyan(BITI)</Author>
-->
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
  <xsl:template match="/">
		<w:document>
      <xsl:if test ="w:document/w:background">
        <xsl:copy-of select ="w:document/w:background"/>
      </xsl:if>
			<xsl:apply-templates select="w:document/w:body" mode="pretreat2"/>
		</w:document>
	</xsl:template>
	<xsl:template match="w:body" mode="pretreat2">
		<w:body>
			<xsl:for-each select="w:sectPr">
				<xsl:copy-of select="."/>
				<xsl:variable name="id1" select="generate-id(.)"/>
				<xsl:variable name="id" select="generate-id(preceding-sibling::w:sectPr)"/>
				<xsl:apply-templates select="preceding-sibling::node()[generate-id(following-sibling::w:sectPr) = $id1 and generate-id(.) != $id and generate-id(.) != $id1 and name(.)!='w:sectPr']" mode="pretreatment2"/>
			</xsl:for-each>
		</w:body>
	</xsl:template>
	<xsl:template match="node()" mode="pretreatment2">
		<xsl:copy-of select="."/>
	</xsl:template>
</xsl:stylesheet>
