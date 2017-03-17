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

<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" 
  xmlns:o="urn:schemas-microsoft-com:office:office" 
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" 
  xmlns:w10="urn:schemas-microsoft-com:office:word" 
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  
  <xsl:import href="paragraph.xsl"/>
	<xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
	<xsl:template name="SimpleField">
    <xsl:param name ="splfldPartFrom"/>
    <xsl:param name ="type"/>
	  <字:域开始_419E>
	    <xsl:attribute name="类型_416E">
        <xsl:value-of select ="$type"/>
      </xsl:attribute>
			<xsl:if test="@fldLock='true'">
			  <xsl:attribute name="是否锁定_416F">
          <xsl:value-of select ="'true'"/>
        </xsl:attribute>
			</xsl:if>
      <xsl:if test ="not(@fldLock) or @fldLock='false'">
        <xsl:attribute name="是否锁定_416F">
          <xsl:value-of select ="'false'"/>
        </xsl:attribute>
      </xsl:if>
		</字:域开始_419E>
	  <字:域代码_419F>
	    <字:段落_416B>
	      <字:句_419D>
	        <字:句属性_4158/>
	        <字:文本串_415B>
            <xsl:variable name ="instr" select ="normalize-space(@w:instr | ..//w:instrText | .//w:instrText)"/><!--当前结点为w:fidSimple-->
            <xsl:if test ="($type='revision') or ($type='pageinsection')">
              <xsl:variable name ="instrcode2" select ="substring-after($instr,' ')"/>
              <xsl:variable name ="instrcode0" select ="translate($type,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
              <xsl:variable name ="instrcode1" select ="concat($instrcode0,' ')"/>
              <xsl:value-of select="concat($instrcode1,$instrcode2)"/>
            </xsl:if>
            <xsl:if test ="not(($type='revision') or ($type='pageinsection'))">
              <xsl:for-each select="@w:instr | ..//w:instrText | .//w:instrText">
                <xsl:value-of select="concat(.,'')"/>
              </xsl:for-each>
              <!--<xsl:value-of select="normalize-space(@w:instr | ..//w:instrText | .//w:instrText)"/>-->
            </xsl:if>
          </字:文本串_415B>
				</字:句_419D>
			</字:段落_416B>
		</字:域代码_419F>

    <!--2013-03-27，修复索引转换的BUG，start-->
    <xsl:if test ="$type='xe'">
      <xsl:variable name ="instr1">
        <xsl:for-each select="@w:instr | ..//w:instrText | .//w:instrText">
          <xsl:value-of select="concat(.,'')"/>
        </xsl:for-each>
      </xsl:variable>
      <xsl:variable name ="instr2">
        <xsl:value-of select ="substring-after($instr1,'XE')"/>
      </xsl:variable>
      <xsl:variable name ="instr3">
        <xsl:value-of select ="normalize-space($instr2)"/>
      </xsl:variable>
      <xsl:variable name ="strlgth">
        <xsl:value-of select ="string-length($instr3)"/>
      </xsl:variable>
      <xsl:variable name ="lgth">
        <xsl:number value ="$strlgth - 2"/>
      </xsl:variable>
      <字:句_419D>
        <字:句属性_4158/>
        <字:文本串_415B>
          <xsl:value-of select ="substring($instr3,2,$lgth)"/>
        </字:文本串_415B>
      </字:句_419D>
    </xsl:if>
    <!--end-->
    
		<xsl:for-each select="w:r">
		  <字:句_419D>
        <xsl:call-template name="run">
          <xsl:with-param name ="rPartFrom" select ="$splfldPartFrom"/>
        </xsl:call-template>
		  </字:句_419D>
		</xsl:for-each>
	  
	</xsl:template>

  <xsl:template name="EQfield">
    <xsl:param name ="splfldPartFrom"/>
    <xsl:param name ="type"/>
    <xsl:variable name="startid" select="generate-id(./w:fldChar[@w:fldCharType='begin'])"></xsl:variable>
    <xsl:variable name="nextid" select="generate-id(following-sibling::w:r/w:fldChar[@w:fldCharType='end'])"/>
    
    <字:域开始_419E>
      <xsl:attribute name="类型_416E">
        <xsl:value-of select ="$type"/>
      </xsl:attribute>
      <xsl:if test="@fldLock='true'">
        <xsl:attribute name="是否锁定_416F">
          <xsl:value-of select ="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(@fldLock) or @fldLock='false'">
        <xsl:attribute name="是否锁定_416F">
          <xsl:value-of select ="'false'"/>
        </xsl:attribute>
      </xsl:if>
    </字:域开始_419E>
    <字:域代码_419F>
      <字:段落_416B>
        <xsl:for-each select="following-sibling::node()[(w:fldChar/@w:fldCharType!='end' or not(w:fldChar/@w:fldCharType))
                        and generate-id(following-sibling::w:r/w:fldChar[@w:fldCharType='end']) = $nextid ]">
          <字:句_419D>
            <xsl:for-each select="node()">
              <xsl:if test="name(.)='w:rPr'">
                <字:句属性_4158>
                  <xsl:apply-templates select="." mode="RunProperties"/>
                </字:句属性_4158>              
              </xsl:if>
              <xsl:if test="name(.)='w:instrText'">
                <字:文本串_415B>
                  <xsl:value-of select="."/>
                </字:文本串_415B>
              </xsl:if>
            </xsl:for-each>
          </字:句_419D>
        </xsl:for-each>
        <!--<字:句_419D>
          <字:句属性_4158/>
          <字:文本串_415B>
            --><!--<xsl:variable name ="instr" select ="normalize-space(@w:instr | ..//w:instrText | .//w:instrText)"/>--><!--
            --><!--<xsl:variable name="id" select="generate-id(following-sibling::w:r/w:fldChar[@w:fldCharType='end'])"/>
            <xsl:for-each select="following-sibling::w:r[generate-id(following-sibling::w:r/w:fldChar[@w:fldCharType='end'])='$id']/w:instrText">
              <xsl:value-of select="concat(.,'')"/>
            </xsl:for-each>--><!--
            <xsl:value-of select="normalize-space(..//w:instrText | .//w:instrText)"/>
          </字:文本串_415B>
        </字:句_419D>-->
      </字:段落_416B>
    </字:域代码_419F>
    <xsl:for-each select="w:r">
      <字:句_419D>
        <xsl:call-template name="run">
          <xsl:with-param name ="rPartFrom" select ="$splfldPartFrom"/>
        </xsl:call-template>
      </字:句_419D>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
