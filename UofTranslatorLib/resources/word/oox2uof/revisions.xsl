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
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"   xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">
	<xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
	<xsl:template name="revisionSet">
		<xsl:if test="//w:del | //w:ins | //w:r/w:rPr/w:rPrChange | //w:tblPrChange | //w:tblGridChange | //w:tcPrChange|//w:sectPrChange|//w:pPrChange|//w:moveFrom|//w:moveTo">
			<xsl:apply-templates select="document('word/document.xml')/w:document/w:body" mode="revisionSet"/>
		</xsl:if>
    
	</xsl:template>
	<xsl:template match="w:body" mode="revisionSet">
		<xsl:if test="//w:del | //w:ins | //w:r/w:rPr/w:rPrChange | //w:tblPrChange | //w:tblGridChange | //w:tcPrChange|//w:sectPrChange|//w:pPrChange|//w:moveFrom|//w:moveTo">
		  <规则:修订信息集_B60E>
				<xsl:for-each select="w:p | w:tbl | w:sectPr">
					<xsl:for-each select="w:del | w:ins | w:r/w:rPr/w:rPrChange|w:pPr/w:pPrChange | w:tblPr/w:tblPrChange | w:tblGrid/w:tblGridChange | w:tr/w:trPr/w:ins | w:tr/w:trPr/w:del | w:tr/w:tc/w:p/w:pPr/w:rPr/w:del | w:tr/w:tc/w:p/w:pPr/w:rPr/w:ins | w:tr/w:tc/w:p/w:del | w:tr/w:tc/w:p/w:ins | w:tr/w:tc/w:p/w:r/w:rPr/w:rPrChange | w:tr/w:tc/w:tcPr/w:tcPrChange|w:sectPrChange|w:moveFrom|w:moveTo">
						<xsl:choose>
							<xsl:when test="name(.)='w:del'">
								<xsl:call-template name="rprchange"/>
								<xsl:if test="./w:r/w:rPr/w:rPrChange">
									<xsl:for-each select="./w:r/w:rPr/w:rPrChange">
										<xsl:call-template name="rprchange"/>
									</xsl:for-each>
								</xsl:if>
							</xsl:when>
							<xsl:when test="name(.)='w:ins'">
								<xsl:call-template name="rprchange"/>
								<xsl:if test="./w:r/w:rPr/w:rPrChange">
									<xsl:for-each select="./w:r/w:rPr/w:rPrChange">
										<xsl:call-template name="rprchange"/>
									</xsl:for-each>
								</xsl:if>
							</xsl:when>
							<xsl:when test="name(.)='w:rPrChange'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
							<xsl:when test="name(.)='w:moveFrom'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
							<xsl:when test="name(.)='w:moveTo'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
              <xsl:when test="name(.)='w:pPrChange'">
                <xsl:call-template name="rprchange"/>
              </xsl:when>
							<xsl:when test="name(.)='w:tblPrChange'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
							<xsl:when test="name(.)='w:tblGridChange'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
							<xsl:when test="name(.)='w:tcPrChange'">
								<xsl:call-template name="rprchange"/>
							</xsl:when>
						</xsl:choose>
					</xsl:for-each>
				</xsl:for-each>
			</规则:修订信息集_B60E>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rprchange">
	  <规则:修订信息_B60F>
			<xsl:variable name="vfmat1">
				<xsl:number count="w:del | w:ins | w:rPrChange | w:tblPrChange | w:tblGridChange | w:tcPrChange | w:sectPrChange | w:pPrChange| w:moveFrom |w:moveTo" level="any" format="1"/>
			</xsl:variable>
      <xsl:variable name ="vfmat2">
        <xsl:value-of select ="$vfmat1+1"/>
      </xsl:variable>
	    <xsl:attribute name="标识符_B610"><xsl:value-of select="concat('rev_',$vfmat1)"/></xsl:attribute>
			<xsl:if test="@w:author">
			  <xsl:attribute name="作者_B611"><xsl:value-of select="concat('aut1_',./@w:author)"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="@w:date">
			  <xsl:attribute name="日期_B612"><xsl:value-of select="substring(./@w:date,1,string-length(./@w:date)-1)"/></xsl:attribute>
			</xsl:if>
		</规则:修订信息_B60F>
	</xsl:template>


	<xsl:template name="pPrchangeBody">
		<xsl:param name ="revPartFrom"/>
		<xsl:param name="lx"/>
		<xsl:variable name="vdel1">
			<xsl:number count="w:ins|w:rPrChange|w:tblPrChange|w:tblGridChange|w:tcPrChange|w:sectPrChange|w:pPrChange|w:moveFrom|w:moveTo" level="any" format="1"/>
		</xsl:variable>
		<xsl:variable name="vdel">
			<xsl:value-of select ="$vdel1 + 1"/>
		</xsl:variable>
		<字:修订开始_421F>
			<!--attrList="标识符 类型 修订信息引用"-->
			<xsl:attribute name="标识符_4220">
				<xsl:value-of select="concat('xd',$vdel1)"/>
			</xsl:attribute>
			<xsl:attribute name="类型_4221">
				<xsl:value-of select="$lx"/>
			</xsl:attribute>
			<xsl:attribute name="修订信息引用_4222">
				<xsl:value-of select="concat('rev_',$vdel1)"/>
			</xsl:attribute>
		</字:修订开始_421F>
		<字:段落属性_419B>
			<xsl:apply-templates select="../../w:pPr"/>
		</字:段落属性_419B>
		<字:修订结束_4223>
			<xsl:attribute name="开始标识引用_4224">
				<xsl:value-of select="concat('xd',$vdel1)"/>
			</xsl:attribute>
		</字:修订结束_4223>
	</xsl:template>
	
	<xsl:template name="rprchangeBody">
    <xsl:param name ="revPartFrom"/>
		<xsl:param name="lx"/>
    <xsl:variable name="vdel1">
      <xsl:number count="w:del|w:ins|w:rPrChange|w:tblPrChange|w:tblGridChange|w:tcPrChange|w:sectPrChange|w:pPrChange|w:moveFrom|w:moveTo" level="any" format="1"/>
    </xsl:variable>
    <xsl:variable name="vdel">
      <xsl:value-of select ="$vdel1 + 1"/>
    </xsl:variable>
		<字:修订开始_421F><!--attrList="标识符 类型 修订信息引用"--> 
			<xsl:attribute name="标识符_4220"><xsl:value-of select="concat('xd',$vdel1)"/></xsl:attribute>
			<xsl:attribute name="类型_4221"><xsl:value-of select="$lx"/></xsl:attribute>
			<xsl:attribute name="修订信息引用_4222"><xsl:value-of select="concat('rev_',$vdel1)"/></xsl:attribute>
		</字:修订开始_421F>
		<!--修订句属性-->
    <xsl:for-each select ="node()">
      <xsl:choose>
        <xsl:when test ="name(.)='w:pPr'">
          <字:段落属性_419B><!--uof:attrList="式样引用"-->
            <xsl:attribute name="式样引用_419C">
              <xsl:apply-templates select="ancestor::w:pPr[1]"/>
            </xsl:attribute>           
          </字:段落属性_419B>
        </xsl:when>
        <xsl:when test ="name(.)='w:r'">

          <!--2013-05-10，wudi，修复拼音指南BUG，通过EQ域实现的拼音指南不转，start-->
          <xsl:if test ="(contains(ancestor::w:ins/w:r/w:instrText,'EQ') or contains(ancestor::w:ins/w:r/w:instrText,'eq')) and ./w:rPr/w:rPrChange">
            <字:句_419D>
              <xsl:call-template name="run">
                <xsl:with-param name ="rPartFrom" select ="$revPartFrom"/>
              </xsl:call-template>
            </字:句_419D>
          </xsl:if>
          <xsl:if test ="not(contains(ancestor::w:ins/w:r/w:instrText,'EQ') or contains(ancestor::w:ins/w:r/w:instrText,'eq'))">
            <字:句_419D>
              <xsl:call-template name="run">
                <xsl:with-param name ="rPartFrom" select ="$revPartFrom"/>
              </xsl:call-template>
            </字:句_419D>
          </xsl:if>
          <!--end-->

        </xsl:when>
        <xsl:when test ="name(.)='w:rPr'">
          <字:句属性_4158>
			  <xsl:apply-templates select="../../../w:rPr" mode="RunProperties"/>
          </字:句属性_4158>
        </xsl:when>
        <!--<xsl:when test="name(.)='w:bookmarkStart'">
          <xsl:call-template name="bookmarkStart"/>
        </xsl:when>
        <xsl:when test="name(.)='w:bookmarkEnd'">
          <xsl:call-template name="bookmarkEnd"/>
        </xsl:when>-->
        <xsl:when test="name(.)='w:commentRangeStart'">
          <xsl:call-template name="commentStart"/>
        </xsl:when>
        <xsl:when test="name(.)='w:commentRangeEnd'">
          <xsl:call-template name="commentEnd"/>
        </xsl:when>
        <xsl:when test="name(.)='w:hyperlink'">
          <xsl:call-template name="hyperlinkRegion">
            <xsl:with-param name ="filename" select ="$pPartFrom"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="name(.)='w:fldSimple'">
          <xsl:variable name="temp" select="normalize-space(@w:instr)"/>
          <xsl:variable name="type" select="substring-before($temp,' ')"/>
          <xsl:choose>
            <xsl:when test="$type='AUTHOR' or $type='FILENAME' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE'or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS'">
              <xsl:call-template name="SimpleField">
                <xsl:with-param name="type" select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>             
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$type='REVNUM'">
              <xsl:call-template name="SimpleField">
                <xsl:with-param name="type" select="'revision'"/>
                <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$type='SECTIONPAGES'">
              <xsl:call-template name="SimpleField">
                <xsl:with-param name="type" select="'pageinsection'"/>
                <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>  
    </xsl:for-each>
		<字:修订结束_4223>
			<xsl:attribute name="开始标识引用_4224"><xsl:value-of select="concat('xd',$vdel1)"/></xsl:attribute>
		</字:修订结束_4223>
	</xsl:template>
 
</xsl:stylesheet>