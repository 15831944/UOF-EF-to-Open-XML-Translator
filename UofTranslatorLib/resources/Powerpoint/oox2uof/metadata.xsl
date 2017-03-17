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
xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:元="http://schemas.uof.org/cn/2009/metadata">
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<xsl:template name="metaData">
		<!--<uof:元数据 uof:locID="u0001">
			<xsl:if test="document('_rels/.rels')/rel:Relationships/rel:Relationship[@Target='docProps/app.xml']">
				<xsl:apply-templates select="document('docProps/app.xml')/app:Properties"/>
			</xsl:if>
			<xsl:if test="document('_rels/.rels')/rel:Relationships/rel:Relationship[@Target='docProps/core.xml']">
				<xsl:apply-templates select="document('docProps/core.xml')/cp:coreProperties"/>
			</xsl:if>
			<xsl:if test="document('_rels/.rels')/rel:Relationships/rel:Relationship[@Target='docProps/custom.xml']">
				<xsl:apply-templates select="document('docProps/custom.xml')/cus:Properties"/>
			</xsl:if>
		</uof:元数据>-->
	<元:元数据_5200>
      <xsl:if test="/p:presentation/app:Properties">
        <xsl:apply-templates select="/p:presentation/app:Properties"/>
      </xsl:if>
      <xsl:if test="/p:presentation/cp:coreProperties">
        <xsl:apply-templates select="/p:presentation/cp:coreProperties"/>
      </xsl:if>
      <xsl:if test="/p:presentation/cus:Properties">
        <xsl:apply-templates select="/p:presentation/cus:Properties"/>
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
		<xsl:if test="cp:category">
			<元:分类_520B>
				<xsl:value-of select="./cp:category"/>
			</元:分类_520B>
		</xsl:if>

    <!--2014-01-14, tangjiang, 修复Strict OOXML到UOF PPT关键字的转换 start -->
    <xsl:if test="cp:keywords">
      <元:关键字集_520D>
        <xsl:for-each select="./cp:keywords">
          <元:关键字_520E>
            <xsl:value-of select="."/>
          </元:关键字_520E>
        </xsl:for-each>
      </元:关键字集_520D>
    </xsl:if>
    <!--end-->

  </xsl:template>
  <!--2013-1-12 修改自定义元数据丢失 liqiuling start -->
	<xsl:template match="cus:Properties">
		<元:用户自定义元数据集_520F>
			<xsl:for-each select="./cus:property">
				<元:用户自定义元数据_5210>
          <xsl:attribute name="名称_5211">
            <xsl:value-of select="./@name"/>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="name(./vt:lpwstr) = 'vt:lpwstr'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'string'"/>
              </xsl:attribute>
              <xsl:value-of select="./vt:lpwstr"/>
            </xsl:when>
            <xsl:when test="name(./vt:filetime) = 'vt:filetime'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'datetime'"/>
              </xsl:attribute>
              <xsl:value-of select="./vt:filetime"/>
            </xsl:when>
            <xsl:when test="name(./vt:i4) = 'vt:i4'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'float'"/>
              </xsl:attribute>
              <xsl:value-of select="./vt:i4"/>
            </xsl:when>
            <xsl:when test="name(./vt:r8) = 'vt:r8'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'float'"/>
              </xsl:attribute>
              <xsl:value-of select="./vt:r8"/>
            </xsl:when>
            <xsl:when test="name(./vt:bool) = 'vt:bool'">
              <xsl:attribute name="类型_5212">
                <xsl:value-of select="'boolean'"/>
              </xsl:attribute>
              <xsl:variable name="boolVal" select="./vt:bool"/>
              <xsl:if test="$boolVal='true'">
                <xsl:value-of select="'是'"/>
              </xsl:if>
              <xsl:if test="$boolVal='false'">
                <xsl:value-of select="'否'"/>
              </xsl:if>
            </xsl:when>
          </xsl:choose>					
				</元:用户自定义元数据_5210>
			</xsl:for-each>
		</元:用户自定义元数据集_520F>
    <!--2013-1-12 修改自定义元数据丢失 liqiuling start                                                     -->
	</xsl:template>
</xsl:stylesheet>
