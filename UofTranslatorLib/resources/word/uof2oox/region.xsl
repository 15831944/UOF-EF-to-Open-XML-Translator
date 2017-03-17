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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template name="paragraphRevision">
		<xsl:param name="revType"/>
		<xsl:variable name="revId">
			<xsl:value-of select="@修订信息引用_4222"/>
		</xsl:variable>
		<xsl:variable name="revDate">
			<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@日期_B612"/>
		</xsl:variable>
		<xsl:variable name="userId">
			<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@作者_B611"/>
		</xsl:variable>
		<xsl:variable name="username">
			<xsl:value-of select="preceding::规则:用户_B668[@标识符_4100=$userId]/@姓名_41DC"/>
		</xsl:variable>
		<xsl:element name="{$revType}">
			<xsl:attribute name="w:id">
				<xsl:apply-templates select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]" mode="rev"/>
			</xsl:attribute>
			<xsl:attribute name="w:author">
				<xsl:value-of select="$username"/>
			</xsl:attribute>
			<xsl:attribute name="w:date">
				<xsl:value-of select="concat($revDate,'Z')"/>
			</xsl:attribute>
      <xsl:choose>
        <xsl:when test="$revType='w:pPrChange'">
          <xsl:apply-templates select="字:段落属性_419B" mode="revpPr"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="字:句_419D">

            <!--2013-04-23，wudi，修复修订转换的BUG，修订内容为空格符，修订类型为"delete"的情况，start-->
            <xsl:call-template name="run">
              <xsl:with-param name ="revType" select ="$revType"/>
            </xsl:call-template>
            <!--end-->

          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template match="字:段落属性_419B" mode="revpPr">
		<xsl:call-template name="pPrWithpStyle"/>
	</xsl:template>
	<xsl:template match="规则:修订信息_B60F" mode="rev">
		<xsl:variable name="num">
			<xsl:number count="规则:修订信息_B60F" format="1" level="single"/>
		</xsl:variable>
		<xsl:value-of select="$num"/>
	</xsl:template>

	<xsl:template name="runRevision">
		<xsl:param name="revType"/>
		<xsl:variable name="revId">
			<xsl:value-of select="@修订信息引用_4222"/>
		</xsl:variable>
		<xsl:variable name="revDate">
			<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@日期_B612"/>
		</xsl:variable>
		<xsl:variable name="userId">
			<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@作者_B611"/>
		</xsl:variable>
		<xsl:variable name="username">
			<xsl:value-of select="preceding::规则:用户_B668[@标识符_4100=$userId]/@姓名_41DC"/>
		</xsl:variable>
		<xsl:element name="{$revType}">
			<xsl:attribute name="w:id">
				<xsl:apply-templates select="preceding::规则:修订信息_B60F[@字:标识符=$revId]" mode="rev"/>
			</xsl:attribute>
			<xsl:attribute name="w:author">
				<xsl:value-of select="$username"/>
			</xsl:attribute>
			<xsl:attribute name="w:date">
				<xsl:value-of select="concat($revDate,'Z')"/>
			</xsl:attribute>
			<xsl:if test="$revType='w:rPrChange'">
				<xsl:apply-templates select="字:句属性_4158" mode="rpr"/>
			</xsl:if>
		</xsl:element>
	</xsl:template>

	<xsl:template name="bkStart">
		<w:bookmarkStart>
			<xsl:variable name="id">
        <xsl:number count="//字:区域开始_4165[@类型_413B = 'bookmark']" format="1" level="any"/>
				<!--<xsl:value-of select="./@标识符_4100"/>-->
			</xsl:variable>
      
      <!--2012-12-28，wudi，UOF到OOX方向的书签实现，修改w:id的取值，原取值$id，start-->
			<xsl:attribute name="w:id">
        <xsl:value-of select="$id - 1"/>
				<!--<xsl:apply-templates select="preceding::书签:书签_9105[书签:区域_9100/@区域引用_41CE=$id]" mode="bkid"/>-->
			</xsl:attribute>
      <!--end-->
      
			<xsl:attribute name="w:name">
        <xsl:choose>
          <xsl:when test="@名称_4166">
            <xsl:value-of select="@名称_4166"/>
          </xsl:when>
          <xsl:when test="not(@名称_4166)">
            <xsl:value-of select="preceding::书签:书签_9105[书签:区域_9100/@区域引用_41CE=$id]/@名称_4166"/>
          </xsl:when>
        </xsl:choose>
			</xsl:attribute>
		</w:bookmarkStart>
	</xsl:template>
	<xsl:template match="书签:书签_9105" mode="bkid">
		<xsl:variable name="num">
			<xsl:number count="书签:书签_9105" format="1" level="any"/>
		</xsl:variable>
		<xsl:value-of select="$num"/>
	</xsl:template>
	<xsl:template name="bkEnd">
		<w:bookmarkEnd>
      
      <!--2012-11-28，wudi，OOX到UOF方向的书签实现，修改变量id的计算方法及w:id的取值，原id计算方法同bkStart，原取值$id，start-->
			<xsl:variable name="id">
        <xsl:number count="//字:区域结束_4167[starts-with(@标识符引用_4168,'bk_')]" format="1" level="any"/>
				<!--<xsl:value-of select="@标识符引用_4168"/>-->
			</xsl:variable>
			<xsl:attribute name="w:id">
        <xsl:value-of select="$id - 1"/>
				<!--<xsl:apply-templates select="preceding::书签:书签_9105[书签:区域_9100/@区域引用_41CE=$id]" mode="bkid"/>-->
			</xsl:attribute>
      <!--end-->
      
		</w:bookmarkEnd>
	</xsl:template>
  
	<xsl:template name="annoStart">
		<w:commentRangeStart>
			<xsl:variable name="annoid">
				<xsl:value-of select="@标识符_4100"/>
			</xsl:variable>
			<xsl:attribute name="w:id">
				<xsl:apply-templates select="preceding::规则:批注_B66A[@区域引用_41CE=$annoid]" mode="annoid"/>
			</xsl:attribute>
		</w:commentRangeStart>
	</xsl:template>

	<xsl:template match="规则:批注_B66A" mode="annoid">
		<xsl:variable name="num">
			<xsl:number count="规则:批注_B66A" format="1" level="any"/>
		</xsl:variable>
		<xsl:value-of select="$num"/>
	</xsl:template>
	<xsl:template name="annoEnd">
		<xsl:variable name="annoid">
			<xsl:value-of select="@标识符引用_4168"/>
		</xsl:variable>
		<xsl:variable name="id">
			<xsl:apply-templates select="preceding::规则:批注_B66A[@区域引用_41CE=$annoid]" mode="annoid"/>
		</xsl:variable>
		<w:r>
			<xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="rpr"/>
			<w:commentReference>
				<xsl:attribute name="w:id">
					<xsl:apply-templates select="preceding::规则:批注_B66A[@区域引用_41CE=$annoid]" mode="annoid"/>
				</xsl:attribute>
			</w:commentReference>
		</w:r>
		<w:commentRangeEnd>
			<xsl:attribute name="w:id">
				<xsl:apply-templates select="preceding::规则:批注_B66A[@区域引用_41CE=$annoid]" mode="annoid"/>
			</xsl:attribute>
		</w:commentRangeEnd>
	</xsl:template>

  <!--cxl,2012.2.16,currnet node:  字:区域开始_4165--> 
  <xsl:template name="hyperLink">
		<xsl:variable name="hyperid" select="@标识符_4100"/>
		<w:hyperlink>
			<xsl:if test="preceding::超链:超级链接_AA0C/超链:书签_AA0D and preceding::超链:超级链接_AA0C/超链:链源_AA00 = $hyperid">
				<xsl:attribute name="w:anchor">
					<xsl:apply-templates select="preceding::超链:超级链接_AA0C/超链:书签_AA0D" mode="bkid"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="preceding::超链:超级链接_AA0C/超链:目标_AA01 and preceding::超链:超级链接_AA0C/超链:链源_AA00 = $hyperid">
				<xsl:attribute name="r:id">
					<xsl:apply-templates select="preceding::超链:超级链接_AA0C" mode="hyperid"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:attribute name="w:history">
				<xsl:value-of select="'1'"/>
			</xsl:attribute>
			<xsl:call-template name="run"/>
		</w:hyperlink>
	</xsl:template>
	<xsl:template match="超链:超级链接_AA0C" mode="hyperid">
		<xsl:variable name="num">
			<xsl:number count="超链:超级链接_AA0C" format="1" level="single"/>
		</xsl:variable>
		<xsl:variable name="num1">
			<xsl:value-of select="$num + 10"/>
		</xsl:variable>
		<xsl:value-of select="concat('rId',$num1)"/>
	</xsl:template>
  
  
	<xsl:template name="comments">
		<w:comments>
			<xsl:for-each select="规则:批注_B66A">
				<w:comment>
					<xsl:attribute name="w:id">
						<xsl:number count="规则:批注_B66A" format="1" level="single"/>
					</xsl:attribute>
					<xsl:attribute name="w:author">
						<xsl:variable name="author" select="@作者_41DD"/>
						<xsl:value-of select="preceding::规则:用户_B668[@标识符_4100=$author]/@姓名_41DC"/>
					</xsl:attribute>
					<xsl:attribute name="w:date">
						<xsl:value-of select="concat(@日期_41DE,'Z')"/>
					</xsl:attribute>
					<xsl:if test="@作者缩写_41DF">
						<xsl:attribute name="w:initials">
							<xsl:value-of select="@作者缩写_41DF"/>
						</xsl:attribute>
					</xsl:if>
          <xsl:for-each select="字:段落_416B | 字:文字表_416C">
            <xsl:choose>
              <xsl:when test="name(.)='字:段落_416B'">
                <xsl:call-template name="paragraph"/>
              </xsl:when>
              <xsl:when test="name(.)='字:文字表_416C'">
                <xsl:call-template name="table"/>
              </xsl:when>
            </xsl:choose>
					</xsl:for-each>
				</w:comment>
			</xsl:for-each>
		</w:comments>
	</xsl:template>
  
	<xsl:template name="simplyField">
		<xsl:param name="type"/>
    <xsl:variable name="endID" select="generate-id(following-sibling::字:域结束_41A0[1])"/>
    <xsl:if test="$type!='Eq' and $type!='eq'">
      <!--<xsl:comment>
        <xsl:value-of select="$type"/>
      </xsl:comment>-->
		  <w:fldSimple>      
        <xsl:attribute name="w:instr">

          <!--2014-04-08，wudi，页码转换可能不包含域代码，加以区分，start-->
          <xsl:choose>
            <xsl:when test="following-sibling::字:域代码_419F[1]">
              <xsl:apply-templates select="following-sibling::字:域代码_419F[1]" mode="simplyfield">
                <!--cxl,2012.3.17-->
                <xsl:with-param name="type" select="$type"/>
              </xsl:apply-templates>
            </xsl:when>

            <xsl:otherwise>
              <xsl:value-of select="$type"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->

        </xsl:attribute>
        <xsl:apply-templates select="following-sibling::字:句_419D[generate-id(following-sibling::字:域结束_41A0[1])=$endID]" mode="simplyfield"/> 
		  </w:fldSimple>    
    </xsl:if>
    <xsl:if test="$type='Eq' or $type='eq'">
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <!--cxl,2012.3.21带圈字符-->
      <xsl:call-template name="enclose"/>
      <!--<xsl:apply-templates select="following-sibling::字:域代码_419F[1]//字:句_419D" mode="simplyfield"/>-->
    </xsl:if>
	</xsl:template>

	<xsl:template match="字:域代码_419F" mode="simplyfield">
		<xsl:param name="type"/>   
		<xsl:variable name="fieldcode">
      <xsl:for-each select="字:段落_416B/字:句_419D/字:空格符_4161 | 字:段落_416B/字:句_419D/字:文本串_415B">
        <xsl:choose>
          <xsl:when test="name(.)='字:空格符_4161'">
            <xsl:value-of select="concat(.,' ')"/>
          </xsl:when>
          <xsl:when test="name(.)='字:文本串_415B'">
            <xsl:value-of select="concat(.,'')"/>
          </xsl:when>
        </xsl:choose>
			</xsl:for-each>
		</xsl:variable>
		<!--<xsl:if test="$type='revnum'">
			<xsl:value-of select="translate($fieldcode,'REVISION','REVNUM')"/>
		</xsl:if>$type='revnum' or -->
    <xsl:choose>
      <xsl:when test="$type='sectionpages'">
        <xsl:value-of select="translate($fieldcode,'SECTIONPAGES','SECTIONPAGES')"/>
      </xsl:when>
      <xsl:otherwise>
        <!--cxl,2012.3.14修改域转换，去除[DBNum1]-->
        <xsl:choose>
          <xsl:when test="not(contains($fieldcode,'DBNum1'))">
            <xsl:value-of select="$fieldcode"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat(substring-before($fieldcode,'[DBNum1]'),substring-after($fieldcode,'[DBNum1]'))"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
	</xsl:template>

  <!--<xsl:template match="字:域代码_419F" mode="enclosed">
    <xsl:variable name="fieldcode">
      <xsl:for-each select="字:段落_416B/字:句_419D/node()">
        <xsl:if test="name(.)='字:空格符_4161'">
          <xsl:value-of select="concat(.,' ')"/>
        </xsl:if>
        <xsl:if test="name(.)='字:文本串_415B'">
          <xsl:value-of select="concat(.,'')"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <xsl:value-of select="$fieldcode"/>
  </xsl:template>-->
  <xsl:template name="enclose">
    <xsl:for-each select="following-sibling::字:域代码_419F[1]">
      <xsl:for-each select="./字:段落_416B/node()">
        <xsl:choose>
          <xsl:when test="name(.)='字:句_419D'">
            <w:r>
              <xsl:for-each select="node()">
                <xsl:choose>
                  <xsl:when test="name(.)='字:句属性_4158'">
                    <xsl:apply-templates select="." mode="rpr"/>
                  </xsl:when>
                  <xsl:when test="name(.)='字:文本串_415B'">
                    <w:instrText>
                      <xsl:value-of select="."/>
                    </w:instrText>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
            </w:r>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </xsl:for-each>
    <xsl:if test="following-sibling::字:域结束_41A0[1]">
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </xsl:if>
  </xsl:template>
  
	<xsl:template match="字:句_419D" mode="simplyfield">
		<xsl:call-template name="run"/>
	</xsl:template>
</xsl:stylesheet>
