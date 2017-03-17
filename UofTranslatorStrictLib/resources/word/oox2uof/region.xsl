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
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">
	<xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  
	<xsl:template name="regionStart">
		<xsl:param name="id"/>
		<xsl:param name="name"/>
		<xsl:param name="type"/>
		<字:区域开始_4165>
			<xsl:attribute name="标识符_4100"><xsl:value-of select="$id"/></xsl:attribute>
			<xsl:attribute name="名称_4166"><xsl:value-of select="$name"/></xsl:attribute>
			<xsl:attribute name="类型_413B"><xsl:value-of select="$type"/></xsl:attribute>
		</字:区域开始_4165>
	</xsl:template>

	<xsl:template name="regionEnd">
		<xsl:param name="id"/>
		<字:区域结束_4167>
			<xsl:attribute name="标识符引用_4168"><xsl:value-of select="$id"/></xsl:attribute>
		</字:区域结束_4167>
	</xsl:template>
  <!--书签********************************************************************************************-->
	<!--<xsl:template name="bookmarkStart">
		<字:句_419D>
      --><!--<w:bookmarkStart>--><!--
        <xsl:attribute name ="w:id">
          <xsl:call-template name ="numbookmark"/>
        </xsl:attribute>
        <xsl:if test ="@w:name">
          <xsl:attribute name ="w:name">
            <xsl:value-of select ="@w:name"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="@w:colFirst">
          <xsl:attribute name ="w:colFirst">
            <xsl:value-of select ="@w:colFirst"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="@w:colLast">
          <xsl:attribute name ="w:colLast">
            <xsl:value-of select ="@w:colLast"/>
          </xsl:attribute>
        </xsl:if>
      --><!--</w:bookmarkStart>--><!--
		</字:句_419D>
	</xsl:template>
  <xsl:template name ="numbookmark">
    <xsl:value-of select ="concat('bk_',generate-id(.))"/>
  </xsl:template>
  <xsl:template match ="w:bookmarkStart" mode ="num">
    <xsl:call-template name ="numbookmark"/>
  </xsl:template>
	<xsl:template name="bookmarkEnd">	
    <w:bookmarkEnd>
      <xsl:variable name ="id">
        <xsl:value-of select ="@w:id"/>
      </xsl:variable>
      <xsl:attribute name ="w:id">
        <xsl:apply-templates select ="//w:bookmarkStart[@w:id=$id]" mode ="num"/>
      </xsl:attribute>
    </w:bookmarkEnd>
	</xsl:template>-->
  <!--**************************************************************************************************************-->
	<xsl:template name="commentSet">
		<xsl:if test="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Target='comments.xml']">
			<xsl:apply-templates select="document('word/comments.xml')/w:comments" mode="comSet"/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:comments" mode="comSet">
		<规则:批注集_B669>
			<xsl:for-each select="./w:comment">
				<规则:批注_B66A>
					<xsl:attribute name="区域引用_41CE"><xsl:value-of select="concat('cmt_',./@w:id)"/></xsl:attribute>
					<xsl:attribute name="作者_41DD"><xsl:value-of select="concat('aut2_',./@w:author)"/></xsl:attribute>
             
          <xsl:attribute name="日期_41DE"><xsl:value-of select="substring(./@w:date,1,19)"/></xsl:attribute>
          
					<xsl:attribute name="作者缩写_41DF"><xsl:value-of select="./@w:initials"/></xsl:attribute>
          <xsl:for-each select="w:p | w:tbl">
            <xsl:if test ="name(.)='w:p'">
              <xsl:call-template name="paragraph">
                <xsl:with-param name ="pPartFrom" select ="'comments'"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="name(.)='w:tbl'">
              <xsl:call-template name="table">
                <xsl:with-param name ="tblPartFrom" select ="'comments'"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:for-each>
				</规则:批注_B66A>
			</xsl:for-each>
		</规则:批注集_B669>
	</xsl:template>

	<xsl:template name="commentStart">
		<字:句_419D>
			<xsl:call-template name="regionStart">
				<xsl:with-param name="id" select="concat('cmt_',@w:id)"/>
				<xsl:with-param name="name" select="'annotation'"/>
				<xsl:with-param name="type" select="'annotation'"/>
			</xsl:call-template>
		</字:句_419D>
	</xsl:template>
	<xsl:template name="commentEnd">
		<字:句_419D>
      <!--<xsl:comment>comment end</xsl:comment>-->
			<xsl:call-template name="regionEnd">
				<xsl:with-param name="id" select="concat('cmt_',@w:id)"/>
			</xsl:call-template>     
		</字:句_419D>
	</xsl:template>
  
  <!--cxl修改于11月6日,2012.3.14修改，增加由域实现的超链接转换-->
  <xsl:template name="hyperlinkRegion">
    <xsl:param name ="filename"/>
    <字:句_419D>
      <xsl:if test="name(.)='w:hyperlink' or (name(.)='w:r' and ./w:fldChar)">
        <字:区域开始_4165> 
          <xsl:attribute name="标识符_4100">
            <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
          </xsl:attribute>
          <xsl:attribute name="名称_4166">
            <xsl:value-of select="'hyperlink'"/>
          </xsl:attribute>
          <xsl:attribute name="类型_413B">
            <xsl:value-of select="'hyperlink'"/>
          </xsl:attribute>
        
          <!--<xsl:if test="@r:id">
            <xsl:variable name="rid">
              <xsl:value-of select="@r:id"/>
            </xsl:variable>
            <xsl:variable name ="goal">--><!--超链接几种要转移到的目标--><!-- 
              <xsl:choose>
                <xsl:when test ="$filename='document'">
                  <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
                <xsl:when test ="$filename='comments'">
                  <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
                <xsl:when test ="$filename='endnotes'">
                  <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
                <xsl:when test ="$filename='footnotes'">
                  <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
                <xsl:when test ="contains($filename,'header')">
                  <xsl:variable name ="doc">
                    <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
                  </xsl:variable>
                  <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
                <xsl:when test ="contains($filename,'footer')">
                  <xsl:variable name ="doc">
                    <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
                  </xsl:variable>
                  <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
                </xsl:when>
              </xsl:choose>
            </xsl:variable>
            <超链:目标_AA01>--><!--这个地方要不要修改？？？？？？？？？？？？？？？？？？？？？？？链接集还有其它几个子元素--><!-- 
              <xsl:value-of select="$goal"/>
            </超链:目标_AA01> 
          </xsl:if>
          <xsl:if test="@w:anchor">
            <超链:书签_AA0D>
              <xsl:value-of select="./@w:anchor"/>
            </超链:书签_AA0D>      
          </xsl:if>
          <xsl:if test="@w:tooltip">
            <超链:提示_AA05>
              <xsl:value-of select="./@w:tooltip"/>
            </超链:提示_AA05>
          </xsl:if>
          <超链:链源_AA00>
            <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
          </超链:链源_AA00>-->
        </字:区域开始_4165>
      </xsl:if>
    </字:句_419D>

    <!--2013-03-27，wudi，修复超链接转换BUG，增加限制条件not(name(.)='w:hyperlink'，start-->
    <xsl:if test ="not(name(.)='w:hyperlink')">
      <xsl:for-each select="following-sibling::w:r[w:t]">
        <字:句_419D>
          <xsl:call-template name="run">
            <xsl:with-param name ="rPartFrom" select ="$filename"/>
          </xsl:call-template>
        </字:句_419D>
      </xsl:for-each>
    </xsl:if>
    <!--end-->
    
    <xsl:for-each select="w:r">

      <!--2013-03-27，wudi，修复索引丢失的BUG，之前的处理方式有问题，w:instrText取值不全，start-->
      <xsl:if test ="./w:fldChar/@w:fldCharType ='begin'">
        
        <!--2013-04-15，针对空格符做二次处理，以免丢了空格符的转换，增加PageRef的情况，start-->
        <xsl:variable name ="temp1">
          <xsl:for-each select="./ancestor::w:hyperlink//w:instrText">
            <xsl:value-of select ="."/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:variable name ="temp">
          <xsl:value-of select ="normalize-space($temp1)"/>
        </xsl:variable>
        <xsl:variable name ="tmp">
          <xsl:if test ="contains(following-sibling::w:r[1]/w:instrText,'XE')">
            <xsl:value-of select ="substring-before($temp,'PAGEREF')"/>
          </xsl:if>
          <xsl:if test ="contains(following-sibling::w:r[1]/w:instrText,'PAGEREF')">
            <xsl:value-of select ="following-sibling::w:r[1]/w:instrText"/>
          </xsl:if>
          <xsl:if test ="contains(following-sibling::w:r[1]/w:instrText,'PageRef')">
            <xsl:value-of select ="$temp"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name ="type">
          <xsl:choose>
            <xsl:when test ="contains($tmp,'XE')">
              <xsl:value-of select ="'XE'"/>
            </xsl:when>
            <xsl:when test ="contains($tmp,'PAGEREF')">
              <xsl:value-of select ="'PAGEREF'"/>
            </xsl:when>
            <xsl:when test ="contains($tmp,'PageRef')">
              <xsl:value-of select ="'PAGEREF'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <!--end-->
        
        <字:域开始_419E>
          <xsl:attribute name="类型_416E">
            <xsl:value-of select ="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
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
        <xsl:choose>
          <xsl:when test ="$type ='XE'">
            <xsl:variable name ="tmp1">
              <xsl:value-of select ="substring-after($tmp,'XE')"/>
            </xsl:variable>
            <xsl:variable name ="tmp2">
              <xsl:value-of select ="normalize-space($tmp1)"/>
            </xsl:variable>
            <字:域代码_419F>
              <字:段落_416B>
                <字:句_419D>
                  <字:句属性_4158/>
                  <字:空格符_4161/>
                  <字:文本串_415B>XE</字:文本串_415B>
                  <字:空格符_4161/>
                  <字:文本串_415B>
                    <xsl:value-of select ="$tmp2"/>
                  </字:文本串_415B>
                  <字:空格符_4161/>
                </字:句_419D>
              </字:段落_416B>
            </字:域代码_419F>
            <xsl:variable name ="strlgth">
              <xsl:value-of select ="string-length($tmp2)"/>
            </xsl:variable>
            <xsl:variable name ="lgth">
              <xsl:number value ="$strlgth - 2"/>
            </xsl:variable>
            <字:句_419D>
              <字:句属性_4158/>
              <字:文本串_415B>
                <xsl:value-of select ="substring($tmp2,2,$lgth)"/>
              </字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test ="$type ='PAGEREF'">
            <字:域代码_419F>
              <字:段落_416B>
                <字:句_419D>
                  <字:句属性_4158/>
                  <字:文本串_415B>
                    <xsl:value-of select ="$tmp"/>
                  </字:文本串_415B>
                </字:句_419D>
              </字:段落_416B>
            </字:域代码_419F>
          </xsl:when>
        </xsl:choose>
      </xsl:if>

      <xsl:if test="./w:fldChar[@w:fldCharType='separate']">
        <字:句_419D>
          <字:句属性_4158/>
          <字:文本串_415B/>
        </字:句_419D>
      </xsl:if>
      
      <xsl:if test="./w:fldChar[@w:fldCharType='end']">
        <字:域结束_41A0/>
      </xsl:if>

      <xsl:if test="not(./w:fldChar[@w:fldCharType='begin']) and not(./w:instrText) and not(./w:fldChar[@w:fldCharType='separate']) and not(./w:instrText[@xml:space='preserve']) and not(./w:fldChar[@w:fldCharType='end'])">
        <字:句_419D>
          <xsl:call-template name="run">
            <xsl:with-param name ="rPartFrom" select ="$filename"/>
          </xsl:call-template>
        </字:句_419D>
      </xsl:if>
      <!--end-->
      
    </xsl:for-each>

    <!--2013-04-15，wudi，针对部分特例，区域开始和区域结束不在一个段落的情况，加限制条件，start-->
    <xsl:if test ="name(.)='w:hyperlink' or (name(.)='w:r' and following-sibling::w:r/w:fldChar/@w:fldCharType='end')">
      <字:句_419D>
        <字:区域结束_4167>
          <xsl:attribute name="标识符引用_4168">
            <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
          </xsl:attribute>
        </字:区域结束_4167>
      </字:句_419D>
    </xsl:if>
    <!--end-->
    
  </xsl:template>

	<xsl:template name="userSet">
		<xsl:if test ="(document('word/document.xml')/w:document/w:body//node()[@w:author]) or (document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Target='comments.xml'])">
			<规则:用户集_B667>
				<xsl:if test ="document('word/document.xml')/w:document/w:body//node()[@w:author]">
				<xsl:apply-templates select="document('word/document.xml')/w:document/w:body" mode="userSet"/>
				</xsl:if>
				<xsl:if test ="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Target='comments.xml']">
					<xsl:apply-templates select="document('word/comments.xml')/w:comments" mode="userSet"/>
				</xsl:if>
			</规则:用户集_B667>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:body" mode="userSet">
		<xsl:if test=".//node()[@w:author]">
				<xsl:for-each select="//node()[@w:author]">
					<xsl:if test="not(current()/@w:author = preceding::node()[@w:author]/@w:author)">
						<规则:用户_B668>
							<xsl:attribute name="姓名_41DC"><xsl:value-of select="./@w:author"/></xsl:attribute>
							<xsl:attribute name="标识符_4100"><xsl:value-of select="concat('aut1_',./@w:author)"/></xsl:attribute>
						</规则:用户_B668>
					</xsl:if>
				</xsl:for-each>
		</xsl:if>
	</xsl:template>
	<xsl:template match ="w:comments" mode ="userSet">
		<xsl:if test=".//node()[@w:author]">
			<xsl:for-each select="//node()[@w:author]">
				<xsl:if test="not(current()/@w:author = preceding::node()[@w:author]/@w:author)">
          <规则:用户_B668>
						<xsl:attribute name="姓名_41DC">
							<xsl:value-of select="./@w:author"/>
						</xsl:attribute>
						<xsl:attribute name="标识符_4100">
							<xsl:value-of select="concat('aut2_',./@w:author)"/>
						</xsl:attribute>
					</规则:用户_B668>
				</xsl:if>
			</xsl:for-each>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
