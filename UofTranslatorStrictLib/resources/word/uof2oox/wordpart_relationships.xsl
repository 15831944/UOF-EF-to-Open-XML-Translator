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
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="documentXmlRels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>    
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
      <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      
      <!-- relationship for header and footer-->
      <xsl:for-each select="//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:奇数页页眉_41F4
                           |//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:偶数页页眉_41F5
                           |//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:首页页眉_41F6">
        <Relationship Id="{generate-id(.)}"
         Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header">
          <xsl:attribute name="Target">
            <xsl:value-of select="'header'"/>
            <xsl:number count="字:奇数页页眉_41F4|字:偶数页页眉_41F5|字:首页页眉_41F6" level="any" format="1"/>
            <xsl:value-of select="'.xml'" />
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
      <xsl:for-each select="//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:奇数页页脚_41F8
                           |//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:偶数页页脚_41F9
                           |//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:首页页脚_41FA">
        <Relationship  Id="{generate-id(.)}"
         Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer">
          <xsl:attribute name="Target">
            <xsl:value-of select="'footer'"/>
            <xsl:number count="字:奇数页页脚_41F8|字:偶数页页脚_41F9|字:首页页脚_41FA" level="any" format="1"/>
            <xsl:value-of select="'.xml'" />
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
		<Relationship Id="rIdPattenFill1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill1.gif"/>
		<Relationship Id="rIdPattenFill2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill2.gif"/>
		<Relationship Id="rIdPattenFill3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill3.gif"/>
		<Relationship Id="rIdPattenFill4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill4.gif"/>
		<Relationship Id="rIdPattenFill5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill5.gif"/>
		<Relationship Id="rIdPattenFill6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill6.gif"/>
		<Relationship Id="rIdPattenFill7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill7.gif"/>
		<Relationship Id="rIdPattenFill8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill8.gif"/>
		<Relationship Id="rIdPattenFill9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill9.gif"/>
		<Relationship Id="rIdPattenFill10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill10.gif"/>
		<Relationship Id="rIdPattenFill11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill11.gif"/>
		<Relationship Id="rIdPattenFill12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill12.gif"/>
		<Relationship Id="rIdPattenFill13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill13.gif"/>
		<Relationship Id="rIdPattenFill14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill14.gif"/>
		<Relationship Id="rIdPattenFill15" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill15.gif"/>
		<Relationship Id="rIdPattenFill16" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill16.gif"/>
		<Relationship Id="rIdPattenFill17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill17.gif"/>
		<Relationship Id="rIdPattenFill18" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill18.gif"/>
		<Relationship Id="rIdPattenFill19" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill19.gif"/>
		<Relationship Id="rIdPattenFill20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill20.gif"/>
		<Relationship Id="rIdPattenFill21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill21.gif"/>
		<Relationship Id="rIdPattenFill22" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill22.gif"/>
		<Relationship Id="rIdPattenFill23" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill23.gif"/>
		<Relationship Id="rIdPattenFill24" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill24.gif"/>
		<Relationship Id="rIdPattenFill25" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill25.gif"/>
		<Relationship Id="rIdPattenFill26" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill26.gif"/>
		<Relationship Id="rIdPattenFill27" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill27.gif"/>
		<Relationship Id="rIdPattenFill28" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill28.gif"/>
		<Relationship Id="rIdPattenFill29" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill29.gif"/>
		<Relationship Id="rIdPattenFill30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill30.gif"/>
		<Relationship Id="rIdPattenFill31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill31.gif"/>
		<Relationship Id="rIdPattenFill32" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill32.gif"/>
		<Relationship Id="rIdPattenFill33" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill33.gif"/>
		<Relationship Id="rIdPattenFill34" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill34.gif"/>
		<Relationship Id="rIdPattenFill35" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill35.gif"/>
		<Relationship Id="rIdPattenFill36" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill36.gif"/>
		<Relationship Id="rIdPattenFill37" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill37.gif"/>
		<Relationship Id="rIdPattenFill38" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill38.gif"/>
		<Relationship Id="rIdPattenFill39" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill39.gif"/>
		<Relationship Id="rIdPattenFill40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill40.gif"/>
		<Relationship Id="rIdPattenFill41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill41.gif"/>
		<Relationship Id="rIdPattenFill42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill42.gif"/>
		<Relationship Id="rIdPattenFill43" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill43.gif"/>
		<Relationship Id="rIdPattenFill44" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill44.gif"/>
		<Relationship Id="rIdPattenFill45" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill45.gif"/>
		<Relationship Id="rIdPattenFill46" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill46.gif"/>
		<Relationship Id="rIdPattenFill47" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill47.gif"/>
		<Relationship Id="rIdPattenFill48" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Image.PatternFill.imagePattenFill48.gif"/>
		
      <!--footnote and endnote-->
      <xsl:if test ="//字:句_419D/字:尾注_415A">
      <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
			</xsl:if>
      <xsl:if test ="//字:句_419D/字:脚注_4159">
      <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
      </xsl:if>    
      <!--comments-->
      <xsl:if test="/uof:UOF/uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669">
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
      </xsl:if>
      <!--hyperlinks-->
      <xsl:for-each select ="//uof:链接集/超链:超级链接_AA0C"><!--[@uof:目标]-->
        <xsl:variable name ="source">
          <xsl:value-of select ="./超链:链源_AA00"/>
        </xsl:variable>
        <xsl:if test ="/uof:UOF/uof:文字处理/字:文字处理文档_4225/node()[name(.)!='字:分节_416A']//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100=$source">
          <xsl:apply-templates select ="." mode ="rels"/>
        </xsl:if>
      </xsl:for-each>
      <!--numbering-->
      <xsl:if test ="uof:式样集//式样:自动编号集_990E">
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
      </xsl:if>
      <!--picture-->
      <xsl:for-each select ="/uof:UOF/uof:文字处理/字:文字处理文档_4225/字:分节_416A/字:节属性_421B/字:填充_4134/图:图片_8005">
        <xsl:apply-templates select ="." mode ="rels"/>
      </xsl:for-each>

      <!--2014-03-18，wudi，考虑文本框里插入图片的情形，start-->
      <xsl:for-each select ="/uof:UOF/uof:对象集/node()//字:句_419D/uof:锚点_C644">
        <xsl:apply-templates select ="." mode ="rels"/>
      </xsl:for-each>
      <!--end-->
      
      <xsl:for-each select ="/uof:UOF/uof:文字处理/字:文字处理文档_4225/node()[name(.)!='字:分节_416A']//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644">
        <xsl:apply-templates select ="." mode ="rels"/>
      </xsl:for-each>
      <xsl:if test ="//对象:对象数据_D701/@标识符_D704='backgroundpattern006'">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
          <xsl:variable name ="abspath">
            <xsl:value-of select ="//对象:对象数据_D701[@标识符_D704 = 'backgroundpattern006']/u2opic:picture/@u2opic:target" xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"/>
          </xsl:variable>
          <xsl:variable name ="relpath">
            <xsl:value-of select ="substring-after($abspath,'image')"/>
          </xsl:variable>
          <xsl:attribute name="Target">
            <xsl:value-of select ="concat('media/image',$relpath)"/>
          </xsl:attribute>
          <xsl:attribute name="Id">
            <xsl:value-of select="'rIdpattern006'"/>
          </xsl:attribute>
        </Relationship>
      </xsl:if>
        <!--other-->
    </Relationships>
  </xsl:template>
<!--wcz，2013/3/24，start-->
	<xsl:template name="numberingXmlRels">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:for-each select ="//对象:对象数据_D701">
        <xsl:variable name ="imgID">
          <xsl:value-of select ="@标识符_D704"/>
        </xsl:variable>
        <xsl:if test="/uof:UOF/uof:式样集/式样:自动编号集_990E/字:自动编号_4124/字:级别_4112/字:图片符号_411B[@引用_411C=$imgID]">
          <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
            <xsl:attribute name="Target">
              <xsl:choose>
                <xsl:when test="substring-after(./对象:路径_D703,'data\')">
                  <xsl:value-of select ="concat('media/',substring-after(./对象:路径_D703,'data\'))"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="concat('media/',substring-after(./对象:路径_D703,'data/'))"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="Id">
              <xsl:value-of select="concat('rId',substring-after($imgID,'Obj'))"/>
            </xsl:attribute>
          </Relationship>
        </xsl:if>
      </xsl:for-each>
		</Relationships>
	</xsl:template>
  <!--end-->
  <xsl:template match ="图:图片_8005" mode ="rels">
    <xsl:variable name ="picid">
      <xsl:value-of select ="/uof:UOF/uof:文字处理/字:文字处理文档_4225/字:分节_416A/字:节属性_421B/字:填充_4134/图:图片_8005/@图形引用_8007"/>
    </xsl:variable>
    <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
      <xsl:variable name ="abspath">
        <xsl:value-of select ="//对象:对象数据_D701[@标识符_D704 = $picid]/对象:路径_D703"/>
      </xsl:variable>
      <xsl:variable name ="relpath">
        <!--wcz,2013,3,24,start-->
        <xsl:choose>
          <xsl:when test="substring-after($abspath,'data\')">
            <xsl:value-of select ="substring-after($abspath,'data\')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="substring-after($abspath,'data/')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--end-->
      </xsl:variable>
      <xsl:attribute name="Target">
        <xsl:value-of select ="concat('media/',$relpath)"/>
      </xsl:attribute>
      <xsl:attribute name="Id">
        <xsl:value-of select="'rIdpic1'"/>
      </xsl:attribute>
    </Relationship>
    
  </xsl:template>
  <xsl:template match ="uof:锚点_C644" mode ="rels">
    <!--<xsl:variable name="anchorType">
      <xsl:value-of select="@字:类型"/>
    </xsl:variable>-->
    <xsl:variable name="number">
      <xsl:number count="uof:锚点_C644" format="1" level="any"/>
    </xsl:variable>
    <xsl:variable name ="objref" select ="@图形引用_C62E"/>
    <xsl:variable name ="isPic">
      <xsl:choose>
       <xsl:when test ="//对象:对象数据_D701/@标识符_D704 = $objref">
        <xsl:value-of select ="'datapic'"/>
       </xsl:when>
       <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005">
         <xsl:value-of select ="'prstpic'"/>
       </xsl:when>
       <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/@组合列表_8064">
         <xsl:value-of select ="'grsppic'"/>   
       </xsl:when>
      <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037">
        <xsl:value-of select ="'objpic'"/>
      </xsl:when>
     
      <!--<xsl:if test ="not((//uof:对象集/uof:其他对象/@uof:标识符 = $objref) or (//uof:对象集/图:图形[@图:标识符 = $objref]/@图:其他对象))or(//uof:对象集/图:图形[@图:标识符 =$objref]/图:预定义图形/图:属性/图:填充/图:图片)">
        <xsl:value-of select ="'false'"/>
      </xsl:if>-->
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test ="$isPic='datapic' or $isPic='objpic' or $isPic='prstpic' or $isPic='grsppic'">
      <xsl:choose>
        <xsl:when test ="$isPic='datapic'">
          <xsl:variable name ="id" select ="@图形引用_C62E"/>
          <xsl:apply-templates select ="//对象:对象数据_D701[@标识符_D704 = $id]" mode ="rels">
            <xsl:with-param name="num" select="$number"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:when test ="$isPic='objpic'">
          <xsl:variable name ="id" select ="@图形引用_C62E"/>
          <xsl:variable name ="oid" select ="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:图片数据引用_8037"/>
          <xsl:apply-templates select ="//对象:对象数据_D701[@标识符_D704 = $oid]" mode ="rels">
            <xsl:with-param name="num" select="$number"/>
          </xsl:apply-templates>
        </xsl:when>
        <!--<xsl:when test ="$isPic='prstpic'">
        <xsl:variable name ="id" select="字:图形/@字:图形引用"/>
        <xsl:variable name ="oid" select ="//uof:对象集/图:图形[@图:标识符 = $id]/图:预定义图形/图:属性/图:填充/图:图片/@图:图形引用"/>
        <xsl:apply-templates select ="//uof:其他对象[@uof:标识符=$oid]" mode ="rels">
          <xsl:with-param name ="num" select ="$number"/>
        </xsl:apply-templates>
      </xsl:when>-->
        <xsl:when test ="$isPic='prstpic'">
          <xsl:variable name ="id" select ="@图形引用_C62E"/>
          <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007">
            <xsl:variable name ="oid" select ="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>
            <xsl:apply-templates select ="//对象:对象数据_D701[@标识符_D704 = $oid]" mode ="rels">
              <xsl:with-param name ="num" select ="$number"/>
            </xsl:apply-templates >
          </xsl:if>
        </xsl:when>
        <!--组合图形的图片填充-->
        <xsl:when test ="$isPic='grsppic'">
          <xsl:variable name ="id" select ="@图形引用_C62E"/>
          <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B = $id]/@组合列表_8064">
            <xsl:apply-templates select ="//uof:对象集/图:图形_8062[@标识符_804B = $id and @组合列表_8064]" mode ="grprels">
              <xsl:with-param name ="num" select ="$number"/>
              <xsl:with-param name ="level" select ="'1'"/>
            </xsl:apply-templates>
          </xsl:if>
        </xsl:when>
      </xsl:choose>
      
    </xsl:if>
  </xsl:template>

  <xsl:template match ="图:图形_8062" mode ="grprels">
    <xsl:param name ="num"/>
    <xsl:param name ="level"/>
    <xsl:variable name ="templist" select ="@组合列表_8064"/>
    <xsl:for-each select="//uof:对象集/图:图形_8062">
      <xsl:choose>
        <xsl:when test="contains($templist,@标识符_804B) and not(@组合列表_8064)">
          <xsl:if test ="./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007">
            <xsl:variable name ="gid" select ="./图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>
            <xsl:variable name ="number" select ="substring-after(@标识符_804B,'Obj')"/>
            <xsl:apply-templates select ="//对象:对象数据_D701[@标识符_D704 = $gid]" mode ="rels">
              <xsl:with-param name ="num" select ="$num*100+$level*10+$number"/>
            </xsl:apply-templates >
          </xsl:if>
        </xsl:when>
        <xsl:when test ="contains($templist,@标识符_804B) and @组合列表_8064">
          <xsl:apply-templates select ="." mode ="grprels">
            <xsl:with-param name ="num" select ="$num"/>
            <xsl:with-param name ="level" select ="$level+1"/>
          </xsl:apply-templates>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match ="对象:对象数据_D701" mode ="rels">
    <xsl:param name="num"/>
    <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
      <xsl:variable name ="abspath">
        <xsl:value-of select ="对象:路径_D703"/>
      </xsl:variable>
      <xsl:variable name ="relpath">
        
        <!--2012-01-23，wudi，UOF到OOX图片转换BUG：改“data\”为“data/”，start-->
        <xsl:choose>
          <xsl:when test="substring-after($abspath,'data/') != ''">
            <xsl:value-of select ="substring-after($abspath,'data/')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="substring-after($abspath,'data\')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--end-->
        
      </xsl:variable>
      <xsl:attribute name="Target">
        <xsl:value-of select ="concat('media/',$relpath)"/>
      </xsl:attribute>
      <xsl:attribute name="Id">
        <xsl:value-of select="concat('rIdObj',$num)"/>
      </xsl:attribute>
    </Relationship>
  </xsl:template>
  <xsl:template match ="超链:超级链接_AA0C" mode ="rels">
    <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" TargetMode="External">
      <xsl:variable name ="num">
        <xsl:number count ="超链:超级链接_AA0C" format ="1" level ="single"/>
      </xsl:variable>
      <xsl:variable name ="numid">
        <xsl:call-template name="addOne"><!--这个模板对id进行了加10，为了防止产生冲突-->
          <xsl:with-param name="source" select="$num"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="Id">
        <xsl:value-of select ="concat('rId',$numid)"/>
      </xsl:attribute>
      <xsl:attribute name="Target">
        <xsl:variable name ="target">
          <xsl:choose>
            <xsl:when test="./超链:目标_AA01">
              <xsl:value-of select="./超链:目标_AA01"/>
            </xsl:when>
            <xsl:when test="not(./超链:目标_AA01) and ./超链:书签_AA0D">
              <xsl:value-of select="./超链:书签_AA0D"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name ="targetname">
        <xsl:choose>
          <xsl:when test ="starts-with($target,'/')">
            <xsl:variable name ="tartemp">
              <xsl:value-of select ="translate(substring-after($target,'/'),'/','\')"/>
            </xsl:variable>
            <xsl:variable name ="tartemp1">
              <!--<xsl:value-of select ="translate($tartemp,'%5C','\')"/>-->
              <xsl:call-template name ="transtring">
                <xsl:with-param name ="input" select ="$tartemp"/>
                <xsl:with-param name ="from" select ="'%5C'"/>
                <xsl:with-param name="to" select ="'\'"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="tar">
              <!--<xsl:value-of select ="translate($tartemp1,'%20',' ')"/>-->
              <xsl:call-template name ="transtring">
              <xsl:with-param name ="input" select ="$tartemp1"/>
              <xsl:with-param name ="from" select ="'%20'"/>
              <xsl:with-param name="to" select ="' '"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select ="concat('file:///',$tar)"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'\')">
            <xsl:variable name ="tar">
              <xsl:value-of select ="translate(substring-after($target,'\'),'/','\')"/>
            </xsl:variable>
            <xsl:value-of select ="concat('file:///',$tar)"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'http:')">
            <xsl:value-of select ="超链:目标_AA01"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'ftp:')">
            <xsl:value-of select ="超链:目标_AA01"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'mailto:')">
            <xsl:value-of select ="超链:目标_AA01"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'telnet:')">
            <xsl:value-of select ="超链:目标_AA01"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'..')">
            <xsl:value-of select ="$target"/>
          </xsl:when>
          <xsl:when test ="starts-with($target,'.')">
            <xsl:variable name ="tar">
              <xsl:value-of select ="translate($target,'/','\')"/>
            </xsl:variable>
            <xsl:value-of select ="concat('file://',$tar)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="$target"/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:variable>
        <xsl:value-of select ="$targetname"/>
      </xsl:attribute>
    </Relationship>
  </xsl:template>
  <xsl:template name ="transtring">
    <xsl:param name ="input"/>
    <xsl:param name ="from"/>
    <xsl:param name ="to"/>
    <xsl:if test ="$input">
      <xsl:choose>
        <xsl:when test ="contains($input,$from)">
          <xsl:value-of select ="substring-before($input,$from)"/>
          <xsl:value-of select ="$to"/>
          <xsl:call-template name ="transtring">
            <xsl:with-param name ="input" select ="substring-after($input,$from)"/>
            <xsl:with-param name ="from" select ="$from"/>
            <xsl:with-param name ="to" select ="$to"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="$input"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name="addOne">
    <xsl:param name="source"/>
    <xsl:value-of select="$source + 10"/>
  </xsl:template>
  <xsl:template name ="otherXmlRels">
    <xsl:param name ="relType"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--header,footer-->
      <xsl:if test ="contains($relType,'页眉') or contains($relType,'页脚') ">
        <xsl:for-each select =".//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink']">
          <xsl:variable name ="hyperid" select ="@标识符_4100"/>     
          <xsl:apply-templates select ="//uof:链接集/超链:超级链接_AA0C[超链:链源_AA00=$hyperid]" mode ="rels"/>
        </xsl:for-each>
        <xsl:for-each select =".//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644">
          <xsl:variable name="number">
            <xsl:number count="uof:锚点_C644" format="1" level="any"/>
          </xsl:variable>
          <xsl:apply-templates select ="." mode ="rels">
            <xsl:with-param name="num" select="$number"/>
          </xsl:apply-templates>
        </xsl:for-each>
      </xsl:if>
      <!--endnote,footnote-->
      <xsl:if test ="$relType='footnotes'">
 
        <xsl:for-each select ="//uof:链接集/超链:超级链接_AA0C[超链:目标_AA01]">
          <xsl:variable name ="source">
            <xsl:value-of select ="超链:链源_AA00"/>
          </xsl:variable>
          <xsl:if test ="//字:脚注_4159//字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100=$source">
            <xsl:apply-templates select ="." mode ="rels"/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select ="//字:脚注_4159//uof:锚点_C644">
          <xsl:variable name="number">
            <xsl:number count="uof:锚点_C644" format="1" level="any"/>
          </xsl:variable>
          <xsl:variable name ="objref" select ="@图形引用_C62E"/>
          <xsl:variable name ="hasDef">
            <xsl:if test ="preceding::字:尾注_415A//uof:锚点_C644[@图形引用_C62E=$objref]">
              <xsl:value-of select ="'true'"/>
            </xsl:if>
            <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037">
              <xsl:variable name ="oobjid">
                <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037"/>
              </xsl:variable>
              <xsl:if test ="preceding::字:尾注_415A//uof:锚点_C644/@图形引用_C62E = //uof:对象集/图:图形_8062[图:图片数据引用_8037=$oobjid]/@标识符_804B">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:if>
          </xsl:variable>
          <xsl:if test ="$hasDef != 'true'">
            <xsl:apply-templates select ="." mode ="rels">
              <xsl:with-param name="num" select="$number"/>
            </xsl:apply-templates>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test ="$relType='endnotes'">
      
        <xsl:for-each select ="//uof:链接集/超链:超级链接_AA0C[超链:目标_AA01]">
          <xsl:variable name ="source">
            <xsl:value-of select ="超链:链源_AA00"/>
          </xsl:variable>
          <xsl:if test ="//字:尾注_415A//字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100=$source">
            <xsl:apply-templates select ="." mode ="rels"/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select ="//字:尾注_415A//uof:锚点_C644">
          <xsl:variable name="number">
            <xsl:number count="uof:锚点_C644" format="1" level="any"/>
          </xsl:variable>
          <xsl:variable name ="objref" select ="@图形引用_C62E"/>
          <xsl:variable name ="hasDef">
            <xsl:if test ="preceding::字:尾注_415A//uof:锚点_C644[@图形引用_C62E=$objref]">
              <xsl:value-of select ="'true'"/>
            </xsl:if>
            <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037">
              <xsl:variable name ="oobjid">
                <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037"/>
              </xsl:variable>
              <xsl:if test ="preceding::字:尾注_415A//uof:锚点_C644/@图形引用_C62E = //uof:对象集/图:图形_8062[图:图片数据引用_8037=$oobjid]/@标识符_804B">
                <xsl:value-of select ="'true'"/>
              </xsl:if>
            </xsl:if>
          </xsl:variable>
          <xsl:if test ="$hasDef != 'true'">
            <xsl:apply-templates select ="." mode ="rels">
              <xsl:with-param name="num" select="$number"/>
            </xsl:apply-templates>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <!--comment-->
      <xsl:if test ="$relType='comments'">
        <xsl:for-each select ="//uof:链接集/超链:超级链接_AA0C[超链:目标_AA01]">
          <xsl:variable name ="source">
            <xsl:value-of select ="超链:链源_AA00"/>
          </xsl:variable>
          <xsl:if test ="/uof:UOF/uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100=$source">
            <xsl:apply-templates select ="." mode ="rels"/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select ="/uof:UOF/uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644">
          <xsl:variable name="number">
            <xsl:number count="uof:锚点_C644" format="1" level="any"/>
          </xsl:variable>
          <xsl:apply-templates select ="." mode ="rels">
            <xsl:with-param name="num" select="$number"/>
          </xsl:apply-templates>
        </xsl:for-each>
      </xsl:if>

    </Relationships>
  </xsl:template>
</xsl:stylesheet>
