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
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
  xmlns:fo="http://www.w3.org/1999/XSL/Format" 
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
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" 
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <xsl:import href="header-footer.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <!--节属性模板cxl2011/11/12-->
  <xsl:template name="section">
    <字:分节_416A>
      <!--<xsl:if test ="w:sectPrChange">--><!--节属性修订信息--><!--
        <xsl:apply-templates select ="w:sectPrChange/w:sectPr" mode ="revSect"/>
        <xsl:call-template name ="revisionSectProperties"/>
      </xsl:if>-->
      <!--<xsl:if test ="not(w:sectPrChange)">--><!--节属性模板-->
        <xsl:call-template name ="sectProperties"/>
      <!--</xsl:if>-->
    </字:分节_416A>
  </xsl:template>
  
  <!--<xsl:template match ="w:sectPr" mode ="revSect">
    <xsl:call-template name ="sectProperties"/>
  </xsl:template>-->
 
  <xsl:template name ="sectProperties">
	  <字:节属性_421B>
      <xsl:if test="./w:type">
		    <字:节类型_41EA>
          <xsl:choose>
            <xsl:when test="(./w:type/@w:val)='continuous'">
              <xsl:value-of select="'continuous'"/>
            </xsl:when>
            <xsl:when test="(./w:type/@w:val)='evenPage'">
              <xsl:value-of select="'even-page'"/>
            </xsl:when>
            <xsl:when test="(./w:type/@w:val)='oddPage'">
              <xsl:value-of select="'odd-page'"/>
            </xsl:when>
            <xsl:when test="(./w:type/@w:val)='nextColumn'">
              <xsl:value-of select="'new-column'"/>
            </xsl:when>
            <xsl:when test="(./w:type/@w:val)='nextPage'">
              <xsl:value-of select="'new-page'"/>
            </xsl:when>
          </xsl:choose>
        </字:节类型_41EA>
      </xsl:if>
      <xsl:if test="not(./w:type)">
        <字:节类型_41EA>
          <xsl:value-of select="'new-page'"/>
        </字:节类型_41EA>
      </xsl:if>
      <xsl:if test="./w:pgSz">
        <xsl:variable name ="printTwoOnOne">
          <xsl:if test ="document('word/settings.xml')/w:settings/w:printTwoOnOne[(not(@w:val)) or (@w:val='true') or  (@w:val='on') or  (@w:val='1')]">
            <xsl:value-of select ="'true'"/>
          </xsl:if>
          <xsl:if test ="document('word/settings.xml')/w:settings/w:printTwoOnOne[ (@w:val='false') or  (@w:val='off') or  (@w:val='0')]">
            <xsl:value-of select ="'false'"/>
          </xsl:if>
          <xsl:if test ="document('word/settings.xml')/w:settings[not(w:printTwoOnOne)]">
            <xsl:value-of select ="'false'"/>
          </xsl:if>
        </xsl:variable>
		    <字:纸张_41EC>
          <xsl:if test ="$printTwoOnOne='true'">
            <xsl:if test ="not(w:pgSz/@w:orient) or w:pgSz/@w:orient='portrait'"><!-- 纸短的一边做顶和底-->
              <xsl:attribute name="宽_C605">
                <xsl:value-of select="(./w:pgSz/@w:w) div 20"/>
              </xsl:attribute>
              <xsl:attribute name="长_C604">
                <xsl:value-of select="(./w:pgSz/@w:h) div 10"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="w:pgSz/@w:orient='landscape'"><!-- 纸长的一边做顶和底-->
              <xsl:attribute name="宽_C605">
                <xsl:value-of select="(./w:pgSz/@w:w) div 10"/>
              </xsl:attribute>
              <xsl:attribute name="长_C604">
                <xsl:value-of select="(./w:pgSz/@w:h) div 20"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          <xsl:if test ="$printTwoOnOne='false'">
            <xsl:attribute name="宽_C605">
              <xsl:value-of select="(./w:pgSz/@w:w) div 20"/>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:value-of select="(./w:pgSz/@w:h) div 20"/>
            </xsl:attribute>
          </xsl:if>
        </字:纸张_41EC>
        <xsl:if test ="w:pgSz/@w:orient">
	     	  <字:纸张方向_41FF>
            <xsl:value-of select ="w:pgSz/@w:orient"/>
          </字:纸张方向_41FF>
        </xsl:if>
        <xsl:if test ="$printTwoOnOne='true'">
		       <字:是否拼页_41FE>  
              <xsl:value-of select ="'true'"/>
          </字:是否拼页_41FE>
        </xsl:if>
      </xsl:if>
      <xsl:if test="./w:pgMar">
		    <字:页边距_41EB>
            <xsl:attribute name="左_C608">
              <xsl:value-of select="(./w:pgMar/@w:left) div 20"/>
            </xsl:attribute>
            <xsl:attribute name="上_C609">
              <xsl:value-of select="(./w:pgMar/@w:top) div 20"/>
            </xsl:attribute>
            <xsl:attribute name="右_C60A">
              <xsl:value-of select="(./w:pgMar/@w:right) div 20"/>
            </xsl:attribute>
            <xsl:attribute name="下_C60B">
              <xsl:value-of select="(./w:pgMar/@w:bottom) div 20"/>
            </xsl:attribute>
        </字:页边距_41EB>
		    <字:装订线_41FB>
          <xsl:attribute name="位置_4150">
            <xsl:apply-templates select ="document('word/settings.xml')/w:settings" mode ="gutter"/>
          </xsl:attribute>
          <xsl:attribute name="距边界_41FC">
            <xsl:value-of select="(./w:pgMar/@w:gutter) div 20"/>
          </xsl:attribute>
        </字:装订线_41FB>
		    <字:页眉位置_41EF>
          <xsl:attribute name="距边界_41F0">
            <xsl:value-of select="(./w:pgMar/@w:header) div 20"/>
          </xsl:attribute>
        </字:页眉位置_41EF>
		    <字:页脚位置_41F2>
          <xsl:attribute name="距边界_41F0">
            <xsl:value-of select="(./w:pgMar/@w:footer) div 20"/>
          </xsl:attribute>
        </字:页脚位置_41F2>
      </xsl:if>
      <!--杨晓，将./w:cols改为./w:cols[@w:num]，为了页面设置的情况09.12.18-->
      <xsl:if test="./w:cols[@w:num]">
		    <字:分栏_4215>
          <xsl:if test="./w:cols/@w:num">
			      <字:栏数_41E8>
              <xsl:value-of select="./w:cols/@w:num"/>
			      </字:栏数_41E8>
          </xsl:if>
          <xsl:if test="w:cols/@w:sep">
            <字:分隔线_41E3>
				      <xsl:attribute name="分隔线线型_41E4">
					      <xsl:value-of select="'single'"/>
				      </xsl:attribute>
				      <xsl:attribute name="分隔线宽度_41E6">
					      <xsl:value-of select="'1'"/>
				      </xsl:attribute>
				      <xsl:attribute name="分隔线颜色_41E7">
					      <xsl:value-of select="'#000000'"/>
				      </xsl:attribute>
			      </字:分隔线_41E3> 
          </xsl:if>
			    <xsl:if test="not(w:cols/@w:equalWidth) or w:cols/@w:equalWidth='1' or w:cols/@w:equalWidth='true'or w:cols/@w:equalWidth='on'">
            <字:是否等宽_41E9>
				      <xsl:value-of select="'true'"/>
			      </字:是否等宽_41E9>
            <xsl:variable name="colnum" select="./w:cols/@w:num"/>
            <xsl:call-template name ="getCol">
              <xsl:with-param name ="num" select ="$colnum"/>
            </xsl:call-template>
          </xsl:if>
          <xsl:if test="w:cols/@w:equalWidth='0' or w:cols/@w:equalWidth='false'or w:cols/@w:equalWidth='off'">
            <字:是否等宽_41E9>
				      <xsl:value-of select="'false'"/>
			      </字:是否等宽_41E9>
            <xsl:for-each select="./w:cols/w:col">
			        <字:栏_41E0>
                <xsl:if test="@w:w">
                  <xsl:attribute name="宽度_41E1">
                    <xsl:value-of select="(./@w:w) div 20"/>
                  </xsl:attribute>                
                </xsl:if>
                <xsl:if test="@w:space">
                  <xsl:attribute name="间距_41E2">
                    <xsl:value-of select="(./@w:space) div 20"/>
                  </xsl:attribute>
                </xsl:if>
              </字:栏_41E0>
            </xsl:for-each>
          </xsl:if>
        </字:分栏_4215>
      </xsl:if>
      <xsl:if test="w:endnotePr">
		    <字:尾注设置_4204>
          <xsl:if test="w:endnotePr/w:numFmt">
            <xsl:attribute name="格式_4151">
              <xsl:call-template name="notePrNumFmt">
                <xsl:with-param name="val" select="./w:endnotePr/w:numFmt/@w:val"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:endnotePr/w:numStart">
            <xsl:attribute name="起始编号_4152">
              <xsl:value-of select="./w:endnotePr/w:numStart/@w:val"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:endnotePr/w:numRestart">
            <xsl:attribute name="编号方式_4153">
              <xsl:choose>
                <xsl:when test="(./w:endnotePr/w:numRestart/@w:val)='continuous'">
                  <xsl:value-of select="'continuous'"/>
                </xsl:when>
                <xsl:when test="(./w:endnotePr/w:numRestart/@w:val)='eachSect'">
                  <xsl:value-of select="'section'"/>
                </xsl:when>
                <xsl:when test="(./w:endnotePr/w:numRestart/@w:val)='eachPage'">
                  <xsl:value-of select="'page'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </字:尾注设置_4204>
      </xsl:if>
      <xsl:if test="w:footnotePr">
		    <字:脚注设置_4203>
          <xsl:if test="w:footnotePr/w:numFmt">
            <xsl:attribute name="格式_4151">
              <xsl:call-template name="notePrNumFmt">
                <xsl:with-param name="val" select="./w:footnotePr/w:numFmt/@w:val"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:footnotePr/w:numStart">
            <xsl:attribute name="起始编号_4152">
              <xsl:value-of select="./w:footnotePr/w:numStart/@w:val"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:footnotePr/w:numRestart">
            <xsl:attribute name="编号方式_4153">
              <xsl:choose>
                <xsl:when test="(./w:footnotePr/w:numRestart/@w:val)='continuous'">
                  <xsl:value-of select="'continuous'"/>
                </xsl:when>
                <xsl:when test="(./w:footnotePr/w:numRestart/@w:val)='eachSect'">
                  <xsl:value-of select="'section'"/>
                </xsl:when>
                <xsl:when test="(./w:footnotePr/w:numRestart/@w:val)='eachPage'">
                  <xsl:value-of select="'page'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:footnotePr/w:pos">
            <xsl:attribute name="位置_4150">
              <xsl:choose>
                <xsl:when test="(./w:footnotePr/w:pos/@w:val)='pageBottom'">
                  <xsl:value-of select="'page-bottom'"/>
                </xsl:when>
                <xsl:when test="(./w:footnotePr/w:pos/@w:val)='beneathText'">
                  <xsl:value-of select="'below-text'"/>
                </xsl:when>
                <xsl:when test="(./w:footnotePr/w:pos/@w:val)='sectEnd'">
                  <xsl:value-of select="'page-bottom'"/>
                </xsl:when>
                <xsl:when test="(./w:footnotePr/w:pos/@w:val)='docEnd'">
                  <xsl:value-of select="'page-bottom'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </字:脚注设置_4203>
      </xsl:if>
      <xsl:if test="./w:lnNumType">
		    <字:行号设置_420A 是否使用行号_420B="true">
          <xsl:if test="w:lnNumType/@w:countBy">
            <xsl:attribute name="行号间隔_420D">
              <xsl:value-of select="./w:lnNumType/@w:countBy"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:lnNumType/@w:distance">
            <xsl:attribute name="距边界_41F0">
              <xsl:value-of select="(./w:lnNumType/@w:distance) div 20"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:lnNumType/@w:restart">
            <xsl:attribute name="编号方式_4153">
              <xsl:call-template name="lnNumTypeRestart">
                <xsl:with-param name="val" select="./w:lnNumType/@w:restart"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:lnNumType/@w:start">
            <xsl:attribute name="起始编号_420C">
              <xsl:value-of select="./w:lnNumType/@w:start + 1"/>
            </xsl:attribute>
          </xsl:if>
        </字:行号设置_420A>
      </xsl:if>
      <!--bidi-->
      <!--docGrid-->
      <xsl:if test="w:docGrid">
		  <字:网格设置_420E>
        
          <!--2013-03-08，wudi，修复文档网格-行跨度转换BUG，原行跨度_4243和列跨度_4244的取值弄反，start-->
          <xsl:if test ="w:docGrid/@w:charSpace">
            <xsl:attribute name ="列跨度_4244">
              <!--yx，this is my changement,but with this changement there is still a little different,so i recover the previous version<xsl:value-of select ="409.5 div ((w:docGrid/@w:charSpace div 4096)+10.5)"/>-->
              <xsl:value-of select ="(w:docGrid/@w:charSpace div 4096)+10.5"/>              
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="w:docGrid/@w:linePitch">
            <xsl:attribute name ="行跨度_4243">
              <xsl:value-of select ="w:docGrid/@w:linePitch div 20"/>
            </xsl:attribute>
          </xsl:if>
          <!--end-->
        
          <xsl:if test ="w:docGrid/@w:type">
            <xsl:attribute name ="网格类型_420F">
              <xsl:choose>
                <xsl:when test ="w:docGrid/@w:type='default'">
                  <xsl:value-of select ="'none'"/>
                </xsl:when>
                <xsl:when test ="w:docGrid/@w:type='linesAndChars'">
                  <xsl:value-of select ="'line-char'"/>
                </xsl:when>
                <xsl:when test ="w:docGrid/@w:type='lines'">
                  <xsl:value-of select ="'line'"/>
                </xsl:when>
                <xsl:when test ="w:docGrid/@w:type='snapToChars'">
                  <xsl:value-of select ="'char'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </字:网格设置_420E>
      </xsl:if>
      <xsl:if test="w:paperSrc">
		    <字:纸张来源_4200>
          <xsl:if test="w:paperSrc/@w:first">
            <xsl:attribute name="首页_4201">
              <xsl:value-of select="./w:paperSrc/@w:first"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:paperSrc/@w:other">
            <xsl:attribute name="其他页_4202">
              <xsl:value-of select="./w:paperSrc/@w:other"/>
            </xsl:attribute>
          </xsl:if>
        </字:纸张来源_4200>
      </xsl:if>
      <xsl:if test="w:pgBorders"><!--页面边框-->
        <xsl:apply-templates select="w:pgBorders" mode="section"/>
      </xsl:if>
      <xsl:if test="w:pgNumType">
		    <字:页码设置_4205>
          <xsl:if test="w:pgNumType/@w:fmt">
            <xsl:attribute name="格式_4151">
              <xsl:call-template name ="notePrNumFmt">
                <xsl:with-param name ="val" select ="./w:pgNumType/@w:fmt"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:pgNumType/@w:start">
            <xsl:attribute name="起始编号_4152">
              <xsl:value-of select="./w:pgNumType/@w:start"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:pgNumType/@w:chapStyle">
            <xsl:attribute name="章节起始样式引用_4208">
              <xsl:variable name ="id">
                <xsl:value-of select="./w:pgNumType/@w:chapStyle"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="starts-with($id,'id_')">
                  <xsl:value-of select ="$id"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="concat('id_',$id)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="w:pgNumType/@w:chapSep">
            <xsl:attribute name="分隔符_4209">
              <xsl:call-template name ="pgNumChapSep">
                <xsl:with-param name ="val" select ="./w:pgNumType/@w:chapSep"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </字:页码设置_4205>
      </xsl:if>
      <xsl:if test="w:textDirection">
		  <字:文字排列方向_4214>
          <xsl:call-template name ="textDirection">
            <xsl:with-param name ="dir" select ="./w:textDirection/@w:val"/>
          </xsl:call-template>
        </字:文字排列方向_4214>
      </xsl:if>
      <xsl:if test ="w:titlePg">
		  <字:是否首页页眉页脚不同_41EE>       
            <xsl:if test ="w:titlePg[not(@w:val) or @w:val='true' or @w:val='on' or @w:val='1']">
              <xsl:value-of select ="'true'"/>
            </xsl:if>
            <xsl:if test ="w:titlePg[@w:val='false' or @w:val='off' or @w:val='0']">
              <xsl:value-of select ="'false'"/>
            </xsl:if>
        </字:是否首页页眉页脚不同_41EE>
      </xsl:if>
      <xsl:if test="w:vAlign">
		  <字:垂直对齐方式_4213>
          <xsl:if test ="w:vAlign/@w:val='both'">
            <xsl:value-of select="'justified'"/>
          </xsl:if>
          <xsl:if test ="w:vAlign/@w:val!='both'">
            <xsl:value-of select="w:vAlign/@w:val"/>
          </xsl:if>
        </字:垂直对齐方式_4213>
      </xsl:if>
      <xsl:if test="preceding::w:background">
        <xsl:apply-templates select="preceding::w:background" mode="fillColor"/>
      </xsl:if>
      <xsl:variable name="headerType" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'"/>
      <xsl:if test="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Type=$headerType]">
        <xsl:call-template name="header"/>
      </xsl:if>
      <xsl:variable name="footerType" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'"/>
      <xsl:if test="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Type=$footerType]">
        <xsl:call-template name="footer"/>
      </xsl:if>
      <xsl:if test="document('word/settings.xml')/w:settings/w:evenAndOddHeaders">
        <xsl:apply-templates select="document('word/settings.xml')/w:settings/w:evenAndOddHeaders" mode="Header"/>
      </xsl:if>
      <xsl:if test="document('word/settings.xml')/w:settings/w:mirrorMargins">
        <xsl:apply-templates select="document('word/settings.xml')/w:settings/w:mirrorMargins" mode="sectPr"/>
      </xsl:if>
    </字:节属性_421B>    
  </xsl:template>
  
  <xsl:template name ="revisionSectProperties">
	  <字:修订开始_421F>
      <xsl:variable name="vdel">
        <xsl:number count="w:del|w:ins|w:rPrChange|w:tblPrChange|w:tblGridChange|w:tcPrChange|w:sectPrChange" level="any" format="1"/>
      </xsl:variable>
      <xsl:attribute name="标识符_4220">
        <xsl:value-of select="concat('xd',$vdel)"/>
      </xsl:attribute>
      <xsl:attribute name="类型_4221">
        <xsl:value-of select="'format'"/>
      </xsl:attribute>
      <xsl:attribute name="修订信息引用_4222">
        <xsl:value-of select="concat('rev_',$vdel)"/>
      </xsl:attribute>
    </字:修订开始_421F>
    <xsl:call-template name="sectProperties"/>
	  <字:修订结束_4223>
      <xsl:variable name="vdel">
        <xsl:number count="w:del|w:ins|w:rPrChange|w:tblPrChange|w:tblGridChange|w:tcPrChange|w:sectPrChange" level="any" format="1"/>
      </xsl:variable>
      <xsl:attribute name="开始标识引用_4224">
        <xsl:value-of select="concat('xd',$vdel)"/>
      </xsl:attribute>
    </字:修订结束_4223>
  </xsl:template>

  <xsl:template match ="w:settings" mode ="gutter">
    <xsl:if test ="w:gutterAtTop[not(@w:val)] or w:gutterAtTop[@w:val='on'] or w:gutterAtTop[@w:val='1'] or w:gutterAtTop[@w:val='true']">
      <xsl:value-of select="'top'"/>
    </xsl:if>
    <xsl:if test ="w:gutterAtTop[@w:val='off'] or w:gutterAtTop[@w:val='0'] or w:gutterAtTop[@w:val='false'] or not(w:gutterAtTop)">
      <xsl:value-of select="'left'"/>
    </xsl:if>
  </xsl:template>
  <!--cxl,2012.3.5修改等宽分栏宽度、间距模板，默认设置下OOX代码中不显示该属性值-->
  <xsl:template name ="getCol">
    <xsl:param name ="num"/>
    <xsl:if test ="$num > 0">
      <xsl:variable name="space">
        <xsl:if test="w:cols/@w:space">
          <xsl:value-of select="(./w:cols/@w:space) div 20"/>
        </xsl:if>
        <xsl:if test="not(w:cols/@w:space)">
          <xsl:value-of select="number(21.21)"/>
        </xsl:if>
      </xsl:variable>
      <xsl:variable name="width">
        <xsl:if test="w:cols/@w:w">
          <xsl:value-of select="(./w:cols/@w:w) div 20"/>
        </xsl:if>
        <xsl:if test="not(w:cols/@w:w)">
          <xsl:value-of select="(39.55 * 10.5 - ($num - 1) * $space) div $num"/>
        </xsl:if>
      </xsl:variable>
		  <字:栏_41E0>
				<xsl:attribute name="宽度_41E1">
					<xsl:value-of select="$width"/>
				</xsl:attribute>
				<xsl:attribute name="间距_41E2">
					<xsl:value-of select="$space"/>
				</xsl:attribute>
		  </字:栏_41E0>
      <xsl:call-template name ="getCol">
        <xsl:with-param name ="num" select ="$num - 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:evenAndOddHeaders" mode="Header">
    <xsl:if test="not(@w:val) or @w:val='true' or @w:val='1'">
		<字:是否奇偶页页眉页脚不同_41ED>true</字:是否奇偶页页眉页脚不同_41ED>
    </xsl:if>
    <xsl:if test="@w:val='false' or @w:val='0'">
		<字:是否奇偶页页眉页脚不同_41ED>false</字:是否奇偶页页眉页脚不同_41ED>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:mirrorMargins" mode="sectPr">
    <xsl:if test="not(@w:val) or @w:val='true' or @w:val='1'">
		<字:是否对称页边距_41FD>true</字:是否对称页边距_41FD>
    </xsl:if>
    <xsl:if test="@w:val='false' or @w:val='0'">
		<字:是否对称页边距_41FD>false</字:是否对称页边距_41FD>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:background" mode="fillColor">
	  <字:填充_4134 是否填充随图形旋转_8067="false">
      <!--yx,change below,2010.2.28-->
      <xsl:if test="not(v:background) and @w:color!='auto'">
        <图:颜色_8004>
          <xsl:value-of select="concat('#',@w:color)"/>
        </图:颜色_8004>
      </xsl:if>
      <!--auto的情况，没找到实例-->
      <xsl:if test="@w:color='auto'">
		    <图:颜色_8004><!--这个方式不正确，至少该有个#
          <xsl:value-of select="@w:color"/>-->
          <xsl:value-of select ="'#ffffff'"/>
        </图:颜色_8004>
      </xsl:if>
      
      <!--yx,change above,2010.2.28-->
      <xsl:if test="v:background">
        <xsl:apply-templates select="v:background" mode="background"/>
      </xsl:if>
    </字:填充_4134>
  </xsl:template>
  
  <!--2015-04-26，wudi，修复图片填充BUG，start-->
  <xsl:template match="v:background" mode="background">
    <xsl:variable name="filltype" select="v:fill/@type"/>
    <xsl:choose><!--渐变填充-->
      
      <!--2014-01-09，wudi，调整代码，增加对中心辐射，角部辐射的处理，start-->
      <xsl:when test="$filltype='gradient' or $filltype='gradientRadial'">
        <图:渐变_800D 起始浓度_8011="100.0" 终止浓度_8012="100.0">
          <xsl:if test ="$filltype='gradient' or not($filltype='gradientRadial' and v:fill/@focusposition='.5,.5' and not(v:fill/@focus)) or (v:fill/@type='gradient' and (v:fill/@angle='-45') and not(v:fill/@focus)) or (v:fill/@type='gradient' and not(v:fill/@angle='-45') and (v:fill/@focus='100%')) or (v:fill/@type='gradient' and (v:fill/@angle='-90' or v:fill/@angle='-135' or v:fill/@angle='-45') and v:fill/@focus='-50%') or (v:fill/@type='gradient' and not(v:fill/@angle) and v:fill/@focus='50%')">
            <xsl:attribute name ="起始色_800E">
              <xsl:variable name ="color1">
                <xsl:value-of select ="@fillcolor"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test ="$color1='red'">
                  <xsl:value-of select ="'#ff0000'"/>
                </xsl:when>
                <xsl:when test ="$color1='yellow'">
                  <xsl:value-of select ="'#ffff00'"/>
                </xsl:when>
                <xsl:when test ="$color1='black'">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="$color1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="终止色_800F">
              <xsl:variable name ="color2">
                <xsl:value-of select ="v:fill/@color2"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test ="$color2='red'">
                  <xsl:value-of select ="'#ff0000'"/>
                </xsl:when>
                <xsl:when test ="$color2='yellow'">
                  <xsl:value-of select ="'#ffff00'"/>
                </xsl:when>
                <xsl:when test ="$color2='black'">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <xsl:when test ="contains($color2,'darken')">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <xsl:when test ="contains($color2,'lighten')">
                  <xsl:value-of select ="'#ffffff'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="$color2"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="($filltype='gradientRadial' and v:fill/@focusposition='.5,.5' and not(v:fill/@focus)) or (v:fill/@type='gradient' and (v:fill/@angle='-45') and (v:fill/@focus='100%')) or (v:fill/@type='gradient' and not(v:fill/@angle='-45') and not(v:fill/@focus)) or (v:fill/@type='gradient' and (v:fill/@angle='-90' or v:fill/@angle='-135' or v:fill/@angle='-45') and v:fill/@focus='50%') or (v:fill/@type='gradient' and not(v:fill/@angle) and v:fill/@focus='-50%')">
            <xsl:attribute name ="起始色_800E">
              <xsl:variable name ="color2">
                <xsl:value-of select ="v:fill/@color2"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test ="$color2='red'">
                  <xsl:value-of select ="'#ff0000'"/>
                </xsl:when>
                <xsl:when test ="$color2='yellow'">
                  <xsl:value-of select ="'#ffff00'"/>
                </xsl:when>
                <xsl:when test ="$color2='black'">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <xsl:when test ="contains($color2,'darken')">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <xsl:when test ="contains($color2,'lighten')">
                  <xsl:value-of select ="'#ffffff'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="$color2"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name ="终止色_800F">
              <xsl:variable name ="color1">
                <xsl:value-of select ="@fillcolor"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test ="$color1='red'">
                  <xsl:value-of select ="'#ff0000'"/>
                </xsl:when>
                <xsl:when test ="$color1='yellow'">
                  <xsl:value-of select ="'#ffff00'"/>
                </xsl:when>
                <xsl:when test ="$color1='black'">
                  <xsl:value-of select ="'#000000'"/>
                </xsl:when>
                <!-- <xsl:when test ="contains($color1,'fill')">
                <xsl:value-of select =""/> 
              </xsl:when>-->
                <xsl:otherwise>
                  <xsl:value-of select ="$color1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:attribute name ="种子类型_8010">
            <xsl:if test ="v:fill/@type='gradient' and v:fill/@focus='100%'">
              <xsl:value-of select ="'linear'"/>
            </xsl:if>
            <xsl:if test ="v:fill/@type='gradient' and not(v:fill/@focus)">
              <xsl:value-of select ="'linear'"/>
            </xsl:if>
            <xsl:if test ="v:fill/@type='gradient' and v:fill/@focus='-50%'">
              <xsl:value-of select ="'axial'"/>
            </xsl:if>
            <xsl:if test ="v:fill/@type='gradient' and v:fill/@focus='50%'">
              <xsl:value-of select ="'axial'"/>
            </xsl:if>
            <!--2014-01-09，wudi，修正，原先取值是square，start-->
            <xsl:if test ="v:fill/@type='gradientRadial'">
              <xsl:value-of select ="'rectangle'"/>
            </xsl:if>
            <!--end-->

          </xsl:attribute>
          <xsl:attribute name ="渐变方向_8013">
            <xsl:choose>
              <xsl:when test ="v:fill/@angle='-90'">
                <xsl:value-of select ="'90'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@angle='-135'">
                <xsl:value-of select ="'45'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@angle='-45'">
                <xsl:value-of select ="'315'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and v:fill/@focusposition='.5,.5' and not(v:fill/@focus)">
                <xsl:value-of select ="'315'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'0'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name ="边界_8014">
            <xsl:if test ="v:fill/@type='gradient'">
              <xsl:choose>
                <xsl:when test ="v:fill/@angle='-45' and not(v:fill/@focus)">
                  <xsl:value-of select ="'80'"/>
                </xsl:when>
                <xsl:when test ="v:fill/@angle='-135' and (v:fill/@focus='100%')">
                  <xsl:value-of select ="'80'"/>
                </xsl:when>
                <xsl:when test ="not(v:fill/@focus)">
                  <xsl:value-of select ="'100'"/>
                </xsl:when>
                <xsl:when test ="v:fill/@focus='-50%'">
                  <xsl:value-of select ="'100'"/>
                </xsl:when>
                <xsl:when test ="v:fill/@focus='50%'">
                  <xsl:value-of select ="'100'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'50'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <xsl:if test ="v:fill/@type='gradientRadial'">
              <xsl:value-of select ="'100'"/>
            </xsl:if>
          </xsl:attribute>
          <xsl:attribute name ="种子X位置_8015">
            <xsl:choose>
              <xsl:when test ="v:fill/o:fill and not(v:fill/@type='gradientRadial')">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and not(v:fill/@focusposition)">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='1')">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition=',1')">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='1,1')">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='.5,.5')">
                <xsl:value-of select ="'50'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and v:fill/@angle='-45' and v:fill/@focus='100%'">
                <xsl:value-of select ="'17'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and v:fill/@angle='-45'  and not(v:fill/@focus)">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and not(v:fill/@angle='-45') and v:fill/@focus='100%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and not(v:fill/@angle='-45')  and not(v:fill/@focus)">
                <xsl:value-of select ="'17'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and (v:fill/@angle='-90' or v:fill/@angle='-135' or v:fill/@angle='-45') and v:fill/@focus='-50%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and (v:fill/@angle='-90' or v:fill/@angle='-135' or v:fill/@angle='-45') and v:fill/@focus='50%'">
                <xsl:value-of select ="'17'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and not(v:fill/@angle) and v:fill/@focus='50%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and not(v:fill/@angle) and v:fill/@focus='-50%'">
                <xsl:value-of select ="'17'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'100'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name ="种子Y位置_8016">
            <xsl:choose>
              <xsl:when test ="v:fill/o:fill and not(v:fill/@type='gradientRadial')">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and not(v:fill/@focusposition)">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='1')">
                <xsl:value-of select ="'0'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition=',1')">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='1,1')">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test="v:fill/@type='gradientRadial' and (v:fill/@focusposition='.5,.5')">
                <xsl:value-of select ="'50'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and v:fill/@focus='100%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and not(v:fill/@focus)">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and v:fill/@focus='-50%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:when test ="v:fill/@type='gradient' and v:fill/@focus='50%'">
                <xsl:value-of select ="'100'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'100'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </图:渐变_800D>
      </xsl:when>
      <!--end-->
      
      <!--纹理和图片一样-->
      <!--<xsl:when test="$filltype='tile'">
			  <图:图片_8005>
			    <xsl:attribute name="位置_8006">
				    <xsl:choose>
					    <xsl:when test="a:stretch">
						    <xsl:value-of select="'stretch'"/>
					    </xsl:when>
					    <xsl:when test="a:tile">
						    <xsl:value-of select="'tile'"/>
					    </xsl:when>
					    <xsl:when test="a:srcRect">tile</xsl:when>
					    <xsl:otherwise>
						    <xsl:value-of select="'tile'"/>
					    </xsl:otherwise>
				    </xsl:choose>
			    </xsl:attribute>
			    <xsl:attribute name="图形引用_8007"> 
				    <xsl:value-of select="concat('document','Obj',1)"/>
			    </xsl:attribute>
			    <xsl:attribute name ="名称_8009">
				    <xsl:value-of select ="v:fill/@o:title"/>
			    </xsl:attribute>
		    </图:图片_8005>
      </xsl:when>-->
      <!-- 图案通过名称对应,为实现互操作，将图片也转换-->
      <xsl:when test="$filltype='pattern'">
        <图:图案_800A>
          <xsl:attribute name ="类型_8008">
            <xsl:variable name ="title">
              <xsl:value-of select ="v:fill/@o:title"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$title ='5%'">ptn001</xsl:when>
              <xsl:when test="$title ='10%'">ptn002</xsl:when>
              <xsl:when test="$title ='20%'">ptn003</xsl:when>
              <xsl:when test="$title ='25%'">ptn004</xsl:when>
              <xsl:when test="$title ='30%'">ptn005</xsl:when>
              <xsl:when test="$title ='40%'">ptn006</xsl:when>
              <xsl:when test="$title ='50%'">ptn007</xsl:when>
              <xsl:when test="$title ='60%'">ptn008</xsl:when>
              <xsl:when test="$title ='70%'">ptn009</xsl:when>
              <xsl:when test="$title ='75%'">ptn010</xsl:when>
              <xsl:when test="$title ='80%'">ptn011</xsl:when>
              <xsl:when test="$title ='90%'">ptn012</xsl:when>
              <xsl:when test="$title ='浅色下对角线'">ptn013</xsl:when>
              <xsl:when test="$title ='浅色上对角线'">ptn014</xsl:when>
              <xsl:when test="$title ='深色下对角线'">ptn015</xsl:when>
              <xsl:when test="$title ='深色上对角线'">ptn016</xsl:when>
              <xsl:when test="$title ='宽下对角线'">ptn017</xsl:when>
              <xsl:when test="$title ='宽上对角线'">ptn018</xsl:when>
              <xsl:when test="$title ='浅色竖线'">ptn019</xsl:when>
              <xsl:when test="$title ='浅色横线'">ptn020</xsl:when>
              <xsl:when test="$title ='窄竖线'">ptn021</xsl:when>
              <xsl:when test="$title ='窄横线'">ptn022</xsl:when>
              <xsl:when test="$title ='深色竖线'">ptn023</xsl:when>
              <xsl:when test="$title ='深色横线'">ptn024</xsl:when>
              <xsl:when test="$title ='下对角虚线'">ptn025</xsl:when>
              <xsl:when test="$title ='上对角虚线'">ptn026</xsl:when>
              <xsl:when test="$title ='横虚线'">ptn027</xsl:when>
              <xsl:when test="$title ='竖虚线'">ptn028</xsl:when>
              <xsl:when test="$title ='小纸屑'">ptn029</xsl:when>
              <xsl:when test="$title ='大纸屑'">ptn030</xsl:when>
              <xsl:when test="$title ='之字形'">ptn031</xsl:when>
              <xsl:when test="$title ='波浪线'">ptn032</xsl:when>
              <xsl:when test="$title ='对角砖形'">ptn033</xsl:when>
              <xsl:when test="$title ='横向砖形'">ptn034</xsl:when>
              <xsl:when test="$title ='编织物'">ptn035</xsl:when>
              <xsl:when test="$title ='苏格兰方格呢'">ptn036</xsl:when>
              <xsl:when test="$title ='草皮'">ptn037</xsl:when>
              <xsl:when test="$title ='虚线网格'">ptn038</xsl:when>
              <xsl:when test="$title ='点式菱形'">ptn039</xsl:when>
              <xsl:when test="$title ='瓦形'">ptn040</xsl:when>
              <xsl:when test="$title ='棚架'">ptn041</xsl:when>
              <xsl:when test="$title ='球体'">ptn042</xsl:when>
              <xsl:when test="$title ='小网格'">ptn043</xsl:when>
              <xsl:when test="$title ='大网格'">ptn044</xsl:when>
              <xsl:when test="$title ='小棋盘'">ptn045</xsl:when>
              <xsl:when test="$title ='大棋盘'">ptn046</xsl:when>
              <xsl:when test="$title ='轮廓式菱形'">ptn047</xsl:when>
              <xsl:when test="$title ='实心菱形'">ptn048</xsl:when>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name ="前景色_800B">
            <xsl:choose>
              <xsl:when test ="@fillcolor='black'">
                <xsl:value-of select ="'#000000'"/>
              </xsl:when>
              <xsl:when test ="@fillcolor='red'">
                <xsl:value-of select ="'#ff0000'"/>
              </xsl:when>
              <xsl:when test ="@fillcolor='yellow'">
                <xsl:value-of select ="'#ffff00'"/>
              </xsl:when>
              <xsl:when test ="contains(@fillcolor,'#')">
                <xsl:value-of select ="@fillcolor"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'#ffffff'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="背景色_800C">
            <xsl:variable name ="backcolor">
              <xsl:value-of select ="v:fill/@color2"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test ="$backcolor='black'">
                <xsl:value-of select ="'#000000'"/>
              </xsl:when>
              <xsl:when test ="$backcolor='red'">
                <xsl:value-of select ="'#ff0000'"/>
              </xsl:when>
              <xsl:when test ="$backcolor='yellow'">
                <xsl:value-of select ="'#ffff00'"/>
              </xsl:when>
              <xsl:when test ="contains($backcolor,'#')">
                <xsl:value-of select ="$backcolor"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'#ffffff'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </图:图案_800A>
      </xsl:when>
      <!--图片-->
      <xsl:when test="$filltype='frame' or $filltype='tile'">
		    <图:图片_8005>
			    <xsl:attribute name="位置_8006">
				    <xsl:choose>
					    <xsl:when test="a:stretch">
						    <xsl:value-of select="'stretch'"/>
					    </xsl:when>
              <!--<xsl:when test="a:tile">
                <xsl:value-of select="'tile'"/>
              </xsl:when>-->
              <xsl:when test="a:srcRect">
                <xsl:value-of select="'srcRect'"/>
              </xsl:when>
					    <xsl:otherwise>
						    <xsl:value-of select="'tile'"/>
					    </xsl:otherwise>
				    </xsl:choose>
			    </xsl:attribute>
			    <xsl:attribute name="图形引用_8007"> 
				    <xsl:value-of select="concat('document','Obj',1)"/>
			    </xsl:attribute>
			    <xsl:attribute name ="名称_8009">
				    <xsl:value-of select ="v:fill/@o:title"/>
			    </xsl:attribute>
		    </图:图片_8005>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--end-->
  
  <xsl:template name="lnNumTypeRestart">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='newPage'">
        <xsl:value-of select="'page'"/>
      </xsl:when>
      <xsl:when test="$val='newSection'">
        <xsl:value-of select="'section'"/>
      </xsl:when>
      <xsl:when test="$val='continuous'">
        <xsl:value-of select="'continuous'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'continuous'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="notePrNumFmt">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='decimal'">
        <xsl:value-of select="'decimal'"/>
      </xsl:when>
      
      <!--2013-05-03，wudi，修复OOX到UOF方向页码设置转换BUG，遗漏了一种页码格式，start-->
      <xsl:when test ="$val='numberInDash'">
        <xsl:value-of select ="'decimal-in-dash'"/>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="$val='upperRoman'">
        <xsl:value-of select="'upper-roman'"/>
      </xsl:when>
      <xsl:when test="$val='lowerRoman'">
        <xsl:value-of select="'lower-roman'"/>
      </xsl:when>
      <xsl:when test="$val='upperLetter'">
        <xsl:value-of select="'upper-letter'"/>
      </xsl:when>
      <xsl:when test="$val='lowerLetter'">
        <xsl:value-of select="'lower-letter'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal'">
        <xsl:value-of select="'ordinal'"/>
      </xsl:when>
      <xsl:when test="$val='cardinalText'">
        <xsl:value-of select="'cardinal-text'"/>
      </xsl:when>
      <xsl:when test="$val='ordinalText'">
        <xsl:value-of select="'ordinal-text'"/>
      </xsl:when>
      <xsl:when test="$val='hex'">
        <xsl:value-of select="'hex'"/>
      </xsl:when>
      <xsl:when test="$val='decimalFullWidth'">
        <xsl:value-of select="'decimal-full-width'"/>
      </xsl:when>
      <xsl:when test="$val='decimalHalfWidth'">
        <xsl:value-of select="'decimal-half-width'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedCircle'">
        <xsl:value-of select="'decimal-enclosed-circle'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedFullstop'">
        <xsl:value-of select="'decimal-enclosed-fullstop'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedParen'">
        <xsl:value-of select="'decimal-enclosed-paren'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedCircleChinese'">
        <xsl:value-of select="'decimal-enclosed-circle-chinese'"/>
      </xsl:when>
      <xsl:when test="$val='ideographEnclosedCircle'">
        <xsl:value-of select="'ideograph-enclosed-circle'"/>
      </xsl:when>
      <xsl:when test="$val='ideographTraditional'">
        <xsl:value-of select="'ideograph-traditional'"/>
      </xsl:when>
      <xsl:when test="$val='ideographZodiac'">
        <xsl:value-of select="'ideograph-zodiac'"/>
      </xsl:when>
      <xsl:when test="$val='chineseCountingThousand'">
        <xsl:value-of select="'chinese-counting'"/>
      </xsl:when>
      <xsl:when test="$val='chineseLegalSimplified'">
        <xsl:value-of select="'chinese-legal-simplified'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'decimal'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="w:pgBorders" mode="section">
	  <字:边框_4133>
      <xsl:if test ="not(./@w:offsetFrom) or ./@w:offsetFrom!='page'">

        <!--2013-03-19，wudi，修复#2732集成测试OOX到UOF2.0第二页多了边框的BUG，start-->
        <xsl:if test="./w:top/@w:shadow='1'">
          <xsl:attribute name="阴影类型_C645">
            <xsl:value-of select="'right-bottom'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="./@w:display='firstPage'">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'first'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="./@w:display='notfirstPage'">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'except-first'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(./@w:display)">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'all'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="度量依据_4230">
          <xsl:value-of select ="'text'"/>
        </xsl:attribute>
        <!--end-->
        
        <xsl:for-each select="node()">          
          <xsl:variable name="bordername" select="name(.)"/>
          <xsl:call-template name="BorderTemplate">
            <xsl:with-param name ="offset" select ="'text'"/>
            <xsl:with-param name="bname" select="$bordername"/>
          </xsl:call-template>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test ="./@w:offsetFrom='page'"><!--页面边框-->
        
        <!--2013-03-19，wudi，修复#2732集成测试OOX到UOF2.0第二页多了边框的BUG，start-->
        <xsl:if test="./w:top/@w:shadow='1'">
          <xsl:attribute name="阴影类型_C645">
            <xsl:value-of select="'right-bottom'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="./@w:display='firstPage'">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'first'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="./@w:display='notfirstPage'">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'except-first'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(./@w:display)">
          <xsl:attribute name ="应用范围_4229">
            <xsl:value-of select ="'all'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="度量依据_4230">
          <xsl:value-of select ="'page-edge'"/>
        </xsl:attribute>
        <!--end-->
       
        <xsl:for-each select="./node()">
          <xsl:variable name="bordername" select="name(.)"/><!--cxl,2012.3.22-->
          <xsl:call-template name="PageBorder">
            <xsl:with-param name="offset" select="'page'"/>
          </xsl:call-template>
          <!--<xsl:call-template name="BorderTemplate">
            <xsl:with-param name ="offset" select ="'page'"/>
            <xsl:with-param name="bname" select="$bordername"/>
          </xsl:call-template>-->
        </xsl:for-each>
      </xsl:if>
    </字:边框_4133>
  </xsl:template>
  <xsl:template name="PageBorder">
    <xsl:param name="offset"/>
    <xsl:if test="name(.)='w:bottom'">
      <uof:下_C616>
        <xsl:if test="./@w:val">
          <xsl:call-template name="pageLineType">
            <xsl:with-param name="pageLineType" select="@w:val"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="./@w:sz">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="./@w:sz div 8"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@w:space">
          <xsl:if test ="$offset='text'">
            <xsl:variable name ="pgMargin">
              <xsl:call-template name ="getPgMargin"/>
            </xsl:variable>
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="($pgMargin div 20) - @w:space"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="$offset='page'">
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="./@w:space"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>

        <xsl:if test="./@w:color">
          <xsl:attribute name="颜色_C611">
            <xsl:if test ="@w:color!='auto'">
              <xsl:value-of select="concat('#',./@w:color)"/>
            </xsl:if>
            <xsl:if test ="@w:color='auto'">
              <xsl:value-of select ="'auto'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </uof:下_C616>
    </xsl:if>
    <xsl:if test="name(.)='w:top'">
      <uof:上_C614>
        <xsl:if test="./@w:val">
          <xsl:call-template name="pageLineType">
            <xsl:with-param name="pageLineType" select="@w:val"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="./@w:sz">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="./@w:sz div 8"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@w:space">
          <xsl:if test ="$offset='text'">
            <xsl:variable name ="pgMargin">
              <xsl:call-template name ="getPgMargin"/>
            </xsl:variable>
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="($pgMargin div 20) - @w:space"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="$offset='page'">
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="./@w:space"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>

        <xsl:if test="./@w:color">
          <xsl:attribute name="颜色_C611">
            <xsl:if test ="@w:color!='auto'">
              <xsl:value-of select="concat('#',./@w:color)"/>
            </xsl:if>
            <xsl:if test ="@w:color='auto'">
              <xsl:value-of select ="'auto'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </uof:上_C614>
    </xsl:if>
    <xsl:if test="name(.)='w:left'">
      <uof:左_C613>
        <xsl:if test="./@w:val">
          <xsl:call-template name="pageLineType">
            <xsl:with-param name="pageLineType" select="@w:val"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="./@w:sz">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="./@w:sz div 8"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@w:space">
          <xsl:if test ="$offset='text'">
            <xsl:variable name ="pgMargin">
              <xsl:call-template name ="getPgMargin"/>
            </xsl:variable>
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="($pgMargin div 20) - @w:space"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="$offset='page'">
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="./@w:space"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>

        <xsl:if test="./@w:color">
          <xsl:attribute name="颜色_C611">
            <xsl:if test ="@w:color!='auto'">
              <xsl:value-of select="concat('#',./@w:color)"/>
            </xsl:if>
            <xsl:if test ="@w:color='auto'">
              <xsl:value-of select ="'auto'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </uof:左_C613>
    </xsl:if>
    <xsl:if test="name(.)='w:right'">
      <uof:右_C615>
        <xsl:if test="./@w:val">
          <xsl:call-template name="pageLineType">
            <xsl:with-param name="pageLineType" select="@w:val"/>
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="./@w:sz">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="./@w:sz div 8"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@w:space">
          <xsl:if test ="$offset='text'">
            <xsl:variable name ="pgMargin">
              <xsl:call-template name ="getPgMargin"/>
            </xsl:variable>
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="($pgMargin div 20) - @w:space"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="$offset='page'">
            <xsl:attribute name="边距_C610">
              <xsl:value-of select="./@w:space"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>

        <xsl:if test="./@w:color">
          <xsl:attribute name="颜色_C611">
            <xsl:if test ="@w:color!='auto'">
              <xsl:value-of select="concat('#',./@w:color)"/>
            </xsl:if>
            <xsl:if test ="@w:color='auto'">
              <xsl:value-of select ="'auto'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </uof:右_C615>
    </xsl:if>
  </xsl:template>
  <xsl:template name="BorderTemplate">
    <xsl:param name="offset"/>
    <xsl:param name="bname"/>
    <xsl:variable name="uofbname">
      <xsl:call-template name="convertbname">
        <xsl:with-param name="ooxname" select="$bname"/>
      </xsl:call-template>
    </xsl:variable>
    <!--cxl2011/11/27删除，现在的UOF标准里没有这个ID-->
    <!--<xsl:variable name="uofbid">
      <xsl:call-template name="convertbid">
        <xsl:with-param name="ooxname" select="$bname"/>
      </xsl:call-template>
    </xsl:variable>-->
    <xsl:element name="{$uofbname}">
      
      <!--yx,change 类型 to 线型，2010.3.27-->
      <xsl:if test="./@w:val">
        <xsl:call-template name="pageLineType">
          <xsl:with-param name="pageLineType" select="@w:val"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="./@w:sz">
        <xsl:attribute name="宽度_C60F">
          <xsl:value-of select="./@w:sz div 8"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./@w:space">
        <xsl:if test ="$offset='text'">
          <xsl:variable name ="pgMargin">
            <xsl:call-template name ="getPgMargin"/>
          </xsl:variable>
          <xsl:attribute name="边距_C610">
            <xsl:value-of select="($pgMargin div 20) - @w:space"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="$offset='page'">
          <xsl:attribute name="边距_C610">
            <xsl:value-of select="./@w:space"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>

      <xsl:if test="./@w:color">
        <xsl:attribute name="颜色_C611">
          <xsl:if test ="@w:color!='auto'">
             <xsl:value-of select="concat('#',./@w:color)"/>
          </xsl:if>
          <xsl:if test ="@w:color='auto'">
            <xsl:value-of select ="'auto'"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
    </xsl:element>
  </xsl:template>
  
  <xsl:template name="convertbname">
    <xsl:param name="ooxname"/>
    <xsl:choose>
      <xsl:when test="$ooxname='w:bottom'">
        <xsl:value-of select="'uof:下_C616'"/>
      </xsl:when>
      <xsl:when test="$ooxname='w:right'">
        <xsl:value-of select="'uof:右_C615'"/>
      </xsl:when>
      <xsl:when test="$ooxname='w:top'">
        <xsl:value-of select="'uof:上_C614'"/>
      </xsl:when>
      <xsl:when test="$ooxname='w:left'">
        <xsl:value-of select="'uof:左_C613'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="getPgMargin">
    <xsl:choose>
      <xsl:when test ="name(.)='w:top'">
        <xsl:value-of select ="preceding::w:pgMar[1]/@w:top"/>
      </xsl:when>
      <xsl:when test ="name(.)='w:left'">
        <xsl:value-of select ="preceding::w:pgMar[1]/@w:left"/>
      </xsl:when>
      <xsl:when test ="name(.)='w:bottom'">
        <xsl:value-of select ="preceding::w:pgMar[1]/@w:bottom"/>
      </xsl:when>
      <xsl:when test ="name(.)='w:right'">
        <xsl:value-of select ="preceding::w:pgMar[1]/@w:right"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>


  <xsl:template name="pageLineType">
    <xsl:param name="pageLineType" />
    <xsl:choose>
      <xsl:when test="$pageLineType='single'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='dotted'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">square-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='dashSmallGap'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='dashed'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='dotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='dotDotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='double'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='triple'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickSmallGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinSmallGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinThickSmallGap'">
        <xsl:attribute name="线型_C60D">thin-between-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickThinSmallGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickMediumGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinMediumGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickThinMediumGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinThickMediumGap'">
        <xsl:attribute name="线型_C60D">thin-between-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickLargeGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thinThickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='thickThinThickLargeGap'">
        <xsl:attribute name="线型_C60D">thin-between-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      
      <!--2013-03-28，wudi，修复页面边框的BUG，start-->
      <xsl:when test="$pageLineType='wave'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='doubleWave'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="$pageLineType='dashDotStroked'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='threeDEmboss'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='threeDEngrave'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='outset'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$pageLineType='inset'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>      
      <xsl:otherwise>
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:otherwise>      
    </xsl:choose>
  </xsl:template> 
  
  
  <xsl:template name="textDirection">
    <xsl:param name="dir"/>
    <xsl:choose>
      <xsl:when test="$dir='lrTb'">
        <xsl:value-of select="'hori-l2r'"/>
      </xsl:when>
      <!--yx,修改09.12.15-->
      <!--xsl:when test="$dir='tbRl'">
        <xsl:value-of select="'hori-r2l'"/>
      </xsl:when-->
      <xsl:when test="$dir='tbRl'">
        <xsl:value-of select="'r2l-t2b-0e-90w'"/>
      </xsl:when>
      <xsl:when test="$dir='lrTbV'"><!--cxl,将中文悬转270-->
        <xsl:value-of select="'t2b-l2r-270e-0w'"/>
      </xsl:when>
      <!--yx,因为和上面的部分代码重复了，所以去掉了09.12.15-->
      <!--xsl:when test="$dir='tbRl'">
        <xsl:value-of select="'vert-r2l'"/>
      </xsl:when-->
      <!--杨晓，将hori-l2r改为l2r-t2b-0e-0w09.12.18-->
      <xsl:otherwise>
        <xsl:value-of select="'l2r-t2b-0e-0w'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name ="pgNumChapSep">
    <xsl:param name ="val"/>
    <xsl:choose>
      <xsl:when test ="$val='emDash'">
        <xsl:value-of select ="'em-dash'"/>
      </xsl:when>
      <xsl:when test ="$val='enDash'">
        <xsl:value-of select ="'en-dash'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="$val"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
