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
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"                
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special">
  <xsl:import href="document.xsl"/>
  <xsl:import href="package_relationships.xsl"/>
  <xsl:import href="contentTypes.xsl"/>
  <xsl:import href="paragraph.xsl"/>
  <xsl:import href="table.xsl"/>
  <xsl:import href="sectPr.xsl"/>
  <xsl:import href="metadata.xsl"/>
  <xsl:import href="settings.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="footnote-endnote.xsl"/>
  <xsl:import href="header-footer.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:import href="object.xsl"/>
  <xsl:import href="wordpart_relationships.xsl"/>
  <xsl:import href="numbering.xsl"/>
  <xsl:import href="region.xsl"/>
  <xsl:import href="hyperlinks.xsl"/>
  <xsl:import href="run.xsl"/>
  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>

  <xsl:template match="/uof:UOF">
    <pzip:archive pzip:target="{$outputFile}">

      <xsl:if test="//w:media">
        <xsl:copy-of select="//w:media/*"/>
      </xsl:if>
      <!-- content types -->
      <pzip:entry pzip:target="[Content_Types].xml">
        <xsl:call-template name="contentTypes"/>
      </pzip:entry>
      <!-- package relationship item -->
      <pzip:entry pzip:target="_rels/.rels">
        <xsl:call-template name="package-relationships"/>
      </pzip:entry>
      <!--metadata-->
      <xsl:if test="//元:元数据_5200">
        <pzip:entry pzip:target="docProps/app.xml">
          <xsl:apply-templates select=".//元:元数据_5200" mode="zip1"/>
        </pzip:entry>
        <pzip:entry pzip:target="docProps/core.xml">
          <xsl:apply-templates select=".//元:元数据_5200" mode="zip2"/>
        </pzip:entry>
        <xsl:if test=".//元:元数据_5200/元:用户自定义元数据集_520F">
          <pzip:entry pzip:target="docProps/custom.xml">
            <xsl:apply-templates select=".//元:元数据_5200" mode="zip3"/>
          </pzip:entry>
        </xsl:if>
      </xsl:if>
      
      <!-- styles -->
      <pzip:entry pzip:target="word/styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>
      <!--footnote and endnote-->
      <xsl:if test="//字:句_419D/字:脚注_4159">
        <pzip:entry pzip:target="word/footnotes.xml">
          <xsl:call-template name="footnote"/>
        </pzip:entry>
      </xsl:if>
      <xsl:if test="//字:句_419D/字:尾注_415A">
        <pzip:entry pzip:target="word/endnotes.xml">
          <xsl:call-template name="endnote"/>
        </pzip:entry>
      </xsl:if>

      <!-- 页眉和页脚-->
      <xsl:for-each select=".//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:奇数页页眉_41F4
                           |.//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:偶数页页眉_41F5
                           |.//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:首页页眉_41F6">
        <pzip:entry>
          <xsl:variable name="num">
            <xsl:number count="字:奇数页页眉_41F4|字:偶数页页眉_41F5|字:首页页眉_41F6" level="any" format="1"/>
          </xsl:variable>
          <xsl:attribute name="pzip:target">
            <xsl:value-of select="concat('word/header',$num,'.xml')"/>
          </xsl:attribute>
          <xsl:call-template name="header"/>
        </pzip:entry>
      </xsl:for-each>
      <xsl:for-each
        select=".//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:奇数页页脚_41F8
                           |.//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:偶数页页脚_41F9
                           |.//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:首页页脚_41FA">
        <pzip:entry>
          <xsl:attribute name="pzip:target">
            <xsl:variable name="num">
              <xsl:number count="字:奇数页页脚_41F8|字:偶数页页脚_41F9|字:首页页脚_41FA" level="any" format="1"/>
            </xsl:variable>
            <xsl:value-of select="concat('word/footer',$num,'.xml')"/>
          </xsl:attribute>
          <xsl:call-template name="footer"/>
        </pzip:entry>
      </xsl:for-each>

      <!-- fonts declaration -->
      <pzip:entry pzip:target="word/fontTable.xml">
        <xsl:call-template name="fontTable"/>
      </pzip:entry>

      <!--custom xml-->
      <xsl:if test="//扩展:扩展区_B200/w:customItem">
        <xsl:for-each select="//扩展:扩展区_B200/w:customItem">
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="@fileName"/>
            </xsl:attribute>
              <xsl:copy-of select="./*"/>           
          </pzip:entry>
        </xsl:for-each>
      </xsl:if>

      <!--numbering-->
      <xsl:if test="uof:式样集//式样:自动编号集_990E">
        <pzip:entry pzip:target="word/numbering.xml">
          <xsl:apply-templates select="uof:式样集//式样:自动编号集_990E" mode="zip"/>
        </pzip:entry>
      </xsl:if>

      <!--comments-->
      <xsl:if test="uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669">
        <pzip:entry pzip:target="word/comments.xml">
          <xsl:apply-templates select="uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669" mode="zip"/>
        </pzip:entry>
      </xsl:if>
  
      <!-- settings  -->
      <pzip:entry pzip:target="word/settings.xml">
        <xsl:call-template name="settings"/>
      </pzip:entry>
      
      <!--theme/theme1.xml-->
      <pzip:entry pzip:target="word/theme/theme1.xml">
        <xsl:call-template name="theme1"/>
      </pzip:entry>
      
      <!-- webSettings  -->
      <!--
			<pzip:entry pzip:target="word/webSettings.xml">
				<xsl:call-template name="webSettings"/>
			</pzip:entry>-->

      <!-- part relationship item -->
		
		<pzip:entry pzip:target="word/_rels/document.xml.rels">
			<xsl:call-template name="documentXmlRels"/>
		</pzip:entry>
		
		<xsl:if test="/uof:UOF/uof:式样集/式样:自动编号集_990E/字:自动编号_4124/字:级别_4112/字:图片符号_411B">
			<pzip:entry pzip:target="word/_rels/numbering.xml.rels">
				<xsl:call-template name="numberingXmlRels"/>
			</pzip:entry>
		</xsl:if>
      <!--comments part relationship item -->
      <xsl:if
        test="(uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink']) or (uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644)">
        <pzip:entry pzip:target="word/_rels/comments.xml.rels">
          <xsl:call-template name="otherXmlRels">
            <xsl:with-param name="relType">
              <xsl:value-of select="'comments'"/>
            </xsl:with-param>
          </xsl:call-template>
        </pzip:entry>
      </xsl:if>
      <!--headers footers part relationship item -->
      <xsl:for-each select=".//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:奇数页页眉_41F4
                           |.//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:偶数页页眉_41F5
                           |.//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:首页页眉_41F6">

        <xsl:if
          test=".//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink'] 
          or (uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644)
          or (.//uof:锚点_C644[@图形引用_C62E=//uof:对象集/图:图形_8062[图:图片数据引用_8037=//uof:对象集/对象:对象数据_D701/@标识符_D704]/@标识符_804B])">
          <pzip:entry>
            <xsl:variable name="num">
              <xsl:number count="字:奇数页页眉_41F4|字:偶数页页眉_41F5|字:首页页眉_41F6" level="any" format="1"/>
            </xsl:variable>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('word/_rels/header',$num,'.xml.rels')"/>
            </xsl:attribute>

            <xsl:call-template name="otherXmlRels">
              <xsl:with-param name="relType">
                <!--<xsl:value-of select ="'header'"/>-->
                <xsl:value-of select="name(.)"/>
              </xsl:with-param>
            </xsl:call-template>
          </pzip:entry>
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select=".//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:奇数页页脚_41F8
                           |.//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:偶数页页脚_41F9
                           |.//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:首页页脚_41FA">

        <xsl:if
          test=".//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/字:区域开始_4165[@类型_413B='hyperlink'] 
          or (uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669//字:句_419D[not(ancestor::字:脚注_4159 or ancestor::字:尾注_415A)]/uof:锚点_C644)
          or (.//uof:锚点_C644[@图形引用_C62E=//uof:对象集/图:图形_8062[图:图片数据引用_8037=//uof:对象集/对象:对象数据_D701/@标识符_D704]/@标识符_804B])">
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:variable name="num">
                <xsl:number count="字:奇数页页脚_41F8|字:偶数页页脚_41F9|字:首页页脚_41FA" level="any" format="1"/>
              </xsl:variable>
              <xsl:value-of select="concat('word/_rels/footer',$num,'.xml.rels')"/>
            </xsl:attribute>
            <xsl:call-template name="otherXmlRels">
              <xsl:with-param name="relType">
                <!--<xsl:value-of select ="'footer'"/>-->
                <xsl:value-of select="name(.)"/>
              </xsl:with-param>
            </xsl:call-template>
          </pzip:entry>
        </xsl:if>

      </xsl:for-each>
      <!--endnotes footnotes part relationship item -->
      <xsl:if test="//字:脚注_4159//字:区域开始_4165[@类型_413B='hyperlink'] or //字:脚注_4159//uof:锚点_C644">
        <pzip:entry pzip:target="word/_rels/footnotes.xml.rels">
          <xsl:call-template name="otherXmlRels">
            <xsl:with-param name="relType">
              <xsl:value-of select="'footnotes'"/>
            </xsl:with-param>
          </xsl:call-template>
        </pzip:entry>
      </xsl:if>
      <xsl:if test="//字:尾注_415A//字:区域开始_4165[@类型_413B='hyperlink'] or //字:尾注_415A//uof:锚点_C644">
        <pzip:entry pzip:target="word/_rels/endnotes.xml.rels">
          <xsl:call-template name="otherXmlRels">
            <xsl:with-param name="relType">
              <xsl:value-of select="'endnotes'"/>
            </xsl:with-param>
          </xsl:call-template>
        </pzip:entry>
      </xsl:if>

      <!-- main content -->
      <xsl:if test=".//字:文字处理文档_4225">
        <pzip:entry pzip:target="word/document.xml">
          <xsl:apply-templates select=".//字:文字处理文档_4225" mode="zip"/>
        </pzip:entry>
      </xsl:if>

    </pzip:archive>
  </xsl:template>

  <xsl:template match="元:元数据_5200" mode="zip1">
    <xsl:call-template name="metadataApp"/>
  </xsl:template>
  <xsl:template match="元:元数据_5200" mode="zip2">
    <xsl:call-template name="metadataCore"/>
  </xsl:template>
  <xsl:template match="元:元数据_5200" mode="zip3">
    <xsl:call-template name="metadataCustom"/>
  </xsl:template>
  <xsl:template match="规则:批注集_B669" mode="zip">
    <xsl:call-template name="comments"/>
  </xsl:template>
  <xsl:template match="字:文字处理文档_4225" mode="zip">
    <xsl:call-template name="document"/>
  </xsl:template>
  <xsl:template match="式样:自动编号集_990E" mode="zip">
    <xsl:call-template name="numbering"/>
  </xsl:template>
  
  <!--cxl2012.2.8-->
  <xsl:template name="theme1">
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1>
            <a:sysClr val="windowText" lastClr="000000"/>
          </a:dk1>
          <a:lt1>
            <a:sysClr val="window" lastClr="FFFFFF"/>
          </a:lt1>
          <a:dk2>
            <a:srgbClr val="1F497D"/>
          </a:dk2>
          <a:lt2>
            <a:srgbClr val="EEECE1"/>
          </a:lt2>
          <a:accent1>
            <a:srgbClr val="4F81BD"/>
          </a:accent1>
          <a:accent2>
            <a:srgbClr val="C0504D"/>
          </a:accent2>
          <a:accent3>
            <a:srgbClr val="9BBB59"/>
          </a:accent3>
          <a:accent4>
            <a:srgbClr val="8064A2"/>
          </a:accent4>
          <a:accent5>
            <a:srgbClr val="4BACC6"/>
          </a:accent5>
          <a:accent6>
            <a:srgbClr val="F79646"/>
          </a:accent6>
          <a:hlink>
            <a:srgbClr val="0000FF"/>
          </a:hlink>
          <a:folHlink>
            <a:srgbClr val="800080"/>
          </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
          <a:majorFont>
            <a:latin typeface="Cambria"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
            <a:font script="Jpan" typeface="ＭＳ ゴシック"/>
            <a:font script="Hang" typeface="맑은 고딕"/>
            <a:font script="Hans" typeface="宋体"/>
            <a:font script="Hant" typeface="新細明體"/>
            <a:font script="Arab" typeface="Times New Roman"/>
            <a:font script="Hebr" typeface="Times New Roman"/>
            <a:font script="Thai" typeface="Angsana New"/>
            <a:font script="Ethi" typeface="Nyala"/>
            <a:font script="Beng" typeface="Vrinda"/>
            <a:font script="Gujr" typeface="Shruti"/>
            <a:font script="Khmr" typeface="MoolBoran"/>
            <a:font script="Knda" typeface="Tunga"/>
            <a:font script="Guru" typeface="Raavi"/>
            <a:font script="Cans" typeface="Euphemia"/>
            <a:font script="Cher" typeface="Plantagenet Cherokee"/>
            <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
            <a:font script="Tibt" typeface="Microsoft Himalaya"/>
            <a:font script="Thaa" typeface="MV Boli"/>
            <a:font script="Deva" typeface="Mangal"/>
            <a:font script="Telu" typeface="Gautami"/>
            <a:font script="Taml" typeface="Latha"/>
            <a:font script="Syrc" typeface="Estrangelo Edessa"/>
            <a:font script="Orya" typeface="Kalinga"/>
            <a:font script="Mlym" typeface="Kartika"/>
            <a:font script="Laoo" typeface="DokChampa"/>
            <a:font script="Sinh" typeface="Iskoola Pota"/>
            <a:font script="Mong" typeface="Mongolian Baiti"/>
            <a:font script="Viet" typeface="Times New Roman"/>
            <a:font script="Uigh" typeface="Microsoft Uighur"/>
            <a:font script="Geor" typeface="Sylfaen"/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="Calibri"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
            <a:font script="Jpan" typeface="ＭＳ 明朝"/>
            <a:font script="Hang" typeface="맑은 고딕"/>
            <a:font script="Hans" typeface="宋体"/>
            <a:font script="Hant" typeface="新細明體"/>
            <a:font script="Arab" typeface="Arial"/>
            <a:font script="Hebr" typeface="Arial"/>
            <a:font script="Thai" typeface="Cordia New"/>
            <a:font script="Ethi" typeface="Nyala"/>
            <a:font script="Beng" typeface="Vrinda"/>
            <a:font script="Gujr" typeface="Shruti"/>
            <a:font script="Khmr" typeface="DaunPenh"/>
            <a:font script="Knda" typeface="Tunga"/>
            <a:font script="Guru" typeface="Raavi"/>
            <a:font script="Cans" typeface="Euphemia"/>
            <a:font script="Cher" typeface="Plantagenet Cherokee"/>
            <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
            <a:font script="Tibt" typeface="Microsoft Himalaya"/>
            <a:font script="Thaa" typeface="MV Boli"/>
            <a:font script="Deva" typeface="Mangal"/>
            <a:font script="Telu" typeface="Gautami"/>
            <a:font script="Taml" typeface="Latha"/>
            <a:font script="Syrc" typeface="Estrangelo Edessa"/>
            <a:font script="Orya" typeface="Kalinga"/>
            <a:font script="Mlym" typeface="Kartika"/>
            <a:font script="Laoo" typeface="DokChampa"/>
            <a:font script="Sinh" typeface="Iskoola Pota"/>
            <a:font script="Mong" typeface="Mongolian Baiti"/>
            <a:font script="Viet" typeface="Arial"/>
            <a:font script="Uigh" typeface="Microsoft Uighur"/>
            <a:font script="Geor" typeface="Sylfaen"/>
          </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
          <a:fillStyleLst>
            <a:solidFill>
              <a:schemeClr val="phClr"/>
            </a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:tint val="50000"/>
                    <a:satMod val="300000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="35000">
                  <a:schemeClr val="phClr">
                    <a:tint val="37000"/>
                    <a:satMod val="300000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:tint val="15000"/>
                    <a:satMod val="350000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:lin ang="16200000" scaled="1"/>
            </a:gradFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:shade val="51000"/>
                    <a:satMod val="130000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="80000">
                  <a:schemeClr val="phClr">
                    <a:shade val="93000"/>
                    <a:satMod val="130000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:shade val="94000"/>
                    <a:satMod val="135000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:lin ang="16200000" scaled="0"/>
            </a:gradFill>
          </a:fillStyleLst>
          <a:lnStyleLst>
            <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr">
                  <a:shade val="95000"/>
                  <a:satMod val="105000"/>
                </a:schemeClr>
              </a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
            <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
            <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
          </a:lnStyleLst>
          <a:effectStyleLst>
            <a:effectStyle>
              <a:effectLst>
                <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                  <a:srgbClr val="000000">
                    <a:alpha val="38000"/>
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </a:effectStyle>
            <a:effectStyle>
              <a:effectLst>
                <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                  <a:srgbClr val="000000">
                    <a:alpha val="35000"/>
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </a:effectStyle>
            <a:effectStyle>
              <a:effectLst>
                <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                  <a:srgbClr val="000000">
                    <a:alpha val="35000"/>
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
              <a:scene3d>
                <a:camera prst="orthographicFront">
                  <a:rot lat="0" lon="0" rev="0"/>
                </a:camera>
                <a:lightRig rig="threePt" dir="t">
                  <a:rot lat="0" lon="0" rev="1200000"/>
                </a:lightRig>
              </a:scene3d>
              <a:sp3d>
                <a:bevelT w="63500" h="25400"/>
              </a:sp3d>
            </a:effectStyle>
          </a:effectStyleLst>
          <a:bgFillStyleLst>
            <a:solidFill>
              <a:schemeClr val="phClr"/>
            </a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:tint val="40000"/>
                    <a:satMod val="350000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="40000">
                  <a:schemeClr val="phClr">
                    <a:tint val="45000"/>
                    <a:shade val="99000"/>
                    <a:satMod val="350000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:shade val="20000"/>
                    <a:satMod val="255000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:path path="circle">
                <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
              </a:path>
            </a:gradFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:tint val="80000"/>
                    <a:satMod val="300000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:shade val="30000"/>
                    <a:satMod val="200000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:path path="circle">
                <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
              </a:path>
            </a:gradFill>
          </a:bgFillStyleLst>
        </a:fmtScheme>
      </a:themeElements>
      <a:objectDefaults/>
      <a:extraClrSchemeLst/>
    </a:theme>
  </xsl:template>
</xsl:stylesheet>
