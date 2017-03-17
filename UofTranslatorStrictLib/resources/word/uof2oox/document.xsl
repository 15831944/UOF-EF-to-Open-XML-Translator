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
<Author>Li Jingui(BUAA)</Author>
<Author>Zhao Baojing(BUAA</Author>
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:cus="http://purl.oclc.org/ooxml/officeDocument/customProperties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
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
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions" 
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes">
  <xsl:import href="paragraph.xsl"/>
  <xsl:import href="table.xsl"/>
  <xsl:import href="sectPr.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="document">
    <w:document w:conformance="strict">
      
      <!--cxl , current node:字:文字处理文档_4225 , 2012.02.09-->
      <xsl:if test="字:分节_416A/字:节属性_421B/字:填充_4134"><!--background-->
        <w:background>
          <xsl:choose>
            <xsl:when test="字:分节_416A/字:节属性_421B/字:填充_4134/图:颜色_8004">
              <xsl:attribute name="w:color">
                <xsl:value-of select="substring-after(字:分节_416A/字:节属性_421B/字:填充_4134/图:颜色_8004,'#')"/>
              </xsl:attribute>
            </xsl:when>

            <xsl:when test="字:分节_416A/字:节属性_421B/字:填充_4134/图:图片_8005">
              <xsl:apply-templates select="字:分节_416A/字:节属性_421B/字:填充_4134/图:图片_8005" mode="background"/>
            </xsl:when>
            <xsl:when test="字:分节_416A/字:节属性_421B/字:填充_4134/图:图案_800A">
              <xsl:apply-templates select="字:分节_416A/字:节属性_421B/字:填充_4134/图:图案_800A" mode="background"/>
            </xsl:when>
            <xsl:when test="字:分节_416A/字:节属性_421B/字:填充_4134/图:渐变_800D">
              <xsl:apply-templates select="字:分节_416A/字:节属性_421B/字:填充_4134/图:渐变_800D" mode="background"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:color">
                <xsl:value-of select="'auto'"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test="//对象:对象数据_D701/@标识符_D704='documentObj1'">
            <xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = 'documentObj1']/u2opic:picture"/>
          </xsl:if>
        </w:background>
      </xsl:if>
      
      <w:body>
        <xsl:for-each select="字:段落_416B | 字:文字表_416C | 字:分节_416A">
          <xsl:choose>
            <xsl:when test="name(.)='字:段落_416B'">
              <xsl:call-template name="paragraph"/>
            </xsl:when>
            <xsl:when test="name(.)='字:文字表_416C'">
              <xsl:call-template name="table"/>
            </xsl:when>
            <xsl:when test="name(.)='字:分节_416A'">
              <xsl:call-template name="section"/>
            </xsl:when>
            <!--<xsl:when test="name(.)='字:修订开始_421F'">
              --><!----><!--
            </xsl:when>
            <xsl:when test="name(.)='字:逻辑章节_421C'">
              --><!----><!--
            </xsl:when>-->
          </xsl:choose>
        </xsl:for-each>
      </w:body>
    </w:document>
  </xsl:template>

  <xsl:template match="图:图片_8005" mode="background">
    <xsl:attribute name="w:color">
      <xsl:value-of select="'FFFFFF'"/>
    </xsl:attribute>

    <xsl:variable name="picid">
      <xsl:value-of select="@图形引用_8007"/>
    </xsl:variable>
    <!--<xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = $picid]/u2opic:picture"/>-->
    <v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">
      <v:fill recolor="t" type="frame">
        <xsl:attribute name="r:id">
          <xsl:value-of select="'rIdpic1'"/>
        </xsl:attribute>
        <xsl:attribute name="o:title">
          <xsl:value-of select="@名称_8009"/>
        </xsl:attribute>
      </v:fill>
    </v:background>
  </xsl:template>

  <xsl:template match="图:图案_800A" mode="background">
    <!--两个方向表示不一致，为实现互操作-->
    <!--<xsl:if test="//对象:对象数据_D701/@标识符_D704='backgroundpattern006'">-->
      <!--<xsl:copy-of select ="//uof:其他对象[@uof:标识符 = 'backgroundpattern006']/u2opic:picture"/>-->

      <xsl:attribute name="w:color">
        <xsl:value-of select="substring-after(@前景色_800B,'#')"/>
      </xsl:attribute>
      <v:background id="_x0000_s1025">
        <xsl:attribute name="fillcolor">
          <xsl:value-of select="@前景色_800B"/>
        </xsl:attribute>
        <v:fill type="pattern">
			<xsl:attribute name="r:id">
				<xsl:choose>
					<xsl:when test="@类型_8008 ='ptn001'">rIdPattenFill1</xsl:when>
					<xsl:when test="@类型_8008 ='ptn007'">rIdPattenFill2</xsl:when>
					<xsl:when test="@类型_8008 ='ptn013'">rIdPattenFill3</xsl:when>
					<xsl:when test="@类型_8008 ='ptn019'">rIdPattenFill4</xsl:when>
					<xsl:when test="@类型_8008 ='ptn025'">rIdPattenFill5</xsl:when>
					<xsl:when test="@类型_8008 ='ptn031'">rIdPattenFill6</xsl:when>
					<xsl:when test="@类型_8008 ='ptn037'">rIdPattenFill7</xsl:when>
					<xsl:when test="@类型_8008 ='ptn043'">rIdPattenFill8</xsl:when>
					<xsl:when test="@类型_8008 ='ptn002'">rIdPattenFill9</xsl:when>
					<xsl:when test="@类型_8008 ='ptn008'">rIdPattenFill10</xsl:when>
					<xsl:when test="@类型_8008 ='ptn014'">rIdPattenFill11</xsl:when>
					<xsl:when test="@类型_8008 ='ptn020'">rIdPattenFill12</xsl:when>
					<xsl:when test="@类型_8008 ='ptn026'">rIdPattenFill13</xsl:when> <!--匹配-->
					<xsl:when test="@类型_8008 ='ptn032'">rIdPattenFill14</xsl:when>
					<xsl:when test="@类型_8008 ='ptn038'">rIdPattenFill15</xsl:when>
					<xsl:when test="@类型_8008 ='ptn044'">rIdPattenFill16</xsl:when>
					<xsl:when test="@类型_8008 ='ptn003'">rIdPattenFill17</xsl:when>
					<xsl:when test="@类型_8008 ='ptn009'">rIdPattenFill18</xsl:when>
					<xsl:when test="@类型_8008 ='ptn015'">rIdPattenFill19</xsl:when>
					<xsl:when test="@类型_8008 ='ptn021'">rIdPattenFill20</xsl:when>
					<xsl:when test="@类型_8008 ='ptn027'">rIdPattenFill21</xsl:when>
					<xsl:when test="@类型_8008 ='ptn033'">rIdPattenFill22</xsl:when>
					<xsl:when test="@类型_8008 ='ptn039'">rIdPattenFill23</xsl:when>
					<xsl:when test="@类型_8008 ='ptn045'">rIdPattenFill24</xsl:when>
					<xsl:when test="@类型_8008 ='ptn004'">rIdPattenFill25</xsl:when>
					<xsl:when test="@类型_8008 ='ptn010'">rIdPattenFill26</xsl:when>
					<xsl:when test="@类型_8008 ='ptn016'">rIdPattenFill27</xsl:when>
					<xsl:when test="@类型_8008 ='ptn022'">rIdPattenFill28</xsl:when>
					<xsl:when test="@类型_8008 ='ptn028'">rIdPattenFill29</xsl:when>
					<xsl:when test="@类型_8008 ='ptn034'">rIdPattenFill30</xsl:when>
					<xsl:when test="@类型_8008 ='ptn040'">rIdPattenFill31</xsl:when>
					<xsl:when test="@类型_8008 ='ptn046'">rIdPattenFill32</xsl:when>
					<xsl:when test="@类型_8008 ='ptn005'">rIdPattenFill33</xsl:when>
					<xsl:when test="@类型_8008 ='ptn011'">rIdPattenFill34</xsl:when>
					<xsl:when test="@类型_8008 ='ptn017'">rIdPattenFill35</xsl:when>
					<xsl:when test="@类型_8008 ='ptn023'">rIdPattenFill36</xsl:when>
					<xsl:when test="@类型_8008 ='ptn029'">rIdPattenFill37</xsl:when>
					<xsl:when test="@类型_8008 ='ptn035'">rIdPattenFill38</xsl:when>
					<xsl:when test="@类型_8008 ='ptn041'">rIdPattenFill39</xsl:when>
					<xsl:when test="@类型_8008 ='ptn047'">rIdPattenFill40</xsl:when>
					<xsl:when test="@类型_8008 ='ptn006'">rIdPattenFill41</xsl:when>
					<xsl:when test="@类型_8008 ='ptn012'">rIdPattenFill42</xsl:when>
					<xsl:when test="@类型_8008 ='ptn018'">rIdPattenFill43</xsl:when>
					<xsl:when test="@类型_8008 ='ptn024'">rIdPattenFill44</xsl:when>
					<xsl:when test="@类型_8008 ='ptn030'">rIdPattenFill45</xsl:when>
					<xsl:when test="@类型_8008 ='ptn036'">rIdPattenFill46</xsl:when>
					<xsl:when test="@类型_8008 ='ptn042'">rIdPattenFill47</xsl:when>
					<xsl:when test="@类型_8008 ='ptn048'">rIdPattenFill48</xsl:when>
				</xsl:choose>
			</xsl:attribute>
          <xsl:attribute name="o:title">
            <xsl:choose>
              <xsl:when test="@类型_8008 ='ptn001'">5%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn002'">10%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn003'">20%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn004'">25%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn005'">30%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn006'">40%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn007'">50%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn008'">60%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn009'">70%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn010'">75%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn011'">80%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn012'">90%</xsl:when>
              <xsl:when test="@类型_8008 ='ptn013'">浅色下对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn014'">浅色上对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn015'">深色下对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn016'">深色上对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn017'">宽下对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn018'">宽上对角线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn019'">浅色竖线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn020'">浅色横线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn021'">窄竖线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn022'">窄横线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn023'">深色竖线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn024'">深色横线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn025'">下对角虚线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn026'">上对角虚线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn027'">横虚线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn028'">竖虚线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn029'">小纸屑</xsl:when>
              <xsl:when test="@类型_8008 ='ptn030'">大纸屑</xsl:when>
              <xsl:when test="@类型_8008 ='ptn031'">之字形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn032'">波浪线</xsl:when>
              <xsl:when test="@类型_8008 ='ptn033'">对角砖形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn034'">横向砖形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn035'">编织物</xsl:when>
              <xsl:when test="@类型_8008 ='ptn036'">苏格兰方格呢</xsl:when>
              <xsl:when test="@类型_8008 ='ptn037'">草皮</xsl:when>
              <xsl:when test="@类型_8008 ='ptn038'">虚线网格</xsl:when>
              <xsl:when test="@类型_8008 ='ptn039'">点式菱形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn040'">瓦形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn041'">棚架</xsl:when>
              <xsl:when test="@类型_8008 ='ptn042'">球体</xsl:when>
              <xsl:when test="@类型_8008 ='ptn043'">小网格</xsl:when>
              <xsl:when test="@类型_8008 ='ptn044'">大网格</xsl:when>
              <xsl:when test="@类型_8008 ='ptn045'">小棋盘</xsl:when>
              <xsl:when test="@类型_8008 ='ptn046'">大棋盘</xsl:when>
              <xsl:when test="@类型_8008 ='ptn047'">轮廓式菱形</xsl:when>
              <xsl:when test="@类型_8008 ='ptn048'">实心菱形</xsl:when>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="color2">
            <xsl:value-of select="@背景色_800C"/>
          </xsl:attribute>
        </v:fill>

      </v:background>
    <!--</xsl:if>-->
  </xsl:template>

  <xsl:template match="图:渐变_800D" mode="background">
    <xsl:attribute name="w:color">
      <xsl:value-of select="substring-after(@起始色_800E,'#')"/>
    </xsl:attribute>
    <v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">
      <xsl:attribute name="fillcolor">
        <xsl:value-of select="@起始色_800E"/>
      </xsl:attribute>
      <v:fill focus="100%">
        <xsl:attribute name="color2">
          <xsl:value-of select="@终止色_800F"/>
        </xsl:attribute>
        <xsl:if test="not(@渐变方向_8013='0')">
          <xsl:attribute name="angle">
            <xsl:choose>
              <xsl:when test="@渐变方向_8013='90'">
                <xsl:value-of select="-90"/>
              </xsl:when>
              <xsl:when test="@渐变方向_8013='45'">
                <xsl:value-of select="-135"/>
              </xsl:when>
              <xsl:when test="@渐变方向_8013='315'">
                <xsl:value-of select="-45"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="type"><!--cxl,2012.3.18-->
          <xsl:choose>
            <xsl:when test="@种子类型_8010='linear'">
              <xsl:value-of select="'gradient'"/>
            </xsl:when>
            <xsl:when test="@种子类型_8010='square'">
              <xsl:value-of select="'gradientRadial'"/>
            </xsl:when>
            <xsl:when test="@种子类型_8010='radar'">
              <xsl:value-of select="'gradientRadial'"/>
            </xsl:when>
            <xsl:when test="@种子类型_8010='rectangle'">
              <xsl:value-of select="'gradientRadial'"/>
            </xsl:when>
            <xsl:when test="@种子类型_8010='oval'">
              <xsl:value-of select="'gradient'"/>
            </xsl:when>
            <xsl:when test="@种子类型_8010='axial'">
              <xsl:value-of select="'gradient'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="@种子类型_8010='square' and @种子X位置_8015='100' and  @种子Y位置_8016='100'">
          <xsl:attribute name="focusposition">
            <xsl:value-of select="'.5,.5'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@种子类型_8010='square' and @种子X位置_8015='30' and  @种子Y位置_8016='30'">
          <o:fill v:ext="view" type="gradientCenter"/>
        </xsl:if>
        <xsl:if test="@种子类型_8010='rectangle' and @种子X位置_8015='0' and  @种子Y位置_8016='0'">
          <o:fill v:ext="view" type="gradientCenter"/>
        </xsl:if>
      </v:fill>
    </v:background>
  </xsl:template>
</xsl:stylesheet>
