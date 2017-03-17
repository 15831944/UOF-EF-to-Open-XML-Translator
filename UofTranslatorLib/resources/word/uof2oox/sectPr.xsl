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
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" 
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <xsl:import href="common.xsl"/>
  <xsl:template name="section">
    <w:sectPr>
      <!--<w:endnotePr>
        <w:numFmt w:val="decimal"/>
      </w:endnotePr>-->
      <xsl:choose>
        <xsl:when test="not(字:修订开始_421F)">
          <xsl:apply-templates select="字:节属性_421B" mode="section"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="字:节属性_421B">
            <xsl:if test="preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224">
              <xsl:call-template name="sectPr"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </w:sectPr>
  </xsl:template>

  <xsl:template match="字:节属性_421B" mode="section">
    <xsl:call-template name="sectPr"/>
  </xsl:template>

  <xsl:template name="sectPr">
    <xsl:call-template name="HeaderFooter"/>
    <xsl:if test="字:脚注设置_4203">
      <w:footnotePr>
        <xsl:call-template name="footnotePr"/>
      </w:footnotePr>
    </xsl:if>
    <xsl:if test="字:尾注设置_4204">
      <w:endnotePr>
        <xsl:call-template name="endnotePr"/>
      </w:endnotePr>
    </xsl:if>

    <xsl:if test="字:节类型_41EA">
      <w:type>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="字:节类型_41EA='continuous'">
              <xsl:value-of select="'continuous'"/>
            </xsl:when>
            <xsl:when test="字:节类型_41EA='even-page'">
              <xsl:value-of select="'evenPage'"/>
            </xsl:when>
            <xsl:when test="字:节类型_41EA='odd-page'">
              <xsl:value-of select="'oddPage'"/>
            </xsl:when>
            <xsl:when test="字:节类型_41EA='new-column'">
              <xsl:value-of select="'nextColumn'"/>
            </xsl:when>
            <xsl:when test="字:节类型_41EA='new-page'">
              <xsl:value-of select="'nextPage'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </w:type>
    </xsl:if>
    <xsl:if test="字:纸张_41EC | 字:纸张方向_41FF">
      <w:pgSz>
        <xsl:variable name="IsPrintTwoOnOne">
          <xsl:if
            test="(//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='true') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='1') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='on')">
            <xsl:value-of select="'true'"/>
          </xsl:if>
          <xsl:if
            test="(//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='false') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='0') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='off') or not(//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE)">
            <xsl:value-of select="'false'"/>
          </xsl:if>
        </xsl:variable>
        <xsl:if test="(字:纸张方向_41FF='portrait') or (not(字:纸张方向_41FF))">
          <xsl:attribute name="w:w">
            <xsl:call-template name="twipsMeasure"><!--common.xsl-->
              <xsl:with-param name="lengthVal" select="字:纸张_41EC/@宽_C605"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="w:h">
            <xsl:choose>
              <xsl:when test="$IsPrintTwoOnOne='true'">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:纸张_41EC/@长_C604 div 2"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:纸张_41EC/@长_C604"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:if test="字:纸张方向_41FF">
            <xsl:attribute name="w:orient">
              <xsl:value-of select="字:纸张方向_41FF"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
        <xsl:if test="字:纸张方向_41FF='landscape'">
          <xsl:attribute name="w:w">
            <xsl:choose>
              <xsl:when test="$IsPrintTwoOnOne='true'">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:纸张_41EC/@宽_C605 div 2"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:纸张_41EC/@宽_C605"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="w:h">
            <xsl:call-template name="twipsMeasure">
              <xsl:with-param name="lengthVal" select="字:纸张_41EC/@长_C604"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:attribute name="w:orient">
            <xsl:value-of select="字:纸张方向_41FF"/>
          </xsl:attribute>
        </xsl:if>
      </w:pgSz>
    </xsl:if>
    <xsl:if test="字:页边距_41EB | 字:装订线_41FB | 字:页眉位置_41EF | 字:页脚位置_41F2">
      <w:pgMar>
        <xsl:variable name="left">
          <xsl:choose>
            <xsl:when test="字:页边距_41EB/@左_C608">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页边距_41EB/@左_C608"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="top">
          <xsl:choose>
            <xsl:when test="字:页边距_41EB/@上_C609">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页边距_41EB/@上_C609"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="right">
          <xsl:choose>
            <xsl:when test="字:页边距_41EB/@右_C60A">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页边距_41EB/@右_C60A"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="bottom">
          <xsl:choose>
            <xsl:when test="字:页边距_41EB/@下_C60B">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页边距_41EB/@下_C60B"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="leftRight" select="round(($left + $right) div 2)"/>
        <xsl:variable name="topBottom" select="round(($top + $bottom) div 2)"/>
        <xsl:attribute name="w:left">
          <xsl:value-of select="$left"/>
        </xsl:attribute>
        <xsl:attribute name="w:right">
          <xsl:value-of select="$right"/>
        </xsl:attribute>
        <xsl:attribute name="w:top">
          <xsl:value-of select="$top"/>
        </xsl:attribute>
        <xsl:attribute name="w:bottom">
          <xsl:value-of select="$bottom"/>
        </xsl:attribute>
        <xsl:choose>
          <xsl:when test="字:装订线_41FB/@距边界_41FC">
            <xsl:attribute name="w:gutter">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:装订线_41FB/@距边界_41FC"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:gutter">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:attribute name="w:header">
          <xsl:choose>
            <xsl:when test="字:页眉位置_41EF">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页眉位置_41EF/@距边界_41F0"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="w:footer">
          <xsl:choose>
            <xsl:when test="字:页脚位置_41F2">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:页脚位置_41F2/@距边界_41F0"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:pgMar>
    </xsl:if>
    <xsl:if test="字:纸张来源_4200">
      <w:paperSrc>
        <xsl:if test="字:纸张来源_4200/@首页_4201">
          <xsl:attribute name="w:first">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="字:纸张来源_4200/@其他页_4202">
          <xsl:attribute name="w:other">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
      </w:paperSrc>
    </xsl:if>
    <xsl:if test="字:边框_4133">
      <w:pgBorders>
        
        <!--2013-03-19，wudi，修复页面边框BUG，start-->
        <xsl:if test ="字:边框_4133/@应用范围_4229 ='first'">
          <xsl:attribute name ="w:display">
            <xsl:value-of select ="'firstPage'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="字:边框_4133/@应用范围_4229 ='except-first'">
          <xsl:attribute name ="w:display">
            <xsl:value-of select ="'notfirstPage'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="字:边框_4133/@度量依据_4230 ='page-edge'">
          <xsl:attribute name="w:offsetFrom">
            <xsl:value-of select="'page'"/>
          </xsl:attribute>
        </xsl:if>
        <!--end-->
        
        <xsl:if test="字:边框_4133/uof:上_C614">
          <w:top>
            <xsl:if test="字:边框_4133/uof:上_C614/@线型_C60D or 字:边框_4133/uof:上_C614/@虚实_C60E">
              <xsl:attribute name="w:val">
                <xsl:call-template name="olineType">
                  <xsl:with-param name="linetype" select="字:边框_4133/uof:上_C614/@线型_C60D"/>
                  <xsl:with-param name="linexs" select="字:边框_4133/uof:上_C614/@虚实_C60E"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:上_C614/@宽度_C60F">
              <xsl:attribute name="w:sz">
                <xsl:call-template name="eighthPointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:上_C614/@宽度_C60F"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:上_C614/@边距_C610">
              <xsl:attribute name="w:space">
                <xsl:call-template name="pointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:上_C614/@边距_C610"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/@阴影类型_C645='right-bottom'">
              <xsl:attribute name="w:shadow">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:上_C614/@颜色_C611">
              <xsl:variable name="topcolor">
                <xsl:value-of select="字:边框_4133/uof:上_C614/@颜色_C611"/>
              </xsl:variable>
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="contains($topcolor,'#')">
                    <xsl:value-of select="substring-after(字:边框_4133/uof:上_C614/@颜色_C611,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="字:边框_4133/uof:上_C614/@颜色_C611"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:top>
        </xsl:if>
        <xsl:if test="字:边框_4133/uof:左_C613">
          <w:left>
            <xsl:if test="字:边框_4133/uof:左_C613/@线型_C60D or 字:边框_4133/uof:左_C613/@虚实_C60E">
              <xsl:attribute name="w:val">
                <xsl:call-template name="olineType">
                  <xsl:with-param name="linetype" select="字:边框_4133/uof:左_C613/@线型_C60D "/>
                  <xsl:with-param name="linexs" select="字:边框_4133/uof:左_C613/@虚实_C60E "/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:左_C613/@宽度_C60F">
              <xsl:attribute name="w:sz">
                <xsl:call-template name="eighthPointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:左_C613/@宽度_C60F"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:左_C613/@边距_C610">
              <xsl:attribute name="w:space">
                <xsl:call-template name="pointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:左_C613/@边距_C610"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/@阴影类型_C645='right-bottom'">
              <xsl:attribute name="w:shadow">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:左_C613/@颜色_C611">
              <xsl:variable name="leftcolor">
                <xsl:value-of select="字:边框_4133/uof:左_C613/@颜色_C611"/>
              </xsl:variable>
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="contains($leftcolor,'#')">
                    <xsl:value-of select="substring-after(字:边框_4133/uof:左_C613/@颜色_C611,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="字:边框_4133/uof:左_C613/@颜色_C611"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:left>
        </xsl:if>
        <xsl:if test="字:边框_4133/uof:下_C616">
          <w:bottom>
            <xsl:if test="字:边框_4133/uof:下_C616/@线型_C60D or 字:边框_4133/uof:下_C616/@虚实_C60E">
              <xsl:attribute name="w:val">
                <xsl:call-template name="olineType">
                  <xsl:with-param name="linetype" select="字:边框_4133/uof:下_C616/@线型_C60D "/>
                  <xsl:with-param name="linexs" select="字:边框_4133/uof:下_C616/@虚实_C60E "/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:下_C616/@宽度_C60F">
              <xsl:attribute name="w:sz">
                <xsl:call-template name="eighthPointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:下_C616/@宽度_C60F"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:下_C616/@边距_C610">
              <xsl:attribute name="w:space">
                <xsl:call-template name="pointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:下_C616/@边距_C610"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/@阴影类型_C645='right-bottom'">
              <xsl:attribute name="w:shadow">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:下_C616/@颜色_C611">
              <xsl:variable name="bottomcolor">
                <xsl:value-of select="字:边框_4133/uof:下_C616/@颜色_C611"/>
              </xsl:variable>
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="contains($bottomcolor,'#')">
                    <xsl:value-of select="substring-after(字:边框_4133/uof:下_C616/@颜色_C611,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="字:边框_4133/uof:下_C616/@颜色_C611"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:bottom>
        </xsl:if>
        <xsl:if test="字:边框_4133/uof:右_C615">
          <w:right>
            <xsl:if test="字:边框_4133/uof:右_C615/@线型_C60D or 字:边框_4133/uof:右_C615/@虚实_C60E">
              <xsl:attribute name="w:val">
                <xsl:call-template name="olineType">
                  <xsl:with-param name="linetype" select="字:边框_4133/uof:右_C615/@线型_C60D "/>
                  <xsl:with-param name="linexs" select="字:边框_4133/uof:右_C615/@虚实_C60E "/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:右_C615/@宽度_C60F">
              <xsl:attribute name="w:sz">
                <xsl:call-template name="eighthPointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:右_C615/@宽度_C60F"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:右_C615/@边距_C610">
              <xsl:attribute name="w:space">
                <xsl:call-template name="pointMeasure">
                  <xsl:with-param name="lengthVal" select="字:边框_4133/uof:右_C615/@边距_C610"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/@阴影类型_C645='right-bottom'">
              <xsl:attribute name="w:shadow">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:边框_4133/uof:右_C615/@颜色_C611">
              <xsl:variable name="rightcolor">
                <xsl:value-of select="字:边框_4133/uof:右_C615/@颜色_C611"/>
              </xsl:variable>
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="contains($rightcolor,'#')">
                    <xsl:value-of select="substring-after(字:边框_4133/uof:右_C615/@颜色_C611,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="字:边框_4133/uof:右_C615/@颜色_C611"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:right>
        </xsl:if>
      </w:pgBorders>
    </xsl:if>

    <xsl:if test="字:行号设置_420A[(@是否使用行号_420B='true') or (@是否使用行号_420B='1') or(@是否使用行号_420B='on')]">
      <w:lnNumType>
        <xsl:if test="字:行号设置_420A/@编号方式_4153">
          <xsl:attribute name="w:restart">
            <xsl:call-template name="olnNumTypeRestart">
              <xsl:with-param name="val" select="字:行号设置_420A/@编号方式_4153"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>

        <!--<xsl:attribute name="w:countBy">1</xsl:attribute>-->     
        <!--cxl,2012.3.17修改-->        
        <xsl:if test="字:行号设置_420A/@行号间隔_420D">
          <xsl:attribute name="w:countBy">
            <xsl:value-of select="字:行号设置_420A/@行号间隔_420D"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="字:行号设置_420A/@距边界_41F0">
          <xsl:attribute name="w:distance">
            <xsl:call-template name="twipsMeasure">
              <xsl:with-param name="lengthVal" select="字:行号设置_420A/@距边界_41F0"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>       
        
        <xsl:if test="字:行号设置_420A/@起始编号_420C">
          <xsl:attribute name="w:start">
            <xsl:value-of select="字:行号设置_420A/@起始编号_420C - 1"/>
          </xsl:attribute>
        </xsl:if>
      </w:lnNumType>
    </xsl:if>
    <xsl:if test="字:页码设置_4205">
      <w:pgNumType>
        <xsl:if test="字:页码设置_4205/@格式_4151">
          <xsl:attribute name="w:fmt">
            <xsl:variable name="pagenumtype">
              <xsl:value-of select="字:页码设置_4205/@格式_4151"/>
            </xsl:variable>
            <xsl:choose>
              <!--第一种decimal对应默认，第二种uof没有实现-->

              <xsl:when test="$pagenumtype='decimal-full-width'">
                <xsl:value-of select="'decimalFullWidth'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='lower-letter'">
                <xsl:value-of select="'lowerLetter'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='upper-letter'">
                <xsl:value-of select="'upperLetter'"/>
              </xsl:when>

              <!--2013-05-03，wudi，修复UOF到OOX方向页码设置转换BUG，遗漏了一种页码格式，start-->
              <xsl:when test ="$pagenumtype='decimal-in-dash'">
                <xsl:value-of select ="'numberInDash'"/>
              </xsl:when>
              <!--end-->

              <xsl:when test="$pagenumtype='lower-roman'">
                <xsl:value-of select="'lowerRoman'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='upper-roman'">
                <xsl:value-of select="'upperRoman'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='chinese-counting'">
                <xsl:value-of select="'chineseCountingThousand'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='chinese-legal-simplified'">
                <xsl:value-of select="'chineseLegalSimplified'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='ideograph-traditional'">
                <xsl:value-of select="'ideographTraditional'"/>
              </xsl:when>
              <xsl:when test="$pagenumtype='ideograph-zodiac'">
                <xsl:value-of select="'ideographZodiac'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$pagenumtype"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="字:页码设置_4205/@起始编号_4152">
          <xsl:attribute name="w:start">
            <xsl:value-of select="字:页码设置_4205/@起始编号_4152"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="字:页码设置_4205/@章节起始样式引用_4208">
          <xsl:attribute name="w:chapStyle">
            <xsl:variable name="id" select="字:页码设置_4205/@章节起始样式引用_4208"/>
            <xsl:apply-templates select="//式样:段落式样_9912[@标识符_4100=$id]" mode="pageNum"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="字:页码设置_4205/@分隔符_4209">
          <xsl:attribute name="w:chapSep">
            <xsl:call-template name="opgNumChapSep">
              <xsl:with-param name="val" select="字:页码设置_4205/@分隔符_4209"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
      </w:pgNumType>
    </xsl:if>
    <xsl:if test="字:分栏_4215">
      <w:cols>
        <xsl:attribute name="w:num">
          <xsl:choose>
            <xsl:when test="字:分栏_4215/字:栏数_41E8">
              <xsl:value-of select="字:分栏_4215/字:栏数_41E8"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'1'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!--分隔线UOF有多种，OOX只有默认一种-->
        <xsl:if test="字:分栏_4215/字:分隔线_41E3/@分隔线线型_41E4='single'">
          <xsl:attribute name="w:sep">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="w:equalWidth">
          <xsl:value-of select="字:分栏_4215/字:是否等宽_41E9"/>
        </xsl:attribute>
        <xsl:if test="字:分栏_4215/字:是否等宽_41E9='true'">
          <xsl:if test="字:分栏_4215/字:栏_41E0/@间距_41E2">
            <xsl:attribute name="w:space">
              <xsl:call-template name="twipsMeasure">
                <xsl:with-param name="lengthVal" select="字:分栏_4215/字:栏_41E0/@间距_41E2"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
        <xsl:if test="字:分栏_4215/字:是否等宽_41E9='false'">
          <xsl:for-each select="字:分栏_4215/字:栏_41E0">
            <w:col>
              <xsl:if test="./@宽度_41E1">
                <xsl:attribute name="w:w">
                  <xsl:call-template name="twipsMeasure">
                    <xsl:with-param name="lengthVal" select="@宽度_41E1"/>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="./@间距_41E2">
                <xsl:attribute name="w:space">
                  <xsl:call-template name="twipsMeasure">
                    <xsl:with-param name="lengthVal" select="@间距_41E2"/>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:if>
            </w:col>
          </xsl:for-each>
        </xsl:if>
      </w:cols>
    </xsl:if>
    <xsl:if test="字:垂直对齐方式_4213">
      <w:vAlign>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="字:垂直对齐方式_4213='justified'">
              <xsl:value-of select="'both'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="字:垂直对齐方式_4213"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:vAlign>
    </xsl:if>

    <!--titlePg-->
    <xsl:if test="字:是否首页页眉页脚不同_41EE">
      <w:titlePg>
        <xsl:attribute name="w:val">
          <xsl:value-of select="字:是否首页页眉页脚不同_41EE"/>
        </xsl:attribute>
      </w:titlePg>
    </xsl:if>
    
    <xsl:if test="字:文字排列方向_4214">
       <!--文字方向影响集成测试，未找到原因先注释掉 -->
      
      <!--2014-04-08，wudi，修改if条件语句，start-->
      
      <!--2014-03-27，wudi，注释掉if条件语句，start-->
      <xsl:if test ="not(字:文字排列方向_4214 ='t2b-l2r-0e-0w')">        
        <w:textDirection>
          <xsl:attribute name="w:val">
           <xsl:call-template name="textDirection">
              <xsl:with-param name="dir" select="字:文字排列方向_4214"/>
           </xsl:call-template>
          </xsl:attribute>
        </w:textDirection>     
      </xsl:if>
      <!--end-->
      
      <!--end-->
      
    </xsl:if>
    <!--docGrid这里与UOF1.1变化较大，宽度、高度属性分别对应UOF2.0的哪个属性？-->
    <xsl:if test="字:网格设置_420E">
      <w:docGrid>
        <xsl:if test="字:网格设置_420E/@列跨度_4244">
          <xsl:attribute name="w:charSpace">
            <!--yx,add 网格设置的对应,删除下面的内容，改为xsl:value-of的形式，因为用原来的好像没有调用模板why09.1.5-->
            <!--xsl:call-template name ="fourthousandpointMeasure">
              <xsl:with-param name ="lengthVal" select ="字:网格设置/@字:宽度"/>
            </xsl:call-template-->
            <xsl:value-of select="round(((字:网格设置_420E/@列跨度_4244)-10.5)*4096)"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="字:网格设置_420E/@行跨度_4243">
          <xsl:attribute name="w:linePitch">
            <!--yx,add 网格设置的对应,删除下面的内容，改为xsl:value-of的形式，因为用原来的好像没有调用模板why09.1.5-->
            <!--xsl:call-template name ="twipsMeasure">
              <xsl:with-param name ="lengthVal" select ="字:网格设置/@字:高度"/>
            </xsl:call-template-->
            <xsl:value-of select="round((字:网格设置_420E/@行跨度_4243)*20)"/>
          </xsl:attribute>
        </xsl:if>

        <xsl:if test="字:网格设置_420E/@网格类型_420F">
          <xsl:attribute name="w:type">
            <xsl:choose>
              <xsl:when test="字:网格设置_420E/@网格类型_420F='none'">
                <xsl:value-of select="'default'"/>
              </xsl:when>
              <xsl:when test="字:网格设置_420E/@网格类型_420F='line-char'">
                <xsl:value-of select="'linesAndChars'"/>
              </xsl:when>
              <xsl:when test="字:网格设置_420E/@网格类型_420F='line'">
                <xsl:value-of select="'lines'"/>
              </xsl:when>
              <xsl:when test="字:网格设置_420E/@网格类型_420F='char'">
                <xsl:value-of select="'snapToChars'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </w:docGrid>
    </xsl:if>

    <xsl:if
      test="preceding-sibling::字:节属性_421B[not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224)]">
      <w:sectPrChange>
        <xsl:apply-templates select="preceding-sibling::字:节属性_421B" mode="sectPr"/>
      </w:sectPrChange>
    </xsl:if>
  </xsl:template>
  <xsl:template match="字:节属性_421B" mode="sectPr">
    <w:sectPr>
      <xsl:call-template name="sectPr"/>
    </w:sectPr>
  </xsl:template>
  <!--yx,add 网格设置的对应09.1.5-->
  <!--xsl:template name ="fourthousandpointMeasure">
    <xsl:param name ="lengthVal"/>
    <xsl:value-of select="((409.5 div lengthVal)-10.5)*4096"/>
  </xsl:template>
  <xsl:template name ="twipsMeasure">
    <xsl:param name ="lengthVal"/>
    <xsl:value-of select="(686.4 div lengthVal)*20"/>
  </xsl:template-->
  <!--yx,add 网格设置的对应09.1.5-->
  <xsl:template match="式样:段落式样_9912" mode="pageNum">
    <xsl:variable name="name">
      <xsl:choose>
        <xsl:when test="@别名_4103">
          <xsl:value-of select="@别名_4103"/>
        </xsl:when>
        <xsl:when test="not(@别名_4103) and @名称_4101">
          <xsl:value-of select="@名称_4101"/>
        </xsl:when>
        <xsl:when test="not(@名称_4101 and @别名_4103)">
          <xsl:value-of select="'heading 1'"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="contains($name,'heading')">
        <xsl:choose>
          <xsl:when test="contains($name,'1')">
            <xsl:value-of select="'1'"/>
          </xsl:when>
          <xsl:when test="contains($name,'2')">
            <xsl:value-of select="'2'"/>
          </xsl:when>
          <xsl:when test="contains($name,'3')">
            <xsl:value-of select="'3'"/>
          </xsl:when>
          <xsl:when test="contains($name,'4')">
            <xsl:value-of select="'4'"/>
          </xsl:when>
          <xsl:when test="contains($name,'5')">
            <xsl:value-of select="'6'"/>
          </xsl:when>
          <xsl:when test="contains($name,'7')">
            <xsl:value-of select="'7'"/>
          </xsl:when>
          <xsl:when test="contains($name,'8')">
            <xsl:value-of select="'8'"/>
          </xsl:when>
          <xsl:when test="contains($name,'9')">
            <xsl:value-of select="'9'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'9'"/>
            <xsl:message terminate="no"
              >feedback:lost:pgNumType_and_chapStyle_in_Section:Section</xsl:message>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'1'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="footnotePr">

    <xsl:if test="字:脚注设置_4203/@位置_4150">
      <w:pos>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="字:脚注设置_4203/@位置_4150 = 'page-bottom'">
              <xsl:value-of select="'pageBottom'"/>
            </xsl:when>
            <xsl:when test="字:脚注设置_4203/@位置_4150= 'below-text'">
              <xsl:value-of select="'beneathText'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'pageBottom'"/>
              <xsl:message terminate="no"
                >feedback:lost:footnotePr_and_pos_in_Section:Section</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:pos>
    </xsl:if>
    <xsl:if test="字:脚注设置_4203/@格式_4151">
      <w:numFmt>
        <xsl:attribute name="w:val">
          <xsl:call-template name="onotePrNumFmt">
            <xsl:with-param name="val" select="./字:脚注设置_4203/@格式_4151"/>
          </xsl:call-template>
        </xsl:attribute>
      </w:numFmt>
    </xsl:if>
    <xsl:if test="字:脚注设置_4203/@起始编号_4152">
      <w:numStart>
        <xsl:attribute name="w:val">
          <xsl:value-of select="字:脚注设置_4203/@起始编号_4152"/>
        </xsl:attribute>
      </w:numStart>
    </xsl:if>
    <xsl:if test="字:脚注设置_4203/@编号方式_4153">
      <w:numRestart>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="字:脚注设置_4203/@编号方式_4153 = 'continuous'">
              <xsl:value-of select="'continuous'"/>
            </xsl:when>
            <xsl:when test="字:脚注设置_4203/@编号方式_4153 = 'section'">
              <xsl:value-of select="'eachSect'"/>
            </xsl:when>
            <xsl:when test="字:脚注设置_4203/@编号方式_4153 = 'page'">
              <xsl:value-of select="'eachPage'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'continuous'"/>
              <xsl:message terminate="no"
                >feedback:lost:footnotePr_and_numRestart_in_Section:Section</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:numRestart>
    </xsl:if>
  </xsl:template>

  <xsl:template name="endnotePr">
    <xsl:if test="//规则:尾注位置_B607">
      <xsl:variable name="posval">
        <xsl:value-of select="//规则:尾注位置_B607"/>
      </xsl:variable>
      <w:pos>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <xsl:when test="$posval = 'doc-end'">
              <xsl:value-of select="'docEnd'"/>
            </xsl:when>
            <xsl:when test="$posval= 'section-end'">
              <xsl:value-of select="'sectEnd'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'docEnd'"/>
              <xsl:message terminate="no"
                >feedback:lost:endnotePr_and_pos_in_Section:Section</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:pos>
    </xsl:if>


    <xsl:if test="字:尾注设置_4204/@格式_4151">
      <w:numFmt>
        <xsl:attribute name="w:val">
          <xsl:call-template name="onotePrNumFmt">
            <xsl:with-param name="val" select="./字:尾注设置_4204/@格式_4151"/>
          </xsl:call-template>
        </xsl:attribute>
      </w:numFmt>
    </xsl:if>

    <xsl:if test="字:尾注设置_4204/@起始编号_4152">
      <w:numStart>
        <xsl:attribute name="w:val">
          <xsl:value-of select="字:尾注设置_4204/@起始编号_4152"/>
        </xsl:attribute>
      </w:numStart>
    </xsl:if>
    <xsl:if test="字:尾注设置_4204/@编号方式_4153">
      <w:numRestart>
        <xsl:attribute name="w:val">
          <xsl:choose>
            <!--
            <xsl:when test="字:尾注设置_4204/@字:编号方式 = 'continuous'">
              <xsl:value-of select="'continuous'"/>
            </xsl:when>
            -->
            <xsl:when test="字:尾注设置_4204/@编号方式_4153 = 'section'">
              <xsl:value-of select="'eachSect'"/>
            </xsl:when>
            <xsl:when test="字:尾注设置_4204/@编号方式_4153 = 'page'">
              <xsl:value-of select="'eachPage'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'continuous'"/>
              <xsl:message terminate="no">feedback:lost:endnotePr_and_numRestart_in_Section:Section</xsl:message>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </w:numRestart>
    </xsl:if>
  </xsl:template>

  <!--2014-03-27，wudi，修改文字排列方向转换时的处理方法，start-->
  <xsl:template name="textDirection">
    <xsl:param name="dir"/>
    <xsl:choose>
      
      <!--2014-05-20，wudi，修改tbRlV为tbRl,修复文字排列方向转换错误，start-->
      <xsl:when test="$dir='r2l-t2b-0e-90w'">
        <xsl:value-of select="'tbRl'"/>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="$dir='r2l-t2b-90e-90w'">
        <xsl:value-of select="'btLr'"/>
      </xsl:when>
      <xsl:when test="$dir='t2b-l2r-270e-0w'">
        <xsl:value-of select="'lrTbV'"/>
      </xsl:when>
      <xsl:when test="$dir='l2r-b2t-270e-270w'">
        <xsl:value-of select="'btLr'"/>
      </xsl:when>
      
      <!--<xsl:otherwise>
        <xsl:message terminate="no">feedback:lost:textDirection_in_Section:Section</xsl:message>
      </xsl:otherwise>-->
    </xsl:choose>
  </xsl:template>
  <!--end-->
  
  <xsl:template name="onotePrNumFmt">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='decimal'">
        <xsl:value-of select="'decimal'"/>
      </xsl:when>
      <xsl:when test="$val='upper-roman'">
        <xsl:value-of select="'upperRoman'"/>
      </xsl:when>
      <xsl:when test="$val='lower-roman'">
        <xsl:value-of select="'lowerRoman'"/>
      </xsl:when>
      <xsl:when test="$val='upper-letter'">
        <xsl:value-of select="'upperLetter'"/>
      </xsl:when>
      <xsl:when test="$val='lower-letter'">
        <xsl:value-of select="'lowerLetter'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal'">
        <xsl:value-of select="'ordinal'"/>
      </xsl:when>
      <xsl:when test="$val='cardinal-text'">
        <xsl:value-of select="'cardinalText'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal-text'">
        <xsl:value-of select="'ordinalText'"/>
      </xsl:when>
      <xsl:when test="$val='hex'">
        <xsl:value-of select="'hex'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-full-width'">
        <xsl:value-of select="'decimalFullWidth'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-half-width'">
        <xsl:value-of select="'decimalHalfWidth'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-circle'">
        <xsl:value-of select="'decimalEnclosedCircle'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-fullstop'">
        <xsl:value-of select="'decimalEnclosedFullstop'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-paren'">
        <xsl:value-of select="'decimalEnclosedParen'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-circle-chinese'">
        <xsl:value-of select="'decimalEnclosedCircleChinese'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-enclosed-circle'">
        <xsl:value-of select="'ideographEnclosedCircle'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-traditional'">
        <xsl:value-of select="'ideographTraditional'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-zodiac'">
        <xsl:value-of select="'ideographZodiac'"/>
      </xsl:when>
      <xsl:when test="$val='chinese-counting'">
        <xsl:value-of select="'chineseCountingThousand'"/>
      </xsl:when>
      <xsl:when test="$val='chinese-legal-simplified'">
        <xsl:value-of select="'chineseLegalSimplified'"/>
      </xsl:when>
      <!--<xsl:otherwise>
        <xsl:value-of select="'decimal'"/>
        <xsl:message terminate="no"
          >feedback:lost:numFmt_and_val_in_Section_Properties:SectionPr</xsl:message>
      </xsl:otherwise>-->
    </xsl:choose>
  </xsl:template>


  <xsl:template name="olnNumTypeRestart">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='page'">
        <xsl:value-of select="'newPage'"/>
      </xsl:when>
      <xsl:when test="$val='section'">
        <xsl:value-of select="'newSection'"/>
      </xsl:when>
      <xsl:when test="$val='continuous'">
        <xsl:value-of select="'continuous'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message terminate="no">
          feedback:lost:lnNumType_and_lnNumTypeRestart_in_Section_Properties:Section Properties </xsl:message>
        <xsl:value-of select="'continuous'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <!--边框的转换 Lee 2011年5月11日-->
  <xsl:template name="olineType">
    <xsl:param name="linetype"/>
    <xsl:param name="linexs"/>
    <xsl:choose>

      <xsl:when test="$linetype='none'">
        <xsl:value-of select="'none'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='solid'">
        <xsl:value-of select="'single'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='square-dot'">
        <xsl:value-of select="'dotted'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='round-dot'">
        <xsl:value-of select="'dashSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='dash'">
        <xsl:value-of select="'dashed'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='dash-dot'">
        <xsl:value-of select="'dotDash'"/>
      </xsl:when>
      <xsl:when test="$linetype='single' and $linexs='dash-dot-dot'">
        <xsl:value-of select="'dotDotDash'"/>
      </xsl:when>

      <!--dotted-heavy这一条在ISO-UOF中可以不考虑 Lee-->
      <xsl:when test="$linetype='dotted-heavy' or $linexs='dotted-heavy'">
        <xsl:value-of select="'dotted'"/>
      </xsl:when>

      <!--case only have double-->

      <xsl:when test="$linetype='double' and $linexs='solid'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='dash-dot-dot'">
        <xsl:value-of select="'triple'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-thin' and $linexs='solid'">
        <xsl:value-of select="'thinThickSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thin-thick' and $linexs='solid'">
        <xsl:value-of select="'thickThinSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-thin' and $linexs='dash'">
        <xsl:value-of select="'thinThickMediumGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thin-thick' and $linexs='dash'">
        <xsl:value-of select="'thickThinMediumGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thin-thick' and $linexs='long-dash'">
        <xsl:value-of select="'thickThinLargeGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-thin' and $linexs='long-dash'">
        <xsl:value-of select="'thinThickLargeGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-between-thin' and $linexs='solid'">
        <xsl:value-of select="'thinThickThinSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-between-thin' and $linexs='dash'">
        <xsl:value-of select="'thinThickThinMediumGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-between-thin' and $linexs='long-dash'">
        <xsl:value-of select="'thinThickThinLargeGap'"/>
      </xsl:when>
      
      <!--2013-03-28，wudi，修复页面边框的BUG，start-->
      <xsl:when test="$linetype='single' and $linexs='wave'">
        <xsl:value-of select="'wave'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='wave'">
        <xsl:value-of select="'doubleWave'"/>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="$linetype='single' and $linexs='long-dash-dot'">
        <xsl:value-of select="'dashDotStroked'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='dash'">
        <xsl:value-of select="'threeDEmboss'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='dash-dot'">
        <xsl:value-of select="'threeDEngrave'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='long-dash'">
        <xsl:value-of select="'outset'"/>
      </xsl:when>
      <xsl:when test="$linetype='double' and $linexs='long-dash-dot'">
        <xsl:value-of select="'inset'"/>
      </xsl:when>
      <xsl:when test="$linetype='wave' and $linexs='solid'">
        <xsl:value-of select="'wave'"/>
      </xsl:when>

      <xsl:when test="$linetype='dash-dot'">
        <xsl:value-of select="'dashDotStroked'"/>
      </xsl:when>

      <xsl:when test="$linetype='double'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-thin'">
        <xsl:value-of select="'thinThickSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='thin-thick'">
        <xsl:value-of select="'thickThinSmallGap'"/>
      </xsl:when>
      <xsl:when test="$linetype='double'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick-between-thin'">
        <xsl:value-of select="'thinThickThinSmallGap'"/>
      </xsl:when>
      <!--以上枚举 测试没问题 Lee-->
      <xsl:otherwise>
        <xsl:value-of select="'single'"/>
        <xsl:message terminate="no">feedback:lost:table_border_style_in_Table:Table</xsl:message>
      </xsl:otherwise>

    </xsl:choose>
  </xsl:template>
  <xsl:template name="opgNumChapSep">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='em-dash'">
        <xsl:value-of select="'emDash'"/>
      </xsl:when>
      <xsl:when test="$val='en-dash'">
        <xsl:value-of select="'enDash'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$val"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
