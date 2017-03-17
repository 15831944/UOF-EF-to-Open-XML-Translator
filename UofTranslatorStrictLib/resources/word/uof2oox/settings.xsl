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
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"
  xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main">
  <xsl:import href="common.xsl"/>

  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <xsl:template name="settings">
    <w:settings>
      <!--cxl,2012.3.12增加页面视图代码-->
      <xsl:if test="//规则:文档设置_B600/规则:当前视图_B601">
        <w:view>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="//规则:文档设置_B600/规则:当前视图_B601='page'">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="//规则:文档设置_B600/规则:当前视图_B601='outline'">
                <xsl:value-of select="'outline'"/>
              </xsl:when>
              <xsl:when test="//规则:文档设置_B600/规则:当前视图_B601='web'">
                <xsl:value-of select="'web'"/>
              </xsl:when>
              <xsl:when test="//规则:文档设置_B600/规则:当前视图_B601='normal'">
                <xsl:value-of select="'normal'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </w:view>
      </xsl:if>
      <xsl:if test="//规则:公用处理规则_B665/规则:文档设置_B600/规则:缩放_B603">
        <w:zoom>
          <xsl:attribute name="w:percent">
            <xsl:value-of select="//规则:公用处理规则_B665/规则:文档设置_B600/规则:缩放_B603"/>
          </xsl:attribute>
        </w:zoom>
      </xsl:if>
      <xsl:if test="//字:文字处理文档_4225/字:分节_416A/字:节属性_421B/字:填充_4134">
        <w:displayBackgroundShape/>
      </xsl:if>
      <xsl:if
        test="(//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='top') or (//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='1') or (//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='on')">
        <w:gutterAtTop/>
      </xsl:if>
      <!--cxl,2012.2.22添加,有待案例验证，装订线位于左侧-->
      <!--<xsl:if
        test="(//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='left') or (//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='1') or (//字:分节_416A/字:节属性_421B/字:装订线_41FB/@位置_4150='on')">
        <w:gutterAtLeft/>
      </xsl:if>-->
      <xsl:if
        test="(//规则:公用处理规则_B665/规则:文档设置_B600/规则:是否修订_B605='true') or (//规则:公用处理规则_B665/规则:文档设置_B600/规则:是否修订_B605='1') or (//规则:公用处理规则_B665/规则:文档设置_B600/规则:是否修订_B605='on')">
        <w:trackRevisions/>
      </xsl:if>
      <xsl:if test="//规则:公用处理规则_B665/规则:文档设置_B600/规则:默认制表位位置_B604">
        <w:defaultTabStop>
          <xsl:attribute name="w:val">
            <xsl:call-template name="twipsMeasure">
              <xsl:with-param name="lengthVal" select="//规则:公用处理规则_B665/规则:文档设置_B600/规则:默认制表位位置_B604"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:defaultTabStop>
      </xsl:if>
      <xsl:if test="//规则:公用处理规则_B665//规则:文档设置_B600/规则:字距调整_B606='western-only'">
        <w:noPunctuationKerning/>
      </xsl:if>
      <w:characterSpacingControl w:val="compressPunctuation"/>
      <w:compat>
        <w:spaceForUL/>
        <w:balanceSingleByteDoubleByteWidth/>
        <w:ulTrailSpace/>
        <w:doNotExpandShiftReturn/>
        <w:adjustLineHeightInTable/>
        <w:useFELayout/>
        <w:splitPgBreakAndParaMark/>
        <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"
          w:val="14"/>
        <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"
          w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
        <w:compatSetting w:name="enableOpenTypeFeatures"
          w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
        <w:compatSetting w:name="doNotFlipMirrorIndents"
          w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
      </w:compat>
      
      <xsl:if test="//规则:公用处理规则_B665/规则:文档设置_B600/规则:标点禁则_B608">
        <xsl:if test="//规则:公用处理规则_B665/规则:文档设置_B600/规则:标点禁则_B608/规则:行首字符_B609">
          <w:noLineBreaksBefore w:lang="zh-CN">
            <xsl:attribute name="w:val">
              <xsl:value-of select="//规则:公用处理规则_B665/规则:文档设置_B600/规则:标点禁则_B608/规则:行首字符_B609"/>
            </xsl:attribute>
          </w:noLineBreaksBefore>
        </xsl:if>
        <xsl:if test="//规则:公用处理规则_B665/规则:文档设置_B600/规则:标点禁则_B608/规则:行尾字符_B60A">
          <w:noLineBreaksAfter w:lang="zh-CN">
            <xsl:attribute name="w:val">
              <xsl:value-of select="//规则:公用处理规则_B665/规则:文档设置_B600/规则:标点禁则_B608/规则:行尾字符_B60A"/>
            </xsl:attribute>
          </w:noLineBreaksAfter>
        </xsl:if>
      </xsl:if>

      <xsl:if
        test="(//字:分节_416A/字:节属性_421B/字:是否奇偶页页眉页脚不同_41ED='true') or (//字:分节_416A/字:节属性_421B/字:是否奇偶页页眉页脚不同_41ED='1') or (//字:分节_416A/字:节属性_421B/字:是否奇偶页页眉页脚不同_41ED='on')">
        <w:evenAndOddHeaders/>
      </xsl:if>
      <xsl:if
        test="(//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='true') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='1') or (//字:分节_416A[1]/字:节属性_421B/字:是否拼页_41FE='on')">
        <w:printTwoOnOne/>
      </xsl:if>
      <xsl:if
        test="(//字:分节_416A[1]/字:节属性_421B/字:是否对称页边距_41FD='true') or (//字:分节_416A[1]/字:节属性_421B/字:是否对称页边距_41FD='1') or (//字:分节_416A[1]/字:节属性_421B/字:是否对称页边距_41FD='on')">
        <w:mirrorMargins/>
      </xsl:if>
      <xsl:if test=".//字:句_419D/字:脚注_4159">
        <w:footnotePr>
          <w:footnote w:id="0"/>
          <w:footnote w:id="1"/>
        </w:footnotePr>
      </xsl:if>
      <xsl:if test=".//字:句_419D/字:尾注_415A">
        <w:endnotePr>
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
                  <xsl:otherwise>docEnd</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:pos>
          </xsl:if>
          <w:endnote w:id="0"/>
          <w:endnote w:id="1"/>
        </w:endnotePr>
      </xsl:if>      
    </w:settings>
  </xsl:template>
</xsl:stylesheet>
