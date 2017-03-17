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
<Author>Li Yanyan(BITI)</Author>
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">

  <xsl:variable name="unit" select="/uof:UOF/uof:文字处理/规则:公用处理规则_B665/规则:长度单位_B666"/>

  <!--
       1 cm = 28.3 point
       1 inch = 72 point
       1 inch = 2.54cm
       1 mm = 2.83 point
  -->


  <xsl:template name="twipsMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 20)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 20)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 20)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 20)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="tableTwipsMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 10)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 10)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 10)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 10)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>


  <xsl:template name="eighthPointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 8)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 8)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 8)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 8)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="halfPointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 2)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 2)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 2)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 2)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="pointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round($lengthVal)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="fourthousandpointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(($lengthVal - 10.5) * 4096)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal - 10.5) * 28.3 * 4096)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal - 10.5) * 2.83 * 4096)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal - 10.5) * 72 * 4096)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="nineThousandpointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 9525)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 9525)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 9525)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 9525)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="onethousandpointMeasure">
    <xsl:param name="lengthVal"/>
    <xsl:choose>
      <xsl:when test="contains($lengthVal,'E')">
        <xsl:value-of select="0"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$unit='pt'">
            <xsl:value-of select="round(number($lengthVal) * 12700)"/>
          </xsl:when>
          <xsl:when test="$unit='cm'">
            <xsl:value-of select="round(number($lengthVal) * 28.3 * 12700)"/>
          </xsl:when>
          <xsl:when test="$unit='mm'">
            <xsl:value-of select="round(number($lengthVal) * 2.83 * 12700)"/>
          </xsl:when>
          <xsl:when test="$unit='inch'">
            <xsl:value-of select="round(number($lengthVal) * 72 * 12700)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="字:填充_4134" mode="shdCommon">
    <w:shd>
      <xsl:choose>
        <xsl:when test="./图:颜色_8004">
          <xsl:attribute name="w:val">clear</xsl:attribute>
          <xsl:attribute name="w:fill">
            <xsl:choose>
              <xsl:when test="./图:颜色_8004='auto'">auto</xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring(./图:颜色_8004,2)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>

        <xsl:when test="./图:图片_8005">
          <xsl:attribute name="w:val">clear</xsl:attribute>
          <xsl:attribute name="w:fill">auto</xsl:attribute>
          <xsl:message>feedback:lost:shading_for_paragraphs_or_tables_or_phrases</xsl:message>
        </xsl:when>

        <xsl:when test="./图:图案_800A">
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn002'">pct10</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn003'">pct20</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn004'">pct25</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn005'">pct30</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn006'">pct40</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn007'">pct50</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn008'">pct60</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn009'">pct70</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn010'">pct75</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn011'">pct80</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn012'">pct90</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn013'">thinReverseDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn014'">thinDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn015'">reverseDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn016'">diagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn017'">reverseDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn018'">diagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn019'">thinVertStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn020'">thinHorzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn021'">vertStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn022'">horzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn023'">vertStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn024'">horzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn025'">thinReverseDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn026'">thinDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn027'">thinVertStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn028'">thinHorzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn029'">horzCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn030'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn031'">thinHorzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn032'">thinHorzStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn033'">diagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn034'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn035'">reverseDiagStripe</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn036'">horzCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn037'">thinDiagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn038'">thinHorzCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn039'">thinDiagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn040'">thinDiagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn041'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn042'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn043'">horzCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn044'">thinHorzCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn045'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn046'">diagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn047'">thinDiagCross</xsl:when>
              <xsl:when test="图:图案_800A/@类型_8008='ptn048'">diagCross</xsl:when>
              <xsl:otherwise>
                <xsl:choose>
                  <xsl:when test="(图:图案_800A/@背景色_800C) and not(图:图案_800A/@前景色_800B)">clear</xsl:when>
                  <xsl:when test="(图:图案_800A/@背景色_800C) and (图:图案_800A/@前景色_800B)">solid</xsl:when>
                  <xsl:when test="not(图:图案_800A/@背景色_800C) and (图:图案_800A/@前景色_800B)">solid</xsl:when>
                  <xsl:when test="not(图:图案_800A/@背景色_800C) and not(图:图案_800A/@前景色_800B)">clear</xsl:when>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="图:图案_800A/@前景色_800B">
              <xsl:attribute name="w:color">
                <xsl:choose>
                  <xsl:when test="图:图案_800A/@前景色_800B!='auto'">
                    <xsl:value-of select="substring-after(图:图案_800A/@前景色_800B,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>auto</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:color">auto</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="图:图案_800A/@背景色_800C">
              <xsl:attribute name="w:fill">
                <xsl:choose>
                  <xsl:when test="图:图案_800A/@背景色_800C!='auto'">
                    <xsl:value-of select="substring-after(图:图案_800A/@背景色_800C,'#')"/>
                  </xsl:when>
                  <xsl:otherwise>auto</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:fill">auto</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>

        <xsl:when test="./图:渐变_800D">
          <xsl:attribute name="w:val">clear</xsl:attribute>
          <xsl:attribute name="w:fill">
            <xsl:choose>
              <xsl:when test="./图:渐变_800D/@起始色_800E='auto'">auto</xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring(./图:渐变_800D/@起始色_800E,2)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:message>feedback:lost:shading_for_paragraphs_or_tables_or_phrases</xsl:message>
        </xsl:when>
      </xsl:choose>
    </w:shd>
  </xsl:template>

</xsl:stylesheet>
