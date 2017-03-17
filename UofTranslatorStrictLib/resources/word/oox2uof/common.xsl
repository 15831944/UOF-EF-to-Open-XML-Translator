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
<Author>Ban Qianchao(BUAA)</Author>
<Author>Li Yanyan(BITI)</Author>
-->
<xsl:stylesheet version="1.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:fo="http://www.w3.org/1999/XSL/Format" 
xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" 
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles">

  
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  
  <xsl:template name="sdtContentBlock">
    <xsl:param name ="sdtPartFrom"/>
    
    <!--2012-11-27，wudi，OOX到UOF方向的目录与索引的实现，start-->

    <!--2014-03-24，wudi，修复表格转换BUG，单元格里存在sdt节点的情况，start-->
    <xsl:choose>
      <!--<xsl:when test="name(.)='w:sdtPr'">
        
      </xsl:when>
      <xsl:when test="name(.)='w:sdtEndPr'">
        
      </xsl:when>-->
      <xsl:when test="name(.)='w:sdtContent'">
        <xsl:for-each select="w:p | w:tbl | w:sdt">
          <xsl:if test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name="pPartFrom" select="$sdtPartFrom"/>
            </xsl:call-template>
          </xsl:if>
          <xsl:if test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:if>
          <xsl:if test="name(.)='w:sdt'">
            <xsl:call-template name="sdtContentBlock">
              <xsl:with-param name="sdtPartFrom" select="$sdtPartFrom"/>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
    </xsl:choose>
    <!--end-->
    
    <!--end-->
    
  </xsl:template>


  <xsl:template name="sdtContentRun">
    <xsl:param name ="sdtPartFrom"/>
    <xsl:for-each select="./w:sdtContent/node()">
      <xsl:if test="name(.)='w:r'">
        <xsl:if test="not(./w:fldChar) and not(./w:instrText) and not(contains(preceding-sibling::w:r/w:instrText,'HYPERLINK'))">
          <字:句_419D>
            <xsl:call-template name="run">
              <xsl:with-param name ="rPartFrom" select ="$sdtPartFrom"/>
            </xsl:call-template>
          </字:句_419D>
        </xsl:if>
        <xsl:if test="./w:fldChar or ./w:instrText">
          <xsl:if test="./w:fldChar/@w:fldCharType='begin'">
            <!--and(contains(./following::*/w:instrText[1],'FILENAME'))">-->
            <xsl:variable name="temp">
              <xsl:for-each select="./ancestor::w:p//w:instrText">
                <xsl:value-of select ="normalize-space(.)"/>
              </xsl:for-each>
            </xsl:variable>
            <xsl:variable name="type">
              <xsl:choose>
                <!--<xsl:when test="contains($temp,' ')">
                      <xsl:value-of select="substring-before($temp,' ')"/>
                    </xsl:when>-->
                <!--CXL,2012.3.17-->
                <xsl:when test="starts-with($temp,'eq')">
                  <xsl:value-of select="'eq'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'AUTHOR')">
                  <xsl:value-of select="'AUTHOR'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'DATE')">
                  <xsl:value-of select="'DATE'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'FILENAME')">
                  <xsl:value-of select="'FILENAME'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'SECTIONPAGES')">
                  <xsl:value-of select="'SECTIONPAGES'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'REVNUM')">
                  <xsl:value-of select="'REVNUM'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'TIME')">
                  <xsl:value-of select="'TIME'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'PAGE')">
                  <xsl:value-of select="'PAGE'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'SECTION')">
                  <xsl:value-of select="'SECTION'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'REF')">
                  <xsl:value-of select="'REF'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'XE')">
                  <xsl:value-of select="'XE'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'SEQ')">
                  <xsl:value-of select="'SEQ'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'SAVEDATE')">
                  <xsl:value-of select="'SAVEDATE'"/>
                </xsl:when>
                <xsl:when test="contains($temp,'TITLE')">
                  <xsl:value-of select="'TITLE'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'CREATEDATE')">
                  <xsl:value-of select="'CREATEDATE'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'NUMPAGES')">
                  <xsl:value-of select="'NUMPAGES'"/>
                </xsl:when>
                <xsl:when test="starts-with($temp,'NUMCHARS')">
                  <xsl:value-of select="'NUMCHARS'"/>
                </xsl:when>
                <xsl:when test="contains($temp,'HYPERLINK')">
                  <xsl:value-of select="'HYPERLINK'"/>
                </xsl:when>
                <!--<xsl:otherwise>
                      <xsl:value-of select="$temp"/>
                    </xsl:otherwise>-->
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when
                test="$type='AUTHOR' or $type='DATE' or $type='FILENAME' or $type='SECTIONPAGES' or $type='REVNUM' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE' or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS'">
                <xsl:call-template name="SimpleField">
                  <xsl:with-param name="type"
                    select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                  <xsl:with-param name="splfldPartFrom" select="$sdtPartFrom"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$type='HYPERLINK'">
                <xsl:call-template name="hyperlinkRegion">
                  <xsl:with-param name="filename" select="$sdtPartFrom"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$type='eq'">
                <xsl:call-template name="EQfield">
                  <xsl:with-param name="type"
                    select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                  <xsl:with-param name="splfldPartFrom" select="$sdtPartFrom"/>
                </xsl:call-template>
              </xsl:when>
              <!--<xsl:when test="$type='REVNUM'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'revision'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$type='SECTIONPAGES'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'pageinsection'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>-->
            </xsl:choose>
          </xsl:if>
          <xsl:if test="./w:fldChar/@w:fldCharType='end'">
            <字:域结束_41A0/>
          </xsl:if>
        </xsl:if>
      </xsl:if>
      <xsl:if test="name(.)='w:sdt'">
        <xsl:call-template name="sdtContentRun">
          <xsl:with-param name="sdtPartFrom" select="$sdtPartFrom"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="sdtContentCell">
    
    <!--2013-03-25，wudi，修复表格边框转换错误，start-->
    <xsl:param name="postr"/>
    <xsl:param name ="postd"/>
    <xsl:param name="tcPartFrom"/>
    <xsl:for-each select="./w:sdtContent/node()">
      <xsl:if test="name(.)='w:tc'">
        <xsl:call-template name="tblCell">
          <xsl:with-param name ="postr" select ="$postr"/>
          <xsl:with-param name ="postd" select ="$postd"/>
          <xsl:with-param name ="tcPartFrom" select ="$tcPartFrom"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="name(.)='w:sdt'">
        <xsl:call-template name="sdtContentCell"/>
      </xsl:if>
    </xsl:for-each>
    <!--end-->
    
  </xsl:template>

  <xsl:template name="sdtContentRow">
    <xsl:for-each select="./w:sdtContent/node()">
      <xsl:if test="name(.)='w:tr'">
        <xsl:call-template name="tblRow"/>
      </xsl:if>
      <xsl:if test="name(.)='w:sdt'">
        <xsl:call-template name="sdtContentRow"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

<xsl:template name="CustomXmlBlock">
  <xsl:for-each select="./node()">
    <xsl:if test="name(.)='w:p'">
      <xsl:call-template name="paragraph"/>
    </xsl:if>
    <xsl:if test="name(.)='w:tbl'">
      <xsl:call-template name="table"/>
    </xsl:if>
    <xsl:if test="name(.)='w:customXml'">
      <xsl:call-template name="CustomXmlBlock"/>
    </xsl:if>
  </xsl:for-each>
</xsl:template>

<xsl:template name="CustomXmlRun">
  <xsl:for-each select="./node()">
    <xsl:if test="name(.)='w:r'">
      <字:句_419D>
        <xsl:call-template name="run">
          <xsl:with-param name ="rPartFrom" select ="'CustomXml'"/>
        </xsl:call-template>
      </字:句_419D>
    </xsl:if>
    <xsl:if test="name(.)='w:customXml'">
      <xsl:call-template name="CustomXmlRun"/>
    </xsl:if>
  </xsl:for-each>
</xsl:template>

<xsl:template name="CustomXmlCell">
  <xsl:for-each select="./node()">
    <xsl:if test="name(.)='w:tc'">
      <xsl:call-template name="tblCell"/>
    </xsl:if>
    <xsl:if test="name(.)='w:customXml'">
      <xsl:call-template name="CustomXmlCell"/>
    </xsl:if>
  </xsl:for-each>
</xsl:template>

<xsl:template name="CustomXmlRow">
  <xsl:for-each select="./node()">
    <xsl:if test="name(.)='w:tr'">
      <xsl:call-template name="tblRow"/>
    </xsl:if>
    <xsl:if test="name(.)='w:customXml'">
      <xsl:call-template name="CustomXmlRow"/>
    </xsl:if>
  </xsl:for-each>
</xsl:template>

  <xsl:template name="shd">
    <图:图案_800A>
      <xsl:choose>

        <xsl:when test="(./@w:val='clear') and (./@w:color='auto')">
          <xsl:if test="./@w:fill">
            <xsl:attribute name="背景色_800C">
              <xsl:if test="./@w:fill!='auto'">
                <xsl:value-of select="concat('#',./@w:fill)"/>
              </xsl:if>
              <xsl:if test="./@w:fill='auto'">auto</xsl:if>
            </xsl:attribute>
          </xsl:if>
        </xsl:when>

        <xsl:when test="./@w:val='solid'">
          <xsl:if test="./@w:color">
            <xsl:attribute name="前景色_800B">
              <xsl:if test="./@w:color!='auto'">
                <xsl:value-of select="concat('#',./@w:color)"/>
              </xsl:if>
              <xsl:if test="./@w:color='auto'">auto</xsl:if>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="./@w:fill">
            <xsl:attribute name="背景色_800C">
              <xsl:if test="./@w:fill!='auto'">
                <xsl:value-of select="concat('#',./@w:fill)"/>
              </xsl:if>
              <xsl:if test="./@w:fill='auto'">auto</xsl:if>
            </xsl:attribute>
          </xsl:if>
        </xsl:when>

        <xsl:otherwise>
          <xsl:if test="./@w:val">
            <xsl:attribute name="类型_8008">
              <xsl:choose>
                <xsl:when test="./@w:val='pct5'">ptn001</xsl:when>
                <xsl:when test="./@w:val='pct10'">ptn002</xsl:when>
                <xsl:when test="./@w:val='pct12'">ptn002</xsl:when>
                <xsl:when test="./@w:val='pct15'">ptn002</xsl:when>
                <xsl:when test="./@w:val='pct20'">ptn003</xsl:when>
                <xsl:when test="./@w:val='pct25'">ptn004</xsl:when>
                <xsl:when test="./@w:val='pct30'">ptn005</xsl:when>
                <xsl:when test="./@w:val='pct35'">ptn005</xsl:when>
                <xsl:when test="./@w:val='pct37'">ptn005</xsl:when>
                <xsl:when test="./@w:val='pct40'">ptn006</xsl:when>
                <xsl:when test="./@w:val='pct45'">ptn006</xsl:when>
                <xsl:when test="./@w:val='pct50'">ptn007</xsl:when>
                <xsl:when test="./@w:val='pct55'">ptn007</xsl:when>
                <xsl:when test="./@w:val='pct60'">ptn008</xsl:when>
                <xsl:when test="./@w:val='pct62'">ptn008</xsl:when>
                <xsl:when test="./@w:val='pct65'">ptn008</xsl:when>
                <xsl:when test="./@w:val='pct70'">ptn009</xsl:when>
                <xsl:when test="./@w:val='pct75'">ptn010</xsl:when>
                <xsl:when test="./@w:val='pct80'">ptn011</xsl:when>
                <xsl:when test="./@w:val='pct85'">ptn012</xsl:when>
                <xsl:when test="./@w:val='pct87'">ptn012</xsl:when>
                <xsl:when test="./@w:val='pct90'">ptn012</xsl:when>
                <xsl:when test="./@w:val='pct95'">ptn012</xsl:when>
                <xsl:when test="./@w:val='horzStripe'">ptn024</xsl:when>
                <xsl:when test="./@w:val='vertStripe'">ptn023</xsl:when>
                <xsl:when test="./@w:val='reverseDiagStripe'">ptn015</xsl:when>
                <xsl:when test="./@w:val='diagStripe'">ptn016</xsl:when>
                <xsl:when test="./@w:val='horzCross'">ptn045</xsl:when>
                <xsl:when test="./@w:val='diagCross'">ptn041</xsl:when>
                <xsl:when test="./@w:val='thinHorzStripe'">ptn020</xsl:when>
                <xsl:when test="./@w:val='thinVertStripe'">ptn019</xsl:when>
                <xsl:when test="./@w:val='thinReverseDiagStripe'">ptn013</xsl:when>
                <xsl:when test="./@w:val='thinDiagStripe'">ptn014</xsl:when>
                <xsl:when test="./@w:val='thinHorzCross'">ptn043</xsl:when>
                <xsl:when test="./@w:val='thinDiagCross'">ptn047</xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="./@w:color">
            <xsl:attribute name="前景色_800B">
              <xsl:if test="./@w:color!='auto'">
                <xsl:value-of select="concat('#',./@w:color)"/>
              </xsl:if>
              <xsl:if test="./@w:color='auto'">auto</xsl:if>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="./@w:fill">
            <xsl:attribute name="背景色_800C">
              <xsl:if test="./@w:fill!='auto'">
                <xsl:value-of select="concat('#',./@w:fill)"/>
              </xsl:if>
              <xsl:if test="./@w:fill='auto'">auto</xsl:if>
            </xsl:attribute>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </图:图案_800A>
  </xsl:template>
  
</xsl:stylesheet>
