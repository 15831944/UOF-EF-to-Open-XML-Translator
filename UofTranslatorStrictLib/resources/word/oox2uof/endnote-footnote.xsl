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
<Author>Ban Qianchao(BUAA)</Author>
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
  
  <xsl:import href="table.xsl"/>
  <xsl:import href="paragraph.xsl"/>
  
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  
  <xsl:template name="endnoteReference">
    <字:尾注_415A>
      <xsl:if test="@w:customMarkFollows='1' or @w:customMarkFollows='on' or @w:customMarkFollows='true'">
        <xsl:attribute name="引文体_4157">

          <!--2013-03-12，wudi，修复自定义尾注BUG，部分枚举，start-->
          <xsl:choose>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F021']">
              <xsl:value-of select ="'!'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F023']">
              <xsl:value-of select ="'#'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F025']">
              <xsl:value-of select ="'%'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F028']">
              <xsl:value-of select ="'('"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F029']">
              <xsl:value-of select ="')'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02A']">
              <xsl:value-of select ="'*'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02B']">
              <xsl:value-of select ="'+'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02C']">
              <xsl:value-of select ="','"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02D']">
              <xsl:value-of select ="'-'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02E']">
              <xsl:value-of select ="'.'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02F']">
              <xsl:value-of select ="'/'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F030']">
              <xsl:value-of select ="'0'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F031']">
              <xsl:value-of select ="'1'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F032']">
              <xsl:value-of select ="'2'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F033']">
              <xsl:value-of select ="'3'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F034']">
              <xsl:value-of select ="'4'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F035']">
              <xsl:value-of select ="'5'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F036']">
              <xsl:value-of select ="'6'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F037']">
              <xsl:value-of select ="'7'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F038']">
              <xsl:value-of select ="'8'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F039']">
              <xsl:value-of select ="'9'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03A']">
              <xsl:value-of select ="':'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03B']">
              <xsl:value-of select ="';'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03D']">
              <xsl:value-of select ="'='"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03E']">
              <xsl:value-of select ="'>'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03F']">
              <xsl:value-of select ="'?'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F041']">
              <xsl:value-of select ="'Α'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F042']">
              <xsl:value-of select ="'Β'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F043']">
              <xsl:value-of select ="'Χ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F044']">
              <xsl:value-of select ="'Δ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F045']">
              <xsl:value-of select ="'Ε'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F046']">
              <xsl:value-of select ="'Φ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F047']">
              <xsl:value-of select ="'Γ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F048']">
              <xsl:value-of select ="'Η'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F049']">
              <xsl:value-of select ="'Ι'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04B']">
              <xsl:value-of select ="'Κ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04C']">
              <xsl:value-of select ="'Λ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04D']">
              <xsl:value-of select ="'Μ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04E']">
              <xsl:value-of select ="'Ν'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04F']">
              <xsl:value-of select ="'Ο'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F050']">
              <xsl:value-of select ="'Π'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F051']">
              <xsl:value-of select ="'Θ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F052']">
              <xsl:value-of select ="'Ρ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F053']">
              <xsl:value-of select ="'Σ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F054']">
              <xsl:value-of select ="'Τ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F055']">
              <xsl:value-of select ="'Υ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F057']">
              <xsl:value-of select ="'Ω'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F058']">
              <xsl:value-of select ="'Ξ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F059']">
              <xsl:value-of select ="'Ψ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05A']">
              <xsl:value-of select ="'Ζ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F061']">
              <xsl:value-of select ="'α'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F062']">
              <xsl:value-of select ="'β'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F063']">
              <xsl:value-of select ="'χ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F064']">
              <xsl:value-of select ="'δ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F065']">
              <xsl:value-of select ="'ε'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F066']">
              <xsl:value-of select ="'φ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F067']">
              <xsl:value-of select ="'γ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F068']">
              <xsl:value-of select ="'η'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F069']">
              <xsl:value-of select ="'ι'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06B']">
              <xsl:value-of select ="'κ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06C']">
              <xsl:value-of select ="'λ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06D']">
              <xsl:value-of select ="'μ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06E']">
              <xsl:value-of select ="'ν'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06F']">
              <xsl:value-of select ="'ο'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F070']">
              <xsl:value-of select ="'π'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F071']">
              <xsl:value-of select ="'θ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F072']">
              <xsl:value-of select ="'ρ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F073']">
              <xsl:value-of select ="'σ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F074']">
              <xsl:value-of select ="'τ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F075']">
              <xsl:value-of select ="'υ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F076']">
              <xsl:value-of select ="'ϖ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F077']">
              <xsl:value-of select ="'ω'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F078']">
              <xsl:value-of select ="'ξ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F079']">
              <xsl:value-of select ="'ψ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07A']">
              <xsl:value-of select ="'ζ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A7']">
              <xsl:value-of select ="'♣'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A8']">
              <xsl:value-of select ="'♦'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A9']">
              <xsl:value-of select ="'♥'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AA']">
              <xsl:value-of select ="'♠'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AC']">
              <xsl:value-of select ="'←'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AD']">
              <xsl:value-of select ="'↑'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AE']">
              <xsl:value-of select ="'→'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AF']">
              <xsl:value-of select ="'↓'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AB']">
              <xsl:value-of select ="'↔'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05B']">
              <xsl:value-of select ="'['"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05D']">
              <xsl:value-of select ="']'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07B']">
              <xsl:value-of select ="'{'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07C']">
              <xsl:value-of select ="'|'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07D']">
              <xsl:value-of select ="'}'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07E']">
              <xsl:value-of select ="'~'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'sym'"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->
          
          <xsl:if test ="following-sibling::w:t">
            <xsl:value-of select="following-sibling::w:t"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:variable name="endnoteid" select="@w:id"/>
      <xsl:apply-templates select="document('word/endnotes.xml')/w:endnotes/w:endnote[@w:id=$endnoteid]" mode="endnotespart"/>
    </字:尾注_415A>
  </xsl:template>
  
  <xsl:template match="w:endnote" mode="endnotespart">
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name ="pPartFrom" select ="'endnotes'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table">
              <xsl:with-param name ="tblPartFrom" select ="'endnotes'"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    
  </xsl:template>
  
  <xsl:template name="footnoteReference">
    <字:脚注_4159>
      <xsl:if test="@w:customMarkFollows='1' or @w:customMarkFollows='on' or @w:customMarkFollows='true'">
        <xsl:attribute name="引文体_4157">

          <!--2013-03-12，wudi，修复自定义脚注BUG，部分枚举，start-->
          <xsl:choose>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F021']">
              <xsl:value-of select ="'!'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F023']">
              <xsl:value-of select ="'#'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F025']">
              <xsl:value-of select ="'%'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F028']">
              <xsl:value-of select ="'('"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F029']">
              <xsl:value-of select ="')'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02A']">
              <xsl:value-of select ="'*'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02B']">
              <xsl:value-of select ="'+'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02C']">
              <xsl:value-of select ="','"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02D']">
              <xsl:value-of select ="'-'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02E']">
              <xsl:value-of select ="'.'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F02F']">
              <xsl:value-of select ="'/'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F030']">
              <xsl:value-of select ="'0'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F031']">
              <xsl:value-of select ="'1'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F032']">
              <xsl:value-of select ="'2'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F033']">
              <xsl:value-of select ="'3'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F034']">
              <xsl:value-of select ="'4'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F035']">
              <xsl:value-of select ="'5'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F036']">
              <xsl:value-of select ="'6'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F037']">
              <xsl:value-of select ="'7'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F038']">
              <xsl:value-of select ="'8'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F039']">
              <xsl:value-of select ="'9'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03A']">
              <xsl:value-of select ="':'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03B']">
              <xsl:value-of select ="';'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03D']">
              <xsl:value-of select ="'='"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03E']">
              <xsl:value-of select ="'>'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F03F']">
              <xsl:value-of select ="'?'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F041']">
              <xsl:value-of select ="'Α'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F042']">
              <xsl:value-of select ="'Β'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F043']">
              <xsl:value-of select ="'Χ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F044']">
              <xsl:value-of select ="'Δ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F045']">
              <xsl:value-of select ="'Ε'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F046']">
              <xsl:value-of select ="'Φ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F047']">
              <xsl:value-of select ="'Γ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F048']">
              <xsl:value-of select ="'Η'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F049']">
              <xsl:value-of select ="'Ι'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04B']">
              <xsl:value-of select ="'Κ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04C']">
              <xsl:value-of select ="'Λ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04D']">
              <xsl:value-of select ="'Μ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04E']">
              <xsl:value-of select ="'Ν'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F04F']">
              <xsl:value-of select ="'Ο'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F050']">
              <xsl:value-of select ="'Π'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F051']">
              <xsl:value-of select ="'Θ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F052']">
              <xsl:value-of select ="'Ρ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F053']">
              <xsl:value-of select ="'Σ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F054']">
              <xsl:value-of select ="'Τ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F055']">
              <xsl:value-of select ="'Υ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F057']">
              <xsl:value-of select ="'Ω'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F058']">
              <xsl:value-of select ="'Ξ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F059']">
              <xsl:value-of select ="'Ψ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05A']">
              <xsl:value-of select ="'Ζ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F061']">
              <xsl:value-of select ="'α'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F062']">
              <xsl:value-of select ="'β'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F063']">
              <xsl:value-of select ="'χ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F064']">
              <xsl:value-of select ="'δ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F065']">
              <xsl:value-of select ="'ε'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F066']">
              <xsl:value-of select ="'φ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F067']">
              <xsl:value-of select ="'γ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F068']">
              <xsl:value-of select ="'η'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F069']">
              <xsl:value-of select ="'ι'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06B']">
              <xsl:value-of select ="'κ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06C']">
              <xsl:value-of select ="'λ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06D']">
              <xsl:value-of select ="'μ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06E']">
              <xsl:value-of select ="'ν'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F06F']">
              <xsl:value-of select ="'ο'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F070']">
              <xsl:value-of select ="'π'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F071']">
              <xsl:value-of select ="'θ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F072']">
              <xsl:value-of select ="'ρ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F073']">
              <xsl:value-of select ="'σ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F074']">
              <xsl:value-of select ="'τ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F075']">
              <xsl:value-of select ="'υ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F076']">
              <xsl:value-of select ="'ϖ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F077']">
              <xsl:value-of select ="'ω'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F078']">
              <xsl:value-of select ="'ξ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F079']">
              <xsl:value-of select ="'ψ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07A']">
              <xsl:value-of select ="'ζ'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A7']">
              <xsl:value-of select ="'♣'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A8']">
              <xsl:value-of select ="'♦'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0A9']">
              <xsl:value-of select ="'♥'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AA']">
              <xsl:value-of select ="'♠'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AC']">
              <xsl:value-of select ="'←'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AD']">
              <xsl:value-of select ="'↑'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AE']">
              <xsl:value-of select ="'→'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AF']">
              <xsl:value-of select ="'↓'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F0AB']">
              <xsl:value-of select ="'↔'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05B']">
              <xsl:value-of select ="'['"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F05D']">
              <xsl:value-of select ="']'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07B']">
              <xsl:value-of select ="'{'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07C']">
              <xsl:value-of select ="'|'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07D']">
              <xsl:value-of select ="'}'"/>
            </xsl:when>
            <xsl:when test ="following-sibling::w:sym[1][@w:char='F07E']">
              <xsl:value-of select ="'~'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'sym'"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->
            
          <xsl:if test ="following-sibling::w:t">
            <xsl:value-of select="following-sibling::w:t"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:variable name="footnoteid" select="@w:id"/>
      <xsl:apply-templates select="document('word/footnotes.xml')/w:footnotes/w:footnote[@w:id=$footnoteid]" mode="footnotespart"/>
    </字:脚注_4159>
  </xsl:template>
  
  <xsl:template match="w:footnote" mode="footnotespart">
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name ="pPartFrom" select ="'footnotes'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table">
              <xsl:with-param name ="tblPartFrom" select ="'footnotes'"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
