<?xml version="1.0" encoding="UTF-8" ?>
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
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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
  <!-- header -->
  <xsl:template name="header">
    <w:hdr>
      <xsl:for-each select="字:段落_416B | 字:文字表_416C">
        <xsl:choose>
          <xsl:when test="name(.)='字:段落_416B'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='字:文字表_416C'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:hdr>
  </xsl:template>
  <!-- footer-->
  <xsl:template name="footer">
    <w:ftr>
      <xsl:for-each select="字:段落_416B | 字:文字表_416C">
        <xsl:choose>
          <xsl:when test="name(.)='字:段落_416B'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='字:文字表_416C'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:ftr>
  </xsl:template>
  <!-- sectPr state the odd or even page -->
  <xsl:template name="HeaderFooter">
    <xsl:variable name="path" select="/uof:UOF/uof:文字处理/字:文字处理文档_4225/字:分节_416A/字:节属性_421B"/>
    <xsl:for-each select="./字:页眉_41F3/字:奇数页页眉_41F4|./字:页眉_41F3/字:偶数页页眉_41F5|./字:页眉_41F3/字:首页页眉_41F6">
      <w:headerReference>
        <xsl:attribute name="w:type">
          <xsl:choose>
            <xsl:when test="name(.)='字:奇数页页眉_41F4'">
              <xsl:value-of select="'default'"/>
            </xsl:when>
            <xsl:when test="name(.)='字:偶数页页眉_41F5'">
              <xsl:value-of select="'even'"/>
            </xsl:when>
            <xsl:when test="name(.)='字:首页页眉_41F6'">
              <xsl:value-of select="'first'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="r:id">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </w:headerReference>
    </xsl:for-each>
    <xsl:for-each select="./字:页脚_41F7/字:奇数页页脚_41F8|./字:页脚_41F7/字:偶数页页脚_41F9|./字:页脚_41F7/字:首页页脚_41FA">
      <w:footerReference>
        <xsl:attribute name="w:type">
          <xsl:choose>
            <xsl:when test="name(.)='字:奇数页页脚_41F8'">
              <xsl:value-of select="'default'"/>
            </xsl:when>
            <xsl:when test="name(.)='字:偶数页页脚_41F9'">
              <xsl:value-of select="'even'"/>
            </xsl:when>
            <xsl:when test="name(.)='字:首页页脚_41FA'">
              <xsl:value-of select="'first'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="r:id">
          <xsl:value-of select="generate-id(.)"/>
        </xsl:attribute>
      </w:footerReference>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
