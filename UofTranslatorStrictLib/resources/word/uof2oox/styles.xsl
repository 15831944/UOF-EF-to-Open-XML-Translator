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
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="object.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <!--cxl,2012.1.12-->
  <xsl:template name="styles">
    <w:styles>
      <w:docDefaults><!--这里的问题很奇怪-->
        <xsl:variable name="paragraphDefaultStyle"
          select="/uof:UOF/uof:式样集/式样:段落式样集_9911/式样:段落式样_9912[@类型_4102='custom']"/>
        <w:rPrDefault>
          <!--<xsl:choose>
            <xsl:when test="$paragraphDefaultStyle/字:句属性_4158/node()">
              <xsl:apply-templates select="/uof:UOF/uof:式样集//式样:段落式样_9912[@类型_4102='custom']/字:句属性_4158"
                mode="rpr"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:for-each select="/uof:UOF/uof:式样集//式样:句式样_9910[@类型_4102='custom']">
                <xsl:call-template name="RunProperties"/>
              </xsl:for-each>
            </xsl:otherwise>
          </xsl:choose>-->
          <w:rPr>

            <!--2014-04-14，wudi，修复反方向文档默认字体转换BUG，start-->
            <xsl:for-each select="/uof:UOF/uof:式样集/式样:句式样集_990F/式样:句式样_9910[@类型_4102='default']">
              <xsl:choose>
                <xsl:when test ="./字:字体_4128">
                  <xsl:apply-templates select="字:字体_4128"/>
                </xsl:when>
                <xsl:otherwise>
                  <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
            <!--end-->
            
            <w:kern w:val="2"/>
            <w:sz w:val="21"/>
            <w:szCs w:val="22"/>
            <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
          </w:rPr>
          
        </w:rPrDefault>
        
        <!-- Default paragraph properties -->
        <w:pPrDefault/>
          <!--<w:pPr>
            <xsl:for-each select="$paragraphDefaultStyle">
              <xsl:call-template name="pPr"/>
            </xsl:for-each>
          </w:pPr>
        </w:pPrDefault>-->
      </w:docDefaults>
      <!-- Insert normal character paragraph table Styles-->
      <xsl:for-each select="/uof:UOF/uof:式样集/*/式样:句式样_9910 | /uof:UOF/uof:式样集/*/式样:段落式样_9912 | /uof:UOF/uof:式样集/*/式样:文字表式样_9918">
        <xsl:choose>
          <xsl:when test="name(.)='式样:段落式样_9912'">
            <xsl:call-template name="style">
              <xsl:with-param name="type" select="'paragraph'"/>
              <xsl:with-param name="IfDefault" select="@类型_4102"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='式样:文字表式样_9918'">
            <xsl:call-template name="style">
              <xsl:with-param name="type" select="'table'"/>
              <xsl:with-param name="IfDefault" select="@类型_4102"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='式样:句式样_9910'">
            <xsl:call-template name="style">
              <xsl:with-param name="type" select="'character'"/>
              <xsl:with-param name="IfDefault" select="@类型_4102"/>
            </xsl:call-template>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:styles>
  </xsl:template>

  <xsl:template name="style">
    <xsl:param name="type"/>
    <xsl:param name="IfDefault"/>
    <w:style>
      <xsl:attribute name="w:type">
        <xsl:value-of select="$type"/>
      </xsl:attribute>
      
      <!--2012-01-18，wudi，修复UOF到OOX方向BUG：句式样，段落式样转换问题，改@类型_4102='default'，原@类型_4102='custom'，start-->
      <xsl:if test="@类型_4102='default'">
        <xsl:attribute name="w:default">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <!--end-->
      
      <xsl:attribute name="w:styleId">
        <xsl:value-of select="@标识符_4100"/>
      </xsl:attribute>
      <xsl:element name="w:name">
        <xsl:attribute name="w:val">
          <xsl:value-of select="@名称_4101"/>
        </xsl:attribute>
      </xsl:element>
      <xsl:element name="w:aliases">
        <xsl:attribute name="w:val">
          <xsl:value-of select="@别名_4103"/>
        </xsl:attribute>
      </xsl:element>
      <xsl:if test="@基式样引用_4104">
        <xsl:element name="w:basedOn">
          <xsl:attribute name="w:val">
            <xsl:value-of select="@基式样引用_4104"/>
          </xsl:attribute>
        </xsl:element>
      </xsl:if>
      <xsl:if test="@后继式样引用_4105">
        <xsl:element name="w:next">
          <xsl:attribute name="w:val">
            <xsl:value-of select="@后继式样引用_4105"/>
          </xsl:attribute>
        </xsl:element>
      </xsl:if>

      <!-- qFormat is used by Mirosoft office -->
      <xsl:element name="w:qFormat"/>

      <!--<xsl:if test="$type='paragraph' and $IfDefault!='default'">-->
      <xsl:if test="$type='paragraph'">
        <w:pPr>
          <xsl:call-template name="pPr"/>
        </w:pPr>
        <xsl:apply-templates select="./字:句属性_4158" mode="rpr"/>
      </xsl:if>
      <!--<xsl:if test="$type='character' and $IfDefault!='default'">-->
      <xsl:if test="$type='character'">
        <xsl:call-template name="RunProperties"/>
      </xsl:if>
      <xsl:if test="$type='table'">
        <w:tblPr>
          <xsl:call-template name="tblPr"/>
        </w:tblPr>
      </xsl:if>
    </w:style>
  </xsl:template>

  <!-- font table-->
  <xsl:template name="fontTable">
    <w:fonts>
      <xsl:for-each select="./uof:式样集/式样:字体集_990C/式样:字体声明_990D">
        <w:font>
          <xsl:attribute name="w:name">
            <xsl:value-of select="@名称_9903"/>
          </xsl:attribute>
        </w:font>
      </xsl:for-each>
    </w:fonts>
  </xsl:template>

</xsl:stylesheet>
