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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
  <xsl:template match="/">
    <w:document>
      <xsl:if test ="w:document/w:background">
        <xsl:copy-of select ="w:document/w:background"/>
      </xsl:if>
      <xsl:apply-templates select="w:document/w:body" mode="pretreat"/>
    </w:document>
  </xsl:template>
  <xsl:template match="w:body" mode="pretreat">
    <w:body>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="prepara"/>
          </xsl:when>
          <!--杨晓，批注结束,列的后面，2010.1.27-->
          <xsl:when test="name(.)='w:tbl'">
            <w:tbl>
              <xsl:for-each select="node()">
                <xsl:choose>
                  <xsl:when test="name(.)='w:tr'">

                    <xsl:call-template name="trpre"/>

                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:copy-of select="self::*"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </w:tbl>
          </xsl:when>
          <!--杨晓，批注结束,列的后面，增加上面情况2010.1.27-->
          <xsl:otherwise>
            <xsl:copy-of select="self::*"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>


    </w:body>
  </xsl:template>
  <!--杨晓，增加模板，2010.1.27-->
  <xsl:template name="trpre">
    <w:tr>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name()='w:tc' and name(./following-sibling::*[1])='w:commentRangeEnd'">
            <xsl:call-template name="commentendpre"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="self::*"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </w:tr>
  </xsl:template>
  <!--杨晓，增加模板，2010.1.28-->
  <xsl:template name="commentendpre">
    <w:tc>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <w:p>
              <xsl:copy-of select="./@*"/>

              <!--2014-05-05，wudi，修复表格-单元格段落属性转换丢失的BUG，start-->
              <xsl:copy-of select="./w:pPr"/>
              <!--end-->

              <xsl:copy-of select="./following::*[1]"/>
            </w:p>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="self::*"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </w:tc>
  </xsl:template>
  <xsl:template name="prepara">
    <!--current node_____w:p-->
    <xsl:if test="w:pPr/w:sectPr">
      <w:p>
        <w:pPr>
          <xsl:apply-templates select="w:pPr" mode="pretreat1"/>
        </w:pPr>
        <xsl:copy-of select="*[not(self::w:pPr)]"/>
      </w:p>
      <xsl:copy-of select="./w:pPr/w:sectPr"/>
    </xsl:if>
    <xsl:if test="not(w:pPr/w:sectPr)">
      <xsl:copy-of select="."/>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:pPr" mode="pretreat1">
    <xsl:copy-of select="*[not(self::w:sectPr)]"/>
  </xsl:template>
</xsl:stylesheet>