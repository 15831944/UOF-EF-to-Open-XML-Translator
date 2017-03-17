<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
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
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">

  <xsl:import href="table.xsl"/>
  <xsl:import href="paragraph.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>

  <xsl:template name="footnote">
    <w:footnotes>
      <!-- special footnotes -->
      <w:footnote w:type="separator" w:id="0">
        <w:p>
          <!-- there are no styles for endnotes separator in oo -->
          <!--w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr-->
          <w:r>
            <w:separator/>
          </w:r>
        </w:p>
      </w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="1">
        <w:p>
          <!--w:pPr>
            <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr-->
          <w:r>
            <w:continuationSeparator/>
          </w:r>
        </w:p>
      </w:footnote>

      <!-- normal footnotes -->
      <xsl:for-each select="//字:句_419D/字:脚注_4159">
        <w:footnote>
          <xsl:variable name="footnoteNo">
            <xsl:number count="字:脚注_4159" level="any" format="1"/>
          </xsl:variable>

          <xsl:attribute name="w:id">
            <xsl:call-template name="addone">
              <xsl:with-param name="temp" select="$footnoteNo"/>
            </xsl:call-template>
          </xsl:attribute>

          <xsl:for-each select="node()">
            <xsl:choose>
              <xsl:when test="name(.)='字:段落_416B'">
                <xsl:call-template name="paragraph"/>
              </xsl:when>
              <xsl:when test="name(.)='字:文字表_416C'">
                <xsl:call-template name="table"/>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>

        </w:footnote>
      </xsl:for-each>
    </w:footnotes>
  </xsl:template>

  <xsl:template name="endnote">
    <w:endnotes>
      <!-- special endnotes -->
      <w:endnote w:type="separator" w:id="0">
        <w:p>
          <w:r>
            <w:separator/>
          </w:r>
        </w:p>
      </w:endnote>
      <w:endnote w:type="continuationSeparator" w:id="1">
        <w:p>
          <w:r>
            <w:continuationSeparator/>
          </w:r>
        </w:p>
      </w:endnote>
      <!-- normal endnotes -->
      <xsl:for-each select="//字:句_419D/字:尾注_415A">
        <w:endnote>
          <xsl:variable name="endnoteNo">
            <xsl:number count="字:尾注_415A" level="any" format="1"/>
          </xsl:variable>
          <xsl:attribute name="w:id">
            <xsl:call-template name="addone">
              <xsl:with-param name="temp" select="$endnoteNo"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:for-each select="node()">
            <xsl:choose>
              <xsl:when test="name(.)='字:段落_416B'">
                <xsl:call-template name="paragraph"/>
              </xsl:when>
              <xsl:when test="name(.)='字:文字表_416C'">
                <xsl:call-template name="table"/>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </w:endnote>
      </xsl:for-each>
    </w:endnotes>
  </xsl:template>

  <xsl:template name="addone">
    <xsl:param name="temp"/>
    <xsl:value-of select="number($temp)+1"/>
  </xsl:template>
</xsl:stylesheet>
