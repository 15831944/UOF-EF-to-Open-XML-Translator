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
<Author>Gu Yueqiong(BUAA)</Author>
-->
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
  <xsl:template match="/">
    <w:document>
      <xsl:if test ="w:document/w:background">
        <xsl:copy-of select ="w:document/w:background"/>
      </xsl:if>
      <xsl:apply-templates select="w:document/w:body" mode="pretreat3"/>
    </w:document>
  </xsl:template>
  <xsl:template match="w:body" mode="pretreat3">
    <w:body>
      <xsl:for-each select="node()">
        <xsl:choose>
          <!--cxl,2012.3.8增加首字下沉预处理代码-->
          <!--<xsl:when test ="name(.)='w:p'">            
            <xsl:if test="./w:pPr/w:framePr">
              <w:p>
                <w:pPr>
                  <xsl:for-each select="./w:pPr/node()">
                    <xsl:if test="name(.)!='w:rPr'">
                      <xsl:copy-of select="."/>
                    </xsl:if>
                  </xsl:for-each>
                  <xsl:copy-of select="./w:r/w:rPr"/>                 
                </w:pPr>                
                <w:r>
                  <xsl:copy-of select="./w:r/w:t"/>
                </w:r>                
                <xsl:copy-of select="following-sibling::w:p[1]/*"/>
              </w:p>
            </xsl:if>
            <xsl:if test="not(./w:pPr/w:framePr) and not(preceding-sibling::w:p[1]/w:pPr/w:framePr)">
              <xsl:copy-of select="."/>
              --><!--<w:p>
                <xsl:variable name="preNo" select="count(preceding-sibling::w:p)"/>
                    <w:pPr>
                      <xsl:apply-templates select="parent::w:body/w:p[$preNo]" mode="PrePpr"/>
                      <xsl:copy-of select="./w:pPr/*"/>
                    </w:pPr>
                      <xsl:apply-templates select="parent::w:body/w:p[$preNo]" mode="PreRun"/>
                    <xsl:copy-of select="./*[not(self::w:pPr)]"/>
              </w:p>--><!--
            </xsl:if>
          </xsl:when>-->
          <!--cxl,2012.3.9新增垂直合并单元格预处理-->
          <xsl:when test="name(.)='w:tbl'">           
            <xsl:if test="not(.//w:vMerge)">
              <xsl:copy-of select="."/>
            </xsl:if>
            <xsl:if test=".//w:vMerge">
              <w:tbl>
                <xsl:for-each select="node()">
                  <xsl:choose>
                    <xsl:when test="name(.)='w:tr'">
                      <w:tr>
                        <xsl:for-each select="node()">                         
                          <xsl:choose>
                            <xsl:when test="name(.)='w:tc'">
                              <xsl:if test="./w:tcPr/w:vMerge/@w:val='restart'">
                                <xsl:variable name="position">
                                  <xsl:value-of select="position()"/>
                                </xsl:variable>
                                <xsl:comment>
                                  <xsl:value-of select="name()"/>
                                </xsl:comment>
                                <xsl:variable name="id">
                                  <xsl:value-of select="generate-id(.//w:vMerge)"/>
                                </xsl:variable>
                                <!--<xsl:variable name="id1">
                                  <xsl:value-of select="generate-id(following::w:tc//w:vMerge[1][@w:val='restart'])"/>
                                </xsl:variable>-->
                                <xsl:variable name="rowsNum"><!--cxl,2012.3.27修改-->
                                  <!--<xsl:value-of select="count(following::w:tr//w:tc[position() = $position]//w:vMerge[not(@w:val='restart') and generate-id(preceding::w:tr//w:tc[position() = $position]//w:vMerge[@w:val='restart']) = $id])"/>-->
                                  <xsl:value-of select="count(following::w:tr/w:tc[position() = $position]//w:vMerge
                                                [not(@w:val='restart') and generate-id(preceding::w:tr/w:tc[position() = $position]//w:vMerge[@w:val='restart'])=$id])"/>
                                </xsl:variable>
                                <w:tc>
                                  <xsl:for-each select="node()">
                                    <xsl:choose>
                                      <xsl:when test="name(.)='w:tcPr'">
                                        <w:tcPr>
                                          <xsl:copy-of select="./*"/>
                                          <w:rowsnum>
                                            <xsl:value-of select="$rowsNum +1"/>
                                          </w:rowsnum>
                                        </w:tcPr>
                                      </xsl:when>
                                      <xsl:otherwise>
                                        <xsl:copy-of select ="."/>
                                      </xsl:otherwise>
                                    </xsl:choose>
                                  </xsl:for-each>
                                </w:tc>                               
                              </xsl:if>
                              <xsl:if test="not(./w:tcPr/w:vMerge)">
                                <xsl:copy-of select="."/>
                              </xsl:if>
                              <xsl:if test="not(./w:tcPr/w:vMerge/@w:val='restart')">
                              </xsl:if>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:copy-of select="."/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:for-each>
                      </w:tr>                     
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:copy-of select="."/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
              </w:tbl>             
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </w:body>
  </xsl:template>

  <xsl:template match="w:p" mode="PrePpr">
    <xsl:if test="./w:pPr/w:framePr">
      <xsl:copy-of select="./w:pPr/w:framePr"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:p" mode="PreRun">
    <xsl:if test="./w:pPr/w:framePr">
      <w:r>
      <xsl:apply-templates select="./w:r" mode="RealRun"/>
      </w:r>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:r" mode="RealRun">
    <xsl:copy-of select="./*[not(self::w:rPr)]"/>
  </xsl:template>
</xsl:stylesheet>