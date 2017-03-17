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
-->

<xsl:stylesheet version="2.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:fo="http://www.w3.org/1999/XSL/Format" 
xmlns:xs="http://www.w3.org/2001/XMLSchema" 
xmlns:fn="http://www.w3.org/2005/xpath-functions" 
xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
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
  
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <!--CXL修改于2011/11/12-->
  <xsl:template name="header">
    <字:页眉_41F3>
      <xsl:call-template name="headerRef">
        <xsl:with-param name="type" select="'default'"/>
        <xsl:with-param name="typeName" select="'字:奇数页页眉_41F4'"/>
        <!--<xsl:with-param name="id" select="'t0028'"/>-->
      </xsl:call-template>
      <xsl:call-template name="headerRef">
        <xsl:with-param name="type" select="'even'"/>
        <xsl:with-param name="typeName" select="'字:偶数页页眉_41F5'"/>
        <!--<xsl:with-param name="id" select="'t0029'"/>-->
      </xsl:call-template>
      <xsl:call-template name="headerRef">
        <xsl:with-param name="type" select="'first'"/>
        <xsl:with-param name="typeName" select="'字:首页页眉_41F6'"/>
        <!--<xsl:with-param name="id" select="'t0030'"/>-->
      </xsl:call-template>
    </字:页眉_41F3>
  </xsl:template>

  <xsl:template name="headerRef">
    <xsl:param name="type"/>
    <xsl:param name="typeName"/>
    <!--<xsl:param name="id"/>-->

    <xsl:for-each select="./w:headerReference[@w:type=$type]">
      <xsl:variable name="temp" select="@r:id"/>
      <xsl:variable name="target" select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$temp]/@Target"/>

      <xsl:apply-templates select="document(concat('word/',$target))/w:hdr">
        <xsl:with-param name="type" select="$typeName"/>
        <!--<xsl:with-param name="id" select="$id"/>-->
        <xsl:with-param name ="doc" select ="$target"/>
      </xsl:apply-templates>      
    </xsl:for-each>
  </xsl:template>


  <xsl:template match="w:hdr">
    <xsl:param name="type"/>
    <!--<xsl:param name="id"/>-->
    <xsl:param name ="doc"/>
    <xsl:element name="{$type}">
      <!--<xsl:attribute name="uof:locID">
        <xsl:value-of select ="$id"/>
      </xsl:attribute>-->
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name ="pPartFrom" select ="$doc"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table">
              <xsl:with-param name ="tblPartFrom" select ="$doc"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:sdt'">
            
            <!--2013-03-22，wudi，修复页眉转换的BUG，此情况之前处理不正确，会丢失页眉，start-->
            
            <!--2014-04-14，wudi，修复页码丢失的BUG，sdt节点下面还可能存在sdt节点，start-->
            <xsl:for-each select="w:sdtContent/w:p | w:sdtContent/w:sdt">
              <!--<xsl:choose>
                <xsl:when test ="name(.)='w:sdtContent'">-->
                  <!--<xsl:for-each select="w:p | w:sdt">-->
                    <xsl:choose>
                      <xsl:when test ="name(.) = 'w:p'">
                        <xsl:call-template name="paragraph">
                          <xsl:with-param name ="pPartFrom" select ="$doc"/>
                        </xsl:call-template>
                      </xsl:when>

                      <xsl:when test="name(.) = 'w:sdt'">
                        <xsl:for-each select="w:sdtContent/w:p">
                          <!--<xsl:choose>
                            <xsl:when test="name(.) = 'w:sdtContent'">-->
                              <!--<xsl:for-each select="w:p">-->
                                <!--<xsl:choose>
                                  <xsl:when test="name(.) = 'w:p'">-->
                                    <xsl:call-template name="paragraph">
                                      <xsl:with-param name ="pPartFrom" select ="$doc"/>
                                    </xsl:call-template>
                                  <!--</xsl:when>
                                </xsl:choose>
                              </xsl:for-each>
                            </xsl:when>
                          </xsl:choose>-->
                        </xsl:for-each>
                      </xsl:when>
                    </xsl:choose>
                  <!--</xsl:for-each>
                </xsl:when>
              </xsl:choose>-->
            </xsl:for-each>
            <!--end-->
            
            <!--end-->
            
          </xsl:when>
          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlBlock"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </xsl:element>

  </xsl:template>

  <xsl:template name="footer">
    <字:页脚_41F7>
      <xsl:call-template name="footerRef">
        <xsl:with-param name="type" select="'default'"/>
        <xsl:with-param name="typeName" select="'字:奇数页页脚_41F8'"/>
        <!--<xsl:with-param name="id" select="'t0032'"/>-->
      </xsl:call-template>
      <xsl:call-template name="footerRef">
        <xsl:with-param name="type" select="'even'"/>
        <xsl:with-param name="typeName" select="'字:偶数页页脚_41F9'"/>
        <!--<xsl:with-param name="id" select="'t0033'"/>-->
      </xsl:call-template>
      <xsl:call-template name="footerRef">
        <xsl:with-param name="type" select="'first'"/>
        <xsl:with-param name="typeName" select="'字:首页页脚_41FA'"/>
        <!--<xsl:with-param name="id" select="'t0034'"/>-->
      </xsl:call-template>
    </字:页脚_41F7>
  </xsl:template>

  <xsl:template name="footerRef">
    <xsl:param name="type"/>
    <xsl:param name="typeName"/>
    <!--<xsl:param name="id"/>-->

    <xsl:for-each select="./w:footerReference[@w:type=$type] ">
      <xsl:variable name="temp" select="@r:id"/>
      <xsl:variable name="target" select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$temp]/@Target"/>

      <xsl:apply-templates select="document(concat('word/',$target))/w:ftr">
        <xsl:with-param name="type" select="$typeName"/>
        <!--<xsl:with-param name="id" select="$id"/>-->
        <xsl:with-param name="doc" select="$target"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="w:ftr">
    <xsl:param name="type"/>
    <!--<xsl:param name="id"/>-->
    <xsl:param name="doc"/>
    <xsl:element name="{$type}">
      <!--<xsl:attribute name="uof:locID">
        <xsl:value-of select ="$id"/>
      </xsl:attribute>-->
      <xsl:for-each select="w:p | w:tbl | w:sdt | w:customXml">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name ="pPartFrom" select ="$doc"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table">
              <xsl:with-param name ="tblPartFrom" select ="$doc"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:sdt'">
            
            <!--2013-03-22，wudi，修复页脚转换的BUG，此情况之前处理不正确，会丢失页脚，start-->
            
            <!--2014-04-14，wudi，修复页码丢失的BUG，sdt节点下面还可能存在sdt节点，start-->
            <xsl:for-each select="w:sdtContent/w:p | w:sdtContent/w:sdt">
              <!--<xsl:choose>
                <xsl:when test ="name(.)='w:sdtContent'">-->
                  <!--<xsl:for-each select="w:p | w:sdt">-->
                    <xsl:choose>
                      <xsl:when test ="name(.) = 'w:p'">
                        <xsl:call-template name="paragraph">
                          <xsl:with-param name ="pPartFrom" select ="$doc"/>
                        </xsl:call-template>
                      </xsl:when>
                                            
                      <xsl:when test="name(.) = 'w:sdt'">
                        <xsl:for-each select="w:sdtContent/w:p">
                          <!--<xsl:choose>
                            <xsl:when test="name(.) = 'w:sdtContent'">-->
                              <!--<xsl:for-each select="w:p">
                                <xsl:choose>
                                  <xsl:when test="name(.) = 'w:p'">-->
                                    <xsl:call-template name="paragraph">
                                      <xsl:with-param name ="pPartFrom" select ="$doc"/>
                                    </xsl:call-template>
                                  <!--</xsl:when>
                                </xsl:choose>
                              </xsl:for-each>-->
                            <!--</xsl:when>
                          </xsl:choose>-->
                        </xsl:for-each>
                      </xsl:when>
                    </xsl:choose>
                  <!--</xsl:for-each>-->
                <!--</xsl:when>
              </xsl:choose>-->
            </xsl:for-each>
            <!--end-->
            
            <!--end-->
            
          </xsl:when>
          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlBlock"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </xsl:element>
  </xsl:template>

</xsl:stylesheet>
