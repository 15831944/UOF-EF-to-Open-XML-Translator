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

<xsl:stylesheet version="1.0"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks" 
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <xsl:import href="common.xsl"/>
  <xsl:import href="numbering.xsl"/>
  <xsl:import href="sectPr.xsl"/>
  <xsl:import href="region.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  

  <!--yx,add the template:hlkstart,current node:字:区域开始，2010.3.2-->
  <xsl:template name="hlkstart">
    <xsl:variable name="hyperid" select="@标识符_4100"/>

    <!--wcz,2013/1/25 ,uof超链接分为目标和书签两种，需要分类讨论-->
    <xsl:variable name="linksource">
      <xsl:choose>
        <xsl:when test="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:目标_AA01">
          <xsl:value-of select="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:目标_AA01"/>
        </xsl:when>
        <xsl:when test="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:书签_AA0D">
          <xsl:value-of select="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:书签_AA0D"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    
    <!--wcz,2013/1/25 ,删除不必要的代码，uof超链接分为目标和书签两种，需要分类讨论-->
    <xsl:variable name="hlkgoal"><!--cxl,2012.5.6修改，删除开始的引号及空格-->
      <!--<xsl:if test="contains($linksource,'&quot; ')">
        <xsl:value-of select="concat(' HYPERLINK &quot;',substring-after(translate($linksource,'&quot;',''),' '),'&quot;')"/>
      </xsl:if>
      <xsl:if test="not(contains($linksource,'&quot; '))">
        <xsl:if test="not(contains($linksource,'&quot;'))">
          <xsl:value-of select="concat(' HYPERLINK &quot;',$linksource,'&quot;')"/>
        </xsl:if>
        <xsl:if test="contains($linksource,'&quot;')">
          <xsl:value-of select="concat(' HYPERLINK &quot;',translate($linksource,'&quot;',''),'&quot;')"/>
        </xsl:if>
      </xsl:if>-->
      <xsl:choose>
        <xsl:when test="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:目标_AA01">

          <!--2013-04-25，wudi，修复UOF到OOX方向超链接转换的BUG，电子邮件地址，及文件路径问题，start-->
          <xsl:choose>
            <xsl:when test ="contains($linksource,'mailto:mailto')">
              <xsl:value-of select ="concat(' HYPERLINK &quot;',substring-after($linksource,'mailto:'),'&quot; ')"/>
            </xsl:when>
            <xsl:when test ="contains($linksource,':\') and not(contains($linksource,':\\'))">
              <xsl:call-template name ="pathOutput">
                <xsl:with-param name ="str1" select ="$linksource"/>
                <xsl:with-param name ="str2" select ="''"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat(' HYPERLINK &quot;',$linksource,'&quot; ')"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->

        </xsl:when>
        <xsl:when test="preceding::超链:超级链接_AA0C[超链:链源_AA00 = $hyperid]/超链:书签_AA0D">
          <xsl:value-of  select="concat(' HYPERLINK \l &quot;',$linksource,'&quot; ')"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    
    <w:r>
      <w:fldChar>
        <xsl:attribute name="w:fldCharType">
          <xsl:value-of select="'begin'"/>
        </xsl:attribute>
      </w:fldChar>
    </w:r>
    <w:r>
      <w:instrText>
        <xsl:value-of select="$hlkgoal"/>
      </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar>
        <xsl:attribute name="w:fldCharType">
          <xsl:value-of select="'separate'"/>
        </xsl:attribute>
      </w:fldChar>
    </w:r>
  </xsl:template>

  <!--2013-04-25，wudi，修复UOF到OOX方向超链接转换的BUG，电子邮件地址，及文件路径问题，start-->
  <xsl:template name ="pathOutput">
    <xsl:param name ="str1"/>
    <xsl:param name ="str2"/>
    <xsl:choose>
      <xsl:when test ="contains($str1,'\')">
        <xsl:variable name ="tmp1">
          <xsl:value-of select ="substring-before($str1,'\')"/>
        </xsl:variable>
        <xsl:variable name ="tmp2">
          <xsl:value-of select ="substring-after($str1,'\')"/>
        </xsl:variable>
        <xsl:variable name ="tmp3">
          <xsl:value-of select ="concat($str2,$tmp1,'\\')"/>
        </xsl:variable>
        <xsl:call-template name ="pathOutput">
          <xsl:with-param name ="str1" select ="$tmp2"/>
          <xsl:with-param name ="str2" select ="$tmp3"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="concat(' HYPERLINK &quot;',$str2,$str1,'&quot; ')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--end-->
  
  <!--yx,add the template:hlkend,2010.3.2-->
  <xsl:template name="hlkend">
    <w:r>
      <w:fldChar>
        <xsl:attribute name="w:fldCharType">
          <xsl:value-of select="'end'"/>
        </xsl:attribute>
      </w:fldChar>
    </w:r>
  </xsl:template>


</xsl:stylesheet>
