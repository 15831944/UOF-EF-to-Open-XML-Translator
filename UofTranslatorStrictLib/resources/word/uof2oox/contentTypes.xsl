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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:template name="contentTypes">
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels"
        ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <!--picture-->

      <!--2015-03-23，wudi，改and条件为not(@私有类型_D707 != 'jpeg')，start-->
      <!--2015-03-13，wudi，if条件or改为and，start-->
      <!--2013-04-16，wudi，修复jpeg格式图片转换后文档内容错误的BUG，增加限制条件 @私有类型_D707 != 'jpeg'-->
      <xsl:if test="//对象:对象数据_D701[@公共类型_D706='jpg'] and not(//对象:对象数据_D701[@私有类型_D707 = 'jpeg'])">
        <Default Extension="jpg" ContentType="image/jpeg"/>
      </xsl:if>
      <!--end-->
      <!--end-->
      <!--end-->
      
	    <xsl:if test="//对象:对象数据_D701[@公共类型_D706='JPG']">
		    <Default Extension="JPG" ContentType="image/jpeg"/>
	    </xsl:if>
      <xsl:if test="//对象:对象数据_D701[@公共类型_D706='wmf']">
        <Default Extension="wmf" ContentType="image/x-wmf"/>
      </xsl:if>
      
      <!--2013-05-15，wudi，新增emf图片格式，start-->
      <xsl:if test ="//对象:对象数据_D701[@公共类型_D706='emf'] or //对象:对象数据_D701[@私有类型_D707='emf']">
        <Default Extension="emf" ContentType="image/x-emf"/>
      </xsl:if>
      <!--end-->

      <!--2015-03-23，wudi，新增eof图片格式，start-->
      <xsl:if test ="//对象:对象数据_D701[@公共类型_D706='eof'] or //对象:对象数据_D701[@私有类型_D707='eof']">
        <Default Extension="eof" ContentType="image/x-eof"/>
      </xsl:if>
      <!--end-->
      
      <!--2013-04-16，wudi，修复jpeg格式图片转换后文档内容错误的BUG，增加条件 or //对象:对象数据_D701[@私有类型_D707 = 'jpeg']-->
      <xsl:if test="//对象:对象数据_D701[@公共类型_D706='jpeg'] or //对象:对象数据_D701[@私有类型_D707 = 'jpeg']">
        <Default Extension="jpeg" ContentType="image/jpeg"/>
      </xsl:if>
      <!--end-->
      
      <xsl:if test="//对象:对象数据_D701[@公共类型_D706='png']">
        <Default Extension="png" ContentType="image/png"/>
      </xsl:if>
      <xsl:if test="//对象:对象数据_D701[@公共类型_D706='tmp']">
        <Default Extension="tmp" ContentType="image/png"/>
      </xsl:if>
      
      <!--2014-03-12，wudi，注释掉这段代码，重复，start-->
      
      <!--<xsl:if test="//对象:对象数据_D701[@公共类型_D706='emf']">
        <Default Extension="emf" ContentType="image/x-emf"/>
      </xsl:if>-->
      
      <!--end-->
      
	  <!--<xsl:if test="//对象:对象数据_D701">-->
      <!--<xsl:if test="//对象:对象数据_D701[@公共类型_D706='gif']">-->
      <Default Extension="gif" ContentType="image/gif"/>
      <!--</xsl:if>-->
      <Override PartName="/word/document.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/fontTable.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
      <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
      <!--cxl2012.2.9修改-->
      <xsl:if
        test="uof:元数据/元:元数据_5200/元:编辑时间_5209|uof:元数据/元:元数据_5200/元:创建应用程序_520A |uof:元数据/元:元数据_5200/元:文档模板_520C|uof:元数据/元:元数据_5200/元:公司名称_5213|uof:元数据/元:元数据_5200/元:经理名称_5214|uof:元数据/元:元数据_5200/元:页数_5215|uof:元数据/元:元数据_5200/元:字数_5216|uof:元数据/元:元数据_5200/元:行数_5219|uof:元数据/元:元数据_5200/元:段落数_521A">
        <Override PartName="/docProps/app.xml"
          ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      </xsl:if>
      <xsl:if
        test="uof:元数据/元:元数据_5200/元:标题_5201|uof:元数据/元:元数据_5200/元:主题_5202|uof:元数据/元:元数据_5200/元:创建者_5203|uof:元数据/元:元数据_5200/元:最后作者_5205|uof:元数据/元:元数据_5200/元:摘要_5206|uof:元数据/元:元数据_5200/元:创建日期_5207|uof:元数据/元:元数据_5200/元:编辑次数_5208|uof:元数据/元:元数据_5200/元:分类_520B|uof:元数据/元:元数据_5200/元:关键字集_520D">
        <Override PartName="/docProps/core.xml"
          ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      </xsl:if>
      <xsl:if test="uof:元数据/元:元数据_5200/元:用户自定义元数据集_520F">
        <Override PartName="/docProps/custom.xml"
          ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
      </xsl:if>
      <!--setting-->
      <Override PartName="/word/settings.xml"
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
      <!--numbering-->
      <xsl:if test="uof:式样集//式样:自动编号集_990E">
        <Override PartName="/word/numbering.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
        />
      </xsl:if>
      <!--comments-->
      <xsl:if test="uof:文字处理/规则:公用处理规则_B665/规则:批注集_B669">
        <Override PartName="/word/comments.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
        />
      </xsl:if>
      <!--footnote and endnote-->
      <xsl:if test="//字:句_419D/字:脚注_4159">
        <Override PartName="/word/footnotes.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        />
      </xsl:if>
      <xsl:if test="//字:句_419D/字:尾注_415A">
        <Override PartName="/word/endnotes.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
        />
      </xsl:if>
      <!-- header footer-->
      <xsl:for-each
        select="//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:奇数页页眉_41F4
                           |//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:偶数页页眉_41F5
                           |//字:分节_416A/字:节属性_421B/字:页眉_41F3/字:首页页眉_41F6">
        <Override
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml">
          <xsl:attribute name="PartName">
            <xsl:value-of select="'/word/header'"/>
            <xsl:number count="字:奇数页页眉_41F4|字:偶数页页眉_41F5|字:首页页眉_41F6" level="any" format="1"/>
            <xsl:value-of select="'.xml'"/>
          </xsl:attribute>
        </Override>
      </xsl:for-each>
      <xsl:for-each
        select="//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:奇数页页脚_41F8
                           |//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:偶数页页脚_41F9
                           |//字:分节_416A/字:节属性_421B/字:页脚_41F7/字:首页页脚_41FA">
        <Override
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml">
          <xsl:attribute name="PartName">
            <xsl:value-of select="'/word/footer'"/>
            <xsl:number count="字:奇数页页脚_41F8|字:偶数页页脚_41F9|字:首页页脚_41FA" level="any" format="1"/>
            <xsl:value-of select="'.xml'"/>
          </xsl:attribute>
        </Override>
      </xsl:for-each>

    </Types>
  </xsl:template>
</xsl:stylesheet>
