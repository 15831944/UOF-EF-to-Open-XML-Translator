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
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <xsl:template name="package-relationships">
    <Relationships>
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Target="word/document.xml"/>
      <!--cxl2012/1/5修改matadata包联系-->
      <xsl:if
        test=".//元:元数据_5200/元:编辑时间_5209 | .//元:元数据_5200/元:创建应用程序_520A |.//元:元数据_5200/元:文档模板_520C |.//元:元数据_5200/元:公司名称_5213 |.//元:元数据_5200/元:经理名称_5214 |.//元:元数据_5200/元:页数_5215 |.//元:元数据_5200/元:字数_5216 |.//元:元数据_5200/元:行数_5219 |.//元:元数据_5200/元:段落数_521A">
        <Relationship Id="rId3"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
          Target="docProps/app.xml"/>
      </xsl:if>
      <xsl:if
        test=".//元:元数据_5200/元:标题_5201|.//元:元数据_5200/元:主题_5202|.//元:元数据_5200/元:创建者_5203 |.//元:元数据_5200/元:最后作者_5205 |.//元:元数据_5200/元:摘要_5206|.//元:元数据_5200/元:创建日期_5207 |.//元:元数据_5200/元:编辑次数_5208|.//元:元数据_5200/元:分类_520B |.//元:元数据_5200/元:关键字集_520D">
        <Relationship Id="rId2"
          Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
          Target="docProps/core.xml"/>
      </xsl:if>
      <xsl:if test=".//元:元数据_5200/元:用户自定义元数据集_520F">
        <Relationship Id="rId4"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
          Target="docProps/custom.xml"/>
      </xsl:if>
      <!--<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>-->
    </Relationships>
  </xsl:template>
</xsl:stylesheet>
