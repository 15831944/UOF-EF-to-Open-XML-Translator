<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:uof="http://schemas.uof.org/cn/2009/uof" 
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
				xmlns:元="http://schemas.uof.org/cn/2009/metadata">
	<xsl:template name="package-relationships">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
			
        <!-- 修改元数据  李娟 11.12.28-->
    
      <!--Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/-->
      <!--2月20日改:处理图片填充时需要,可能有错>
      <xsl:if test="uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:背景/图:图片">
         <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>
      </xsl:if-->
      <!--matadata-->
			<xsl:if test="uof:元数据/元:元数据_5200/元:编辑时间_5209|uof:元数据/元:元数据_5200/元:创建应用程序_520A|uof:元数据/元:元数据_5200/元:文档模板_520C|uof:元数据/元:元数据_5200/元:公司名称_5213|uof:元数据/元:元数据_5200/元:经理名称_5214|uof:元数据/元:元数据_5200/元:页数_5215|uof:元数据/元:元数据_5200/元:字数_5216|uof:元数据/元:元数据_5200/元:行数_5219|uof:元数据/元:元数据_5200/元:段落数_521A">
				<!--<xsl:if test="uof:元数据/uof:编辑时间|uof:元数据/uof:创建应用程序|uof:元数据/uof:文档模板|uof:元数据/uof:公司名称|uof:元数据/uof:经理名称|uof:元数据/uof:页数|uof:元数据/uof:字数|uof:元数据/uof:行数|uof:元数据/uof:段落数">-->
				<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
			</xsl:if>
			<xsl:if test="uof:元数据/元:元数据_5200/元:标题_5201|uof:元数据/元:元数据_5200/元:主题_5202|uof:元数据/元:元数据_5200/元:创建者_5203|uof:元数据/元:元数据_5200/元:最后作者_5205|uof:元数据/元:元数据_5200/元:作者_5204|uof:元数据/元:元数据_5200/元:摘要_5206|uof:元数据/元:元数据_5200/元:编辑次数_5208|uof:元数据/元:元数据_5200/元:分类_520B|uof:元数据/元:元数据_5200/元:关键字集_520D|uof:元数据/元:元数据_5200/元:创建日期_5207">
				<!--<xsl:if test="uof:元数据/uof:标题|uof:元数据/uof:主题|uof:元数据/uof:创建者|uof:元数据/uof:最后作者|uof:元数据/uof:摘要|uof:元数据/uof:创建日期|uof:元数据/uof:编辑次数|uof:元数据/uof:分类|uof:元数据/uof:关键字集">-->
				<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
			</xsl:if>
			<xsl:if test="uof:元数据/元:元数据_5200/元:用户自定义元数据集_520F">
				<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
			</xsl:if>		
		</Relationships>
	</xsl:template>
</xsl:stylesheet>

