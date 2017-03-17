<?xml version="1.0" encoding="UTF-8"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  xmlns:uof="http://schemas.uof.org/cn/2003/uof">
	<xsl:template name="package-relationships">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <a>test</a>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
			<!--matadata-->
      <xsl:if test ="uof:元数据/uof:编辑时间|uof:元数据/uof:创建应用程序|uof:元数据/uof:文档模板|uof:元数据/uof:公司名称|uof:元数据/uof:经理名称|uof:元数据/uof:页数|uof:元数据/uof:字数|uof:元数据/uof:行数|uof:元数据/uof:段落数">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
      </xsl:if>
      <xsl:if test ="uof:元数据/uof:标题|uof:元数据/uof:主题|uof:元数据/uof:创建者|uof:元数据/uof:最后作者|uof:元数据/uof:摘要|uof:元数据/uof:创建日期|uof:元数据/uof:编辑次数|uof:元数据/uof:分类|uof:元数据/uof:关键字集">
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      </xsl:if>
      </Relationships>
	</xsl:template>
	</xsl:stylesheet>
