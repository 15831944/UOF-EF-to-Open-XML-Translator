<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
				xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:规则="http://schemas.uof.org/cn/2009/rules"
				xmlns:式样="http://schemas.uof.org/cn/2009/styles"
				xmlns:对象="http://schemas.uof.org/cn/2009/objects">
    <xsl:output method="xml" indent="yes"/>
    <!--添加预处理式样单 李娟11.12.26-->
	<xsl:template name="uof:UOF" match="/">
		<uof:UOF>
			<xsl:copy-of select="document('uof.xml')"/>
			<uof:元数据 id="meta.xml">
				<xsl:copy-of select="document('_meta/meta.xml')"/>
			</uof:元数据>
			<uof:式样集  id="styles.xml">
				<xsl:copy-of select="document('styles.xml')"/>
			</uof:式样集>
			<uof:演示文稿 id="content.xml">
				<演:公用处理规则>
					
					<xsl:copy-of select="document('rules.xml')"/>
					<xsl:copy-of select="document('theme.xml')"/>
				</演:公用处理规则>
				<演:主体>
					<xsl:copy-of select="document('content.xml')"/>
				</演:主体>
					
			</uof:演示文稿>
			<uof:扩展区 id="extend.xml">
				<xsl:copy-of select="document('extend.xml')"/>
			</uof:扩展区>
			

			
			<uof:对象集 id="objectdata.xml">
				<xsl:copy-of select="document('graphics.xml')"/>
				<xsl:if test="document('content.xml')//演:声音_6B17/@自定义声音_C632!='none'or document('content.xml')//图:图片_8005/@图形引用_8007 or document('styles.xml')//字:图片符号_411B or document('graphics.xml')//图:图片_8005/@图形引用_8007 or document('graphics.xml')//图:图片数据引用_8037 or document('content.xml')//演:声音_6B22/@自定义声音_C632!='none' or document('content.xml')//演:声音_6B22/@预定义声音_C631!='none'">
				<xsl:copy-of select="document('objectdata.xml')//对象:对象数据集_D700"/>
				</xsl:if>
				<!--<xsl:if  test="document('content.xml')//图:图片_8005/@图形引用_8007 or document('styles.xml')//字:图片符号_411B or  document('graphics.xml')//图:图片_8005/@图形引用_8007 or document('graphics.xml')//图:图片数据引用_8037 or document('content.xml')//演:声音_6B22[@预定义声音_C631!='none'] or document('content.xml')//演:声音_6B22[@自定义声音_C632!='none'] or document('content.xml')//演:声音_6B17[@自定义声音_C632!='none'] or document('content.xml')//演:声音_6B17[@自定义声音_C632!='none'] ">
					<xsl:copy-of select="document('objectdata.xml')//对象:对象数据集_D700"/>
				</xsl:if>-->
			</uof:对象集>
			
			<xsl:variable name="hyperlinks">
				<xsl:if test="document('graphics.xml')//字:区域开始_4165[@类型_413B='hyperlink']">
					<xsl:value-of select="document('hyperlinks.xml')"/>
				</xsl:if>
			</xsl:variable>
				<xsl:if test="$hyperlinks!=''">
				<uof:链接集>
					<xsl:copy-of select="document('hyperlinks.xml')"/>
				</uof:链接集>
			</xsl:if>
	</uof:UOF>
	
		</xsl:template>
</xsl:stylesheet>
