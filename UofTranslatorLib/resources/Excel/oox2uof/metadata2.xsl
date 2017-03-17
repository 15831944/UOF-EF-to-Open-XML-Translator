<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:uof="http://schemas.uof.org/cn/2009/uof" xmlns:表="http://schemas.uof.org/cn/2003/spreadsheet" xmlns:演="http://schemas.uof.org/cn/2003/slideshow" xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc" xmlns:图="http://schemas.uof.org/cn/2003/graph" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--
	
	<xsl:template name="metaData">
    
		<uof:元数据>
			
			
				<xsl:apply-templates select="app:Properties"/>
			
			
				<xsl:apply-templates select="cp:coreProperties"/>
			
			
				<xsl:apply-templates select="cus:Properties"/>
			
		</uof:元数据>
	</xsl:template>

-->

  <!--
	
	<xsl:template match="app:Properties">
	    <xsl:if test="app:TotalTime">
		<xsl:variable name ="Mintime">
        <xsl:value-of select="./app:TotalTime"/>
        </xsl:variable>
        <xsl:variable name ="pretime">
        <xsl:value-of select ="concat('P0Y0M0DT0H',$Mintime)"/>
        </xsl:variable>
        <xsl:variable name ="totaltime">
        <xsl:value-of select ="concat($pretime,'M0S')"/>
        </xsl:variable>
            <uof:编辑时间 uof:locID="u0010">
				<xsl:value-of select="$totaltime"/>
			</uof:编辑时间>
		</xsl:if>
		
		<xsl:if test="app:Application">
			<uof:创建应用程序 uof:locID="u0011">
				<xsl:value-of select="./app:Application"/>
			</uof:创建应用程序>
		</xsl:if>
		<xsl:if test="app:Template">
			<uof:文档模板 uof:locID="u0013">
				<xsl:value-of select="./app:Template"/>
			</uof:文档模板>
		</xsl:if>
		<xsl:if test="app:Company">
			<uof:公司名称 uof:locID="u0018">
				<xsl:value-of select="./app:Company"/>
			</uof:公司名称>
		</xsl:if>
		<xsl:if test="app:Manager">
			<uof:经理名称 uof:locID="u0019">
				<xsl:value-of select="./app:Manager"/>
			</uof:经理名称>
		</xsl:if>
		<xsl:if test="app:Pages">
			<uof:页数 uof:locID="u0020">
				<xsl:value-of select="./app:Pages"/>
			</uof:页数>
		</xsl:if>
		<xsl:if test="app:Words">
			<uof:字数 uof:locID="u0021">
				<xsl:value-of select="./app:Words"/>
			</uof:字数>
		</xsl:if>
		<xsl:if test="app:Lines">
			<uof:行数 uof:locID="u0024">
				<xsl:value-of select="./app:Lines"/>
			</uof:行数>
		</xsl:if>
		<xsl:if test="app:Paragraphs">
			<uof:段落数 uof:locID="u0025">
				<xsl:value-of select="./app:Paragraphs"/>
			</uof:段落数>
		</xsl:if>
	</xsl:template>
	
	
	
	
	
	
	
	<xsl:template match="cp:coreProperties">
		<xsl:if test="dc:title">
			<uof:标题 uof:locID="u0002">
				<xsl:value-of select="./dc:title"/>
			</uof:标题>
		</xsl:if>
		<xsl:if test="dc:subject">
			<uof:主题 uof:locID="u0003">
				<xsl:value-of select="./dc:subject"/>
			</uof:主题>
		</xsl:if>
		<xsl:if test="dc:creator">
			<uof:创建者 uof:locID="u0004">
				<xsl:value-of select="./dc:creator"/>
			</uof:创建者>
		</xsl:if>
		<xsl:if test="cp:lastModifiedBy">
			<uof:最后作者 uof:locID="u0006">
				<xsl:value-of select="./cp:lastModifiedBy"/>
			</uof:最后作者>
		</xsl:if>
		<xsl:if test="dc:description">
			<uof:摘要 uof:locID="u0007">
				<xsl:value-of select="./dc:description"/>
			</uof:摘要>
		</xsl:if>
		<xsl:if test="dcterms:created">
			<uof:创建日期 uof:locID="u0008">
				<xsl:value-of select="substring(dcterms:created,1,string-length(dcterms:created)-1)"/>
			</uof:创建日期>
		</xsl:if>
		<xsl:if test="cp:revision">
			<uof:编辑次数 uof:locID="u0009">
				<xsl:value-of select="./cp:revision"/>
			</uof:编辑次数>
		</xsl:if>
		<xsl:if test="cp:category">
			<uof:分类 uof:locID="u0012">
				<xsl:value-of select="./cp:category"/>
			</uof:分类>
		</xsl:if>
		<xsl:if test="cp:keywords">
			<uof:关键字集 uof:locID="u0014">
				<xsl:for-each select="./cp:keywords">
					<uof:关键字 uof:locID="u0015">
						<xsl:value-of select="."/>
					</uof:关键字>
				</xsl:for-each>
			</uof:关键字集>
		</xsl:if>
	</xsl:template>
	<xsl:template match="cus:Properties">
		<uof:用户自定义元数据集 uof:locID="u0016">
			<xsl:for-each select="./cus:property">
				<uof:用户自定义元数据 uof:locID="u0017" uof:attrList="名称 类型">
					<xsl:attribute name="uof:名称"><xsl:value-of select="./@name"/></xsl:attribute>
					<xsl:choose>
						<xsl:when test="name(./node()[position()=1]) = 'vt:lpwstr'">
							<xsl:attribute name="uof:类型"><xsl:value-of select="'string'"/></xsl:attribute>
						</xsl:when>
						<xsl:when test="name(./node()[position()=1]) = 'vt:filetime'">
							<xsl:attribute name="uof:类型"><xsl:value-of select="'datetime'"/></xsl:attribute>
						</xsl:when>
						<xsl:when test="name(./node()[position()=1]) = 'vt:i4'">
							<xsl:attribute name="uof:类型"><xsl:value-of select="'float'"/></xsl:attribute>
						</xsl:when>
						<xsl:when test="name(./node()[position()=1]) = 'vt:bool'">
							<xsl:attribute name="uof:类型"><xsl:value-of select="'boolean'"/></xsl:attribute>
						</xsl:when>
					</xsl:choose>
					<xsl:attribute name="uof:名称"><xsl:value-of select="./@name"/></xsl:attribute>
					<xsl:value-of select="./node()[position()=1]"/>
				</uof:用户自定义元数据>
			</xsl:for-each>
		</uof:用户自定义元数据集>
	</xsl:template>
  -->
</xsl:stylesheet>
