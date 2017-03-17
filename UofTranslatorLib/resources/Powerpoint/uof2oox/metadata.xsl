<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:dc="http://purl.org/dc/elements/1.1/" 
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph" 
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"  
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
		xmlns:元="http://schemas.uof.org/cn/2009/metadata">
	<xsl:template name="metadataApp">
		<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
			<!--修改标签 不知道路径是否对  李娟 11.12.26-->
			<xsl:if test="./元:元数据_5200/元:编辑时间_5209">
				<TotalTime>
					<xsl:call-template name="computerTime">
						<!--<xsl:with-param name="uofTime" select="./uof:编辑时间"/>-->
						<xsl:with-param name="uofTime" select="./元:元数据_5200/元:编辑时间_5209"/>
					</xsl:call-template>
				</TotalTime>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:创建应用程序_520A">
				<!--<xsl:if test="uof:创建应用程序">-->
				<Application>
					<!--<xsl:value-of select="./uof:创建应用程序"/>-->
					<xsl:value-of select="./元:元数据_5200/元:创建应用程序_520A"/>
				</Application>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:文档模板_520C">
				<!--<xsl:if test="uof:文档模板">-->
				<Template>
					<!--<xsl:value-of select="./uof:文档模板"/>-->
					<xsl:value-of select="./元:元数据_5200/元:文档模板_520C"/>
				</Template>
			</xsl:if>
			<!--<xsl:if test="uof:公司名称">-->
			<xsl:if test="./元:元数据_5200/元:公司名称_5213">
				<Company>
					<!--<xsl:value-of select="./uof:公司名称"/>-->
					<xsl:value-of select="./元:元数据_5200/元:公司名称_5213"/>
				</Company>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:经理名称_5214">
				<Manager>
					<xsl:value-of select="./元:元数据_5200/元:经理名称_5214"/>
				</Manager>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:页数_5215">
				<Pages>
					<xsl:value-of select="./元:元数据_5200/元:页数_5215"/>
				</Pages>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:字数_5216">
				<Words>
					<xsl:value-of select="./元:元数据_5200/元:字数_5216"/>
				</Words>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:行数_5219">
				<Lines>
					<xsl:value-of select="./元:元数据_5200/元:行数_5219"/>
				</Lines>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:段落数_521A">
				<Paragraphs>
					<xsl:value-of select="./元:元数据_5200/元:段落数_521A"/>
				</Paragraphs>
			</xsl:if>
		</Properties>
   
    
    
	</xsl:template>
	<xsl:template name="computerTime">
		<xsl:param name="uofTime"/>
		<xsl:variable name="dateTime" select="$uofTime"/>
		<xsl:variable name="Time">
			<xsl:if test="contains($dateTime,'T')">
				<xsl:value-of select="substring-after($dateTime,'T')"/>
			</xsl:if>
			<xsl:if test="not(contains($dateTime,'T'))">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Date">
			<xsl:if test="contains($dateTime,'T')">
				<xsl:value-of select="substring-before($dateTime,'T')"/>
			</xsl:if>
			<xsl:if test="not(contains($dateTime,'T'))">
				<xsl:value-of select="$dateTime"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Ytime">
			<xsl:if test="contains($Date,'Y')">
				<xsl:value-of select="substring-before(substring-after($Date,'P'),'Y')"/>
			</xsl:if>
			<xsl:if test="not(contains($Date,'Y'))">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Mtime">
			<xsl:if test="contains($Date,'M')">
				<xsl:value-of select="substring-before(substring-after($Date,'Y'),'M')"/>
			</xsl:if>
			<xsl:if test="not(contains($Date,'M'))">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Dtime">
			<xsl:if test="contains($Date,'D')">
				<xsl:value-of select="substring-before(substring-after($Date,'M'),'D')"/>
			</xsl:if>
			<xsl:if test="not(contains($Date,'D'))">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Htime">
			<xsl:if test="$Time!=0">
				<xsl:if test="contains($Time,'H')">
					<xsl:value-of select="substring-before($Time,'H')"/>
				</xsl:if>
				<xsl:if test="not(contains($Time,'H'))">
					<xsl:value-of select="'0'"/>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$Time=0">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="Mitime">
			<xsl:if test="$Time!=0">
				<xsl:if test="contains($Time,'M')">
					<xsl:if test="contains($Time,'H')">
						<xsl:value-of select="substring-before(substring-after($Time,'H'),'M')"/>
					</xsl:if>
					<xsl:if test="not(contains($Time,'H'))">
						<xsl:value-of select="substring-before($Time,'M')"/>
					</xsl:if>
				</xsl:if>
				<xsl:if test="not(contains($Time,'M'))">
					<xsl:value-of select="'0'"/>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$Time=0">
				<xsl:value-of select="'0'"/>
			</xsl:if>
		</xsl:variable>
		<!--忽略秒-->
		<xsl:value-of select="$Ytime * 12*30*60*24 +$Mtime * 30*60*24 + $Dtime * 60*24 + $Htime *60 +$Mitime"/>
	</xsl:template>
	<xsl:template name="metadataCore">
		<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcmitype="http://purl.org/dc/dcmitype/">
			<xsl:if test="./元:元数据_5200/元:标题_5201">
				<dc:title>
					<xsl:value-of select="./元:元数据_5200/元:标题_5201"/>
				</dc:title>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:主题_5202">
				<dc:subject>
					<xsl:value-of select="./元:元数据_5200/元:主题_5202"/>
				</dc:subject>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:创建者_5203|./元:元数据_5200/元:作者_5204">
				<dc:creator>
					<xsl:value-of select="./元:元数据_5200/元:创建者_5203"/>
                    
				</dc:creator>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:最后作者_5205">
				<cp:lastModifiedBy>
					<xsl:value-of select="./元:元数据_5200/元:最后作者_5205"/>
				</cp:lastModifiedBy>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:摘要_5206">
				<dc:description>
					<xsl:value-of select="./元:元数据_5200/元:摘要_5206"/>
				</dc:description>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:创建日期_5207">
				<dcterms:created xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">
					<xsl:value-of select="concat(./元:元数据_5200/元:创建日期_5207,'Z')"/>
				</dcterms:created>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:编辑次数_5208">
				<cp:revision>
					<xsl:value-of select="./元:元数据_5200/元:编辑次数_5208"/>
				</cp:revision>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:分类_520B">
				<cp:category>
					<xsl:value-of select="./元:元数据_5200/元:分类_520B"/>
				</cp:category>
			</xsl:if>
			<xsl:if test="./元:元数据_5200/元:关键字集_520D">
				<xsl:for-each select="./元:元数据_5200/元:关键字集_520D/元:关键字_520E">
					<cp:keywords>
						<xsl:value-of select="."/>
					</cp:keywords>
				</xsl:for-each>
			</xsl:if>
		</cp:coreProperties>
	</xsl:template>
	<xsl:template name="metadataCustom">
		<xsl:if test="./元:元数据_5200/元:用户自定义元数据集_520F">
			<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
				<xsl:for-each select="./元:元数据_5200/元:用户自定义元数据集_520F/元:用户自定义元数据_5210">
          
          <!--start liuyin 20121229 修改元数据经过转换后，ooxml文档无法打开-->
          <xsl:if test="./@名称_5211='工作组' or ./@名称_5211='打字员' or ./@名称_5211='记录日期' or ./@名称_5211='编辑者' or ./@名称_5211='语言' or ./@名称_5211='状态' or ./@名称_5211='电话号码'">
          <!--end liuyin 20121229 修改元数据经过转换后，ooxml文档无法打开-->
            
					<xsl:variable name="num">
						<xsl:number count="元:用户自定义元数据_5210" format="1" level="single"/>
					</xsl:variable>
					<property>
						<xsl:attribute name="fmtid"><xsl:value-of select="'{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'"/></xsl:attribute>
						<xsl:attribute name="pid"><xsl:value-of select="$num +1"/></xsl:attribute>
						<xsl:attribute name="name"><xsl:value-of select="./@名称_5211"/></xsl:attribute>
						<xsl:choose>
							<xsl:when test="./@类型_5212='string'">
								<vt:lpwstr>
									<xsl:value-of select="."/>
								</vt:lpwstr>
							</xsl:when>
							<xsl:when test="./@类型_5212='float'">
								<vt:i4>
									<xsl:value-of select="."/>
								</vt:i4>
							</xsl:when>
							<xsl:when test="./@类型_5212='datetime'">
								<vt:filetime>
									<xsl:value-of select="."/>
								</vt:filetime>
							</xsl:when>
							<xsl:when test="./@类型_5212='boolean'">
								<vt:bool>
                  <xsl:choose>
                    <xsl:when test ="'是'">true</xsl:when>
                    <xsl:when test ="'否'">false</xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="."/>
                    </xsl:otherwise>
                  </xsl:choose>
								</vt:bool>
							</xsl:when>
						</xsl:choose>
					</property>
          <!--start liuyin 20121229 修改元数据经过转换后，ooxml文档无法打开-->
          </xsl:if>
          <!--end liuyin 20121229 修改元数据经过转换后，ooxml文档无法打开-->
				</xsl:for-each>
			</Properties>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
