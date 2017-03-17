<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:规则="http://schemas.uof.org/cn/2009/rules"
                xmlns:元="http://schemas.uof.org/cn/2009/metadata"
                xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:对象="http://schemas.uof.org/cn/2009/objects"
                xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" 
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  
	<xsl:template match="表:工作表集" mode="rels">
		<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
			<xsl:variable name="s">
				<xsl:value-of select="count(/uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825[not(@是否隐藏_E73C='true')])"/>
			</xsl:variable>
			<fileVersion>
				<xsl:attribute name="appName">xl</xsl:attribute>
        <xsl:if test ="/uof:UOF/元:元数据_5200/元:编辑次数_5208">
          <xsl:variable name="time" select="/uof:UOF/元:元数据_5200/元:编辑次数_5208"/>
          <xsl:attribute name="lastEdited">
            <xsl:value-of select="$time"/>
          </xsl:attribute>
          <xsl:attribute name="lowestEdited">
            <xsl:value-of select="$time"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test ="not(/uof:UOF/元:元数据_5200/元:编辑次数_5208)">
          <xsl:attribute name ="lastEdited">5</xsl:attribute>
          <xsl:attribute name ="lowestEdited">4</xsl:attribute>
        </xsl:if>
        <xsl:attribute name ="rupBuild">9302</xsl:attribute>
			</fileVersion>
			<!--workbookPr-->
			<workbookPr>
         <!--
        <xsl:attribute name="date1904">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
       -->
        <xsl:if test="/uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:日期系统_B614='1904'">
          <xsl:attribute name="date1904">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="filterPrivacy"><xsl:value-of select="'1'"/></xsl:attribute>
				<xsl:attribute name="defaultThemeVersion"><xsl:value-of select="'124226'"/></xsl:attribute>
			</workbookPr>
			<!--bookViews-->
			<bookViews>
				<workbookView xWindow="600" yWindow="90" windowWidth="19200" windowHeight="11640">
					<xsl:if test="./表:单工作表/表:工作表_E825[@是否隐藏_E73C]">
						<xsl:attribute name="firstSheet"><xsl:value-of select="'1'"/></xsl:attribute>
						<xsl:attribute name="activeTab"><xsl:value-of select="$s"/></xsl:attribute>
					</xsl:if>
          <!--Modified by LDM in 2011/01/23-->
          <!--是否显示工作表标签-->
          <xsl:if test="/uof:UOF /规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否显示工作表标签_B635='false'">
            <xsl:attribute name="showSheetTabs">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </xsl:if>
          <!--Modified by LDM in 2011/01/23-->
          <!--是否显示水平滚动条-->
          <xsl:if test="/uof:UOF /规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否显示水平滚动条_B636='false'">
            <xsl:attribute name="showHorizontalScroll">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </xsl:if>
          <!--Modified by LDM in 2011/01/23-->
          <!--是否显示垂直滚动条-->
          <xsl:if test="/uof:UOF /规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否显示垂直滚动条_B637='false'">
            <xsl:attribute name="showVerticalScroll">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:for-each select="./表:单工作表/表:工作表_E825">
            <xsl:if test="./表:工作表属性_E80D/表:视图_E7D5/表:是否选中_E7D6='true'">
              <xsl:attribute name="activeTab">
                <xsl:value-of select="position() - 1"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
				</workbookView>
			</bookViews>
			<!--sheets-->
			<sheets>
			
				
				<xsl:for-each select="./表:单工作表">
					<sheet>
						<xsl:variable name="sheetName" select="./表:工作表_E825/@名称_E822"/>
						<xsl:attribute name="name"><xsl:value-of select="$sheetName"/></xsl:attribute>
						<xsl:attribute name="sheetId"><xsl:value-of select="@uof:sheetNo"/></xsl:attribute>
						<xsl:if test="./表:工作表_E825[@是否隐藏_E73C='true']">
							<xsl:attribute name="state"><xsl:value-of select="'hidden'"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="r:id"><xsl:value-of select="concat('rId',@uof:sheetNo)"/></xsl:attribute>
					</sheet>
				</xsl:for-each>
        <!--20130206 gaoyuwei bug 2666 图表Sheet的顺序不正确 start-->
        <!--<xsl:if test="../表:图表工作表集/表:图表工作表[图表:图表_E837]">
          <xsl:for-each select="../表:图表工作表集/表:图表工作表[图表:图表_E837]">
            <sheet>
              <xsl:variable name="sheetName" select="./表:工作表_E825/@名称_E822"/>
              <xsl:attribute name="name">
                <xsl:value-of select="$sheetName"/>
              </xsl:attribute>
              <xsl:attribute name="sheetId">
                <xsl:value-of select="translate(./表:工作表_E825/@标识符_E7AC,'SHEET_','')+10"/>
              </xsl:attribute>
              <xsl:if test="./表:工作表_E825/@是否隐藏_E73C='true'">
                <xsl:attribute name="state">
                  <xsl:value-of select="'hidden'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('rIdChartSheet',position())"/>
              </xsl:attribute>
            </sheet>
          </xsl:for-each> 
        </xsl:if> end -->
      
      <!--20140519 凌峰  图表Sheet的顺序不正确 start-->
				<xsl:if test="../表:图表工作表集/表:图表工作表[图表:图表_E837]">
					<xsl:for-each select="../表:图表工作表集/表:图表工作表[图表:图表_E837]">
						<sheet>
							<xsl:variable name="sheetName" select="./表:工作表_E825/@名称_E822"/>
							<xsl:attribute name="name">
								<xsl:value-of select="$sheetName"/>
							</xsl:attribute>
							<xsl:attribute name="sheetId">
								<xsl:value-of select="translate(./表:工作表_E825/@标识符_E7AC,'SHEET_','')+10"/>
							</xsl:attribute>
							<xsl:if test="./表:工作表_E825/@是否隐藏_E73C='true'">
								<xsl:attribute name="state">
									<xsl:value-of select="'hidden'"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:attribute name="r:id">
								<xsl:value-of select="concat('rIdChartSheet',position())"/>
							</xsl:attribute>
						</sheet>
					</xsl:for-each>
				</xsl:if>
				<!--end-->

			</sheets>
			<calcPr calcId="124519">
        <!--<xsl:attribute name="fullPrecision">0</xsl:attribute>-->

        <xsl:if test="/uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:精确度是否以显示值为准_B613='true'">
          <xsl:attribute name="fullPrecision">0</xsl:attribute>
        </xsl:if>

        <xsl:if test="/uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否RC引用_B634='true'">
          <xsl:attribute name="refMode">
            <xsl:value-of select="'R1C1'"/>
          </xsl:attribute>
        </xsl:if>
			</calcPr>
		</workbook>
	</xsl:template>
</xsl:stylesheet>
