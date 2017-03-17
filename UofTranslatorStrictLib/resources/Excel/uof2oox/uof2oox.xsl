<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                 xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:公式="http://schemas.uof.org/cn/2009/equations"
                xmlns:规则="http://schemas.uof.org/cn/2009/rules"
                xmlns:元="http://schemas.uof.org/cn/2009/metadata"
                xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:对象="http://schemas.uof.org/cn/2009/objects"
                xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
                xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:xsd="http://www.ord.com"
				xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">

	<!-- 20130422 add by gaoyuwei 2875_1 uof-oo功能测试 组合图片图片丢失 上数7行代码-->
  <xsl:import href="metadata.xsl"/>
  <xsl:import href="contentTypes.xsl"/>
  <xsl:import href="package_relationships.xsl"/>
  <xsl:import href="theme.xsl"/>
  <xsl:import href="hyperlink.xsl"/>
  <xsl:import href="sharedStrings.xsl"/>
  <xsl:import href="relationships.xsl"/>
  <xsl:import href="sheet.xsl"/>
  <xsl:import href="style.xsl"/>
  <xsl:import href="workbook.xsl"/>
  <xsl:import href="chart.xsl"/>
  <xsl:import href="drawing.xsl"/>
  <xsl:import href="predefined.xsl"/>
  <xsl:import href="txBody.xsl"/>
  <xsl:import href="pPr.xsl"/>
  <xsl:param name="outputFile"/>
  <xsl:output method="xml" encoding="UTF-8"/>
  <xsl:template match="/uof:UOF">
    <pzip:archive pzip:target="{$outputFile}">
      <xsl:if test="//w:media">
        <xsl:copy-of select="//w:media/*"/>
      </xsl:if>
      <!-- content types -->
      <pzip:entry pzip:target="[Content_Types].xml">
        <xsl:call-template name="ContentTypes"/>
        <!--<xsl:for-each select="//uof:UOF/uof:对象集/uof:其他对象">
          <xsl:copy-of select="child::*"/>
        </xsl:for-each>-->
      </pzip:entry>
      <!-- _rels -->
      <pzip:entry pzip:target="_rels/.rels">
        <xsl:call-template name="pr"/>
      </pzip:entry>
      <!--metadata-->
      <xsl:if test="元:元数据_5200">
        <pzip:entry pzip:target="docProps/app.xml">
          <xsl:apply-templates select="元:元数据_5200" mode="app"/>
        </pzip:entry>
        <pzip:entry pzip:target="docProps/core.xml">
          <xsl:apply-templates select="元:元数据_5200" mode="core"/>
        </pzip:entry>
        
    	  <!--20130105,gaoyuwei，解决2634BUG"元数据丢失"UOF-OOXML start-->
		  <xsl:if test="元:元数据_5200/元:用户自定义元数据集_520F">
			  <pzip:entry pzip:target="docProps/custom.xml">
				  <xsl:apply-templates select="元:元数据_5200" mode="custom"/>
			  </pzip:entry>
		  </xsl:if>
		  <!--end-->
		  
      </xsl:if>

      <!--2014-3-11, add by Qihy, 用户自定义数据集 start-->
      <!--custom xml-->
      <xsl:if test="//扩展:扩展区_B200/w:customItem">
        <xsl:for-each select="//扩展:扩展区_B200/w:customItem">
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="@fileName"/>
            </xsl:attribute>
            <xsl:copy-of select="./*"/>
          </pzip:entry>
        </xsl:for-each>
      </xsl:if>
      <!--2014-3-11 end-->
      
      <!--xl/sharedStrings.xml-->

      <xsl:if test="./uof:单元格内容集合[uof:单元格内容]">
        <pzip:entry>
          <xsl:attribute name="pzip:target">
            <xsl:value-of select="'xl/sharedStrings.xml'"/>
          </xsl:attribute>
          <xsl:apply-templates select="uof:单元格内容集合[uof:单元格内容]"/>
          <!---->
        </pzip:entry>
      </xsl:if>

      <!--xl/worksheets-->
      <xsl:if test=".//表:单工作表">
        <xsl:for-each select=".//表:单工作表">
          <xsl:variable name="sheetNo" select="./@uof:sheetNo"/>
          <!--xl/worksheets/sheetX.xml-->
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/worksheets/sheet',$sheetNo,'.xml')"/>
            </xsl:attribute>

            <xsl:call-template name="Sheet">
              <xsl:with-param name="tmpsheet" select="$sheetNo"/>
            </xsl:call-template>
          </pzip:entry>
          <!--xl/worksheets/_rels/sheetX.xml.rels-->
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/worksheets/_rels/sheet',$sheetNo,'.xml.rels')"/>
            </xsl:attribute>
            <xsl:call-template name="SheetRels">
              <xsl:with-param name="relsSheet" select="$sheetNo"/>
            </xsl:call-template>
          </pzip:entry>

          <!--修改 李杨2012-2-21-->
          <xsl:for-each select ="./表:图表集合/表:单图表">
            <!--<xsl:for-each select="./表:工作表_E825/表:工作表内容_E80E/表:图表">-->
            <xsl:variable name="chartPos">
              <xsl:value-of select="position()"/>
            </xsl:variable>
            
            <!--20130116 gaoyuwei bug_2642_2643_2644 标题集图片填充的填充内容丢失 start-->

			  <xsl:if test=".//图:图片_8005[@图形引用_8007]">
          
           <xsl:variable name ="chartpos">
             <!--<xsl:value-of select ="position()"/>-->
             <xsl:value-of select="./@uof:chartNo"/>
           </xsl:variable>
           <xsl:variable name="chartNo">
             <!--<xsl:value-of select="concat($sheetNo,'_',@uof:chartNo)"/>-->
             <xsl:value-of select="concat($sheetNo,'_',$chartpos)"/>
           </xsl:variable>
           <pzip:entry>
             <xsl:attribute name="pzip:target">
               <!--取值形式为：chart+sheet号_chart号.xml/rels，如chart11.xml.rels，表示第一个工作表的第一个图表引用-->
               <xsl:value-of select="concat('xl/charts/_rels/chart',$chartNo,'.xml.rels')"/>
             </xsl:attribute>
             <!--xsl:attribute name="pzip:target">
                  <xsl:value-of select="concat('xl/charts/_rels/chart',$sheetNo,'_',$chartPos,'.xml.rels')"/>
                </xsl:attribute-->
             
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <xsl:for-each select=".//图:图片_8005[@图形引用_8007]">
                    <xsl:call-template name="PhotosInChartsRels">
                    </xsl:call-template>
                  </xsl:for-each>
                </Relationships>
              </pzip:entry>
            </xsl:if>
          </xsl:for-each>
			
			<!--end-->
          
          <!--xl/drawings/drawingX.xml-->
          <xsl:if test=".//表:工作表内容_E80E//uof:锚点_C644[@图形引用_C62E]">
            <pzip:entry>
              <!-- 20130519 update by xuzhenwei 为了匹配C#中的处理的方法以及OO文件命名的规范，修改oo中的drawings下的名称 -->
              <xsl:attribute name="pzip:target">
                <xsl:value-of select="concat('xl/drawings/drawing',$sheetNo,'.xml')"/>
              </xsl:attribute>
              <!-- end -->
              <xdr:wsDr>
                
                <!--修改添加：图片、图表、预定义图形引用。李杨2012-2-16-->
                <xsl:for-each select ="./表:工作表_E825/表:工作表内容_E80E//uof:锚点_C644[@图形引用_C62E]">
                  <xsl:variable name ="grapref">
                    <xsl:value-of select ="@图形引用_C62E"/>
                  </xsl:variable>
                  <xsl:choose >
                    <!--图片对应的drawing-->
                    <xsl:when test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref]/图:图片数据引用_8037">
                      <xsl:call-template name="PictureDrawing">
                        <xsl:with-param name ="id">
                          <xsl:value-of select ="$grapref"/>
                        </xsl:with-param>
                      </xsl:call-template>
                    </xsl:when>
                    <!--图表对应的drawing-->
                    <xsl:when test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref]/图:图表数据引用_8065">
                      <xsl:call-template name="drawing"/>
                    </xsl:when>
                    <!--公式对应的drawing-->
                    <xsl:when test="./ancestor::uof:UOF/公式:公式集_C200/公式:数学公式_C201[@标识符_C202=$grapref]">
                      <xsl:call-template name="equDrawing">
                        <xsl:with-param name="ref" select="$grapref"></xsl:with-param>
                      </xsl:call-template>
                    </xsl:when>
                    <!--预定义图形对应的drawing,包括批注框-->
                    <xsl:otherwise >
                      <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$grapref]">
                        <xsl:call-template name="PredefinedPicDrawing">
                          <xsl:with-param name="seq" select="number($sheetNo * 200) + position()"/>
                        </xsl:call-template>
                      </xsl:if>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:for-each>
             
              </xdr:wsDr>
            </pzip:entry>
          </xsl:if>
          <!--xl/charts/chartX.xml-->
          <!--图表主体-->
          <xsl:if test="./表:图表集合">
            <xsl:for-each select="./表:图表集合/表:单图表">
              <xsl:variable name ="chartpos">
                <!--<xsl:value-of select ="position()"/>-->
                <xsl:value-of select="./@uof:chartNo"/>
              </xsl:variable>
              <xsl:variable name="chartNo">
                <!--<xsl:value-of select="concat($sheetNo,'_',@uof:chartNo)"/>-->
                <xsl:value-of select="concat($sheetNo,'_',$chartpos)"/>
              </xsl:variable>
              <pzip:entry>
                <xsl:attribute name="pzip:target">
                  <!--取值形式为：chart+sheet号_chart号.xml，如chart11.xml，表示第一个工作表的第一个图表-->
                  <xsl:value-of select="concat('xl/charts/chart',$chartNo,'.xml')"/>
                </xsl:attribute>
                <xsl:call-template name="chart"/>
              </pzip:entry>
            </xsl:for-each>
          </xsl:if>
          <!--李杨修改，2012-2-20-->
          <xsl:if test="./表:工作表_E825/表:工作表内容_E80E//uof:锚点_C644">

            <pzip:entry>
              <!-- 20130519 update by xuzhenwei 为了匹配C#中的处理的方法以及OO文件命名的规范，修改oo中的drawings下的名称 -->
              <xsl:attribute name="pzip:target">
                <xsl:value-of select="concat('xl/drawings/_rels/drawing',$sheetNo,'.xml.rels')"/>
              </xsl:attribute>
              <!-- end -->
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				  
                <xsl:for-each select="./表:工作表_E825/表:工作表内容_E80E//uof:锚点_C644">
                  <!--图片对应的xl/drawings/_rels/drawingX.xml.rels-->
                  <xsl:variable name ="graphref">
                    <xsl:value-of select ="./@图形引用_C62E"/>
                  </xsl:variable>

					<!-- 20130423 update by gaoyuwei 2875_1 uof-oo功能测试 组合图片图片丢失 start-->
					<xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphref]">
							<!--end-->
							<xsl:if test="./ancestor::uof:UOF//表:单工作表[@uof:sheetNo=$sheetNo]//表:工作表内容_E80E">
								<xsl:call-template name="AnchorRelationship">
									<xsl:with-param name="graphrefAnchor">
										<xsl:value-of select="./@图形引用_C62E"/>
									</xsl:with-param>
								</xsl:call-template>
								<!-- 20130425 update by gaoyuwei 2875_2 uof-oo功能测试 组合图片图片丢失 start-->
								<xsl:if test="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphref]/@组合列表_8064">
									<!--获取组合图片的节点位置--> 
									<xsl:variable name="groupPosition" select="position()"/>
									  <xsl:variable name="groupLst" select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphref]/@组合列表_8064"/>
									<!--组合图片此处有待修改（3个以上组合）-->
									<xsl:variable name="graphrefzuheBefore">
										<xsl:if test ="contains($groupLst, ./@标识符_804B)">
											<xsl:value-of select="substring-before($groupLst,' ')"/>
										</xsl:if>
									</xsl:variable>
									<xsl:variable name="graphrefzuheAfter">
										<xsl:if test ="contains($groupLst, ./@标识符_804B)">
											<xsl:value-of select="substring-after($groupLst,' ')"/>
										</xsl:if>
									</xsl:variable>
									
									<xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphrefzuheBefore]">
										<xsl:call-template name="ZuheAnchorRelationship">
											<xsl:with-param name="groupPosition" select="$groupPosition +100"/>
										</xsl:call-template>
									</xsl:for-each>
									<xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphrefzuheAfter]">
										<xsl:call-template name="ZuheAnchorRelationship">
											<xsl:with-param name="groupPosition" select="$groupPosition + 100 +1"/>
										</xsl:call-template>
									</xsl:for-each>
								</xsl:if>
								<!--end-->
								
							</xsl:if>
                  </xsl:if>
						
				<!--20130513 update by gaoyuwei bug_2918 第三轮回归 uof-oo 功能测试 预定义图形图片填充丢失 start-->	
                  <!--<xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphref]//图:填充_804C/图:图片_8005">
                    <xsl:call-template name="AnchorRelationship"/>
                  </xsl:if>-->
				<!--end-->                  
                  
                  <!--图表对应的xl/drawings/_rels/drawingX.xml.rels-->
                  <xsl:if test ="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphref]/图:图表数据引用_8065">
                    <xsl:if test="./ancestor::uof:UOF//表:单工作表[@uof:sheetNo=$sheetNo]//表:工作表内容_E80E">
                      <!--<xsl:for-each select="./ancestor::uof:UOF//表:单工作表[@uof:sheetNo=$sheetNo]//表:工作表内容_E80E//uof:锚点_C644[@图形引用_C62E=./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[图:图表数据引用_8065]/@标识符_804B]">-->
                      <xsl:call-template name="ChartRelationship"/>
                      <!--</xsl:for-each>-->
                    </xsl:if>
                  </xsl:if>
                </xsl:for-each>

              </Relationships>
            </pzip:entry>
          </xsl:if>
          <xsl:if test=".//表:批注_E7B7">
            <pzip:entry>
              <xsl:attribute name="pzip:target">
                <xsl:value-of select="concat('xl/comments',$sheetNo,'.xml')"/>
              </xsl:attribute>
              <xsl:call-template name="Comments">
                <xsl:with-param name="relsSheet" select="$sheetNo"/>
              </xsl:call-template>
            </pzip:entry>
          </xsl:if>


        </xsl:for-each>
      </xsl:if>

      <!--chartsheet-->
      <!--修改 李杨2012-5-21-->
      <xsl:if test=".//表:图表工作表集">

        <!--2014-4-15, update by Qihy, chartsheet和chart同时存在时，uof-ooxml转换后需要恢复才能打开， start-->
        <xsl:variable name="temp_count">
          <xsl:choose>
            <xsl:when test="./表:电子表格文档_E826/表:工作表集/表:单工作表/表:图表集合/表:单图表">
              <xsl:value-of select ="count(./表:电子表格文档_E826/表:工作表集/表:单工作表/表:图表集合[表:单图表])"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="sheetNo" select="position()"/>
        <xsl:variable name="drawingNo" select="number($temp_count) + number($sheetNo)"/>
         
        <xsl:for-each select=".//表:图表工作表[图表:图表_E837]">
          <pzip:entry>
            <!-- 20130519 update by xuzhenwei 为了匹配C#中的处理的方法以及OO文件命名的规范，修改oo中的drawings下的名称 -->
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/drawings/drawing',$drawingNo,'.xml')"/>
            </xsl:attribute>
            <!-- end -->
            <xdr:wsDr>
              <!--图表对应的drawing-->
              <xsl:for-each select="./图表:图表_E837">
                <xsl:call-template name="drawingchartsheet">
                  <xsl:with-param name="pos" select="$sheetNo"/>
                </xsl:call-template>
              </xsl:for-each>
            </xdr:wsDr>
          </pzip:entry>
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/chartsheets/chartsheet',position(),'.xml')"/>
            </xsl:attribute>
            <xsl:call-template name="ChartSheet">
              <xsl:with-param name ="seq">
                <xsl:value-of select ="position()"/>
              </xsl:with-param>
            </xsl:call-template>
          </pzip:entry>
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/chartsheets/_rels/chartsheet',position(),'.xml.rels')"/>
            </xsl:attribute>
            <!-- 20130519 update by xuzhenwei 为了匹配C#中的处理的方法以及OO文件命名的规范，修改oo中的drawings下的名称 -->
            <!--<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rIdChartSheet1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/drawing" Target="../drawings/drawing1.xml"/>
            </Relationships>-->
            <!-- end -->
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <!-- 20140429, update by Qihy 修改releationships生成关系有错误, start -->
              <Relationship>
              <xsl:attribute name="Id">
                <xsl:value-of select ="concat('rIdChartSheet', position())"/>
              </xsl:attribute>
              <xsl:attribute name="Type">
                <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
              </xsl:attribute>
              <xsl:attribute name="Target">
                <xsl:value-of select="concat('../drawings/drawing', $drawingNo, '.xml')"/>
              </xsl:attribute>
                </Relationship>
              <!--20140429, end -->
            </Relationships>
            <!--2014-4-15 end-->
            
          </pzip:entry>
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="concat('xl/drawings/_rels/drawing',$drawingNo,'.xml.rels')"/>
            </xsl:attribute>
            <!--<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/chart" Target="../charts/chartsheet1.xml"/>
            </Relationships>-->
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <!-- 20140429, update by Qihy, 修改releationships生成关系有错误, start -->
              <Relationship>
                <xsl:attribute name="Id">
                  <xsl:value-of select ="concat('rIdChartSheet', position())"/>
                </xsl:attribute>
                <xsl:attribute name="Type">
                  <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chart'"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:value-of select="concat('../charts/chartsheet', position(), '.xml')"/>
                </xsl:attribute>
              </Relationship>
              <!-- 20140429, end -->
            </Relationships>
          </pzip:entry>
          <!--xl/charts/chartX.xml-->
          <!--图表主体-->
          <xsl:if test="./图表:图表_E837">
            <xsl:for-each select="./图表:图表_E837">
              <xsl:variable name="chartNo">
                <xsl:value-of select="position()"/>
              </xsl:variable>
              <pzip:entry>
                <xsl:attribute name="pzip:target">
                  <!--取值形式为：chart+sheet号_chart号.xml，如chart11.xml，表示第一个工作表的第一个图表-->
                  <xsl:value-of select="concat('xl/charts/chartsheet',$chartNo,'.xml')"/>
                </xsl:attribute>
                <xsl:call-template name ="chartsheet"/>
                <!--<xsl:call-template name="ChartTrans"/>-->
              </pzip:entry>
            </xsl:for-each>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>

      <!--xl/drawings/_rels/drawingX.xml.rels和xl/drawings/drawingX.xml-->

      <!--xl/drawings/_rels/drawingX.xml.rels-->

      <!--xl/style.xml-->
      <xsl:if test="式样:式样集_990B">
        <pzip:entry pzip:target="xl/styles.xml">
          <xsl:call-template name="Style"/>
        </pzip:entry>
      </xsl:if>

      <!--xl/workbook-->
      <pzip:entry pzip:target="xl/workbook.xml">
        <xsl:apply-templates select="表:电子表格文档_E826/表:工作表集" mode="rels"/>
      </pzip:entry>

      <!--xl/rels/workbook-->
      <pzip:entry pzip:target="xl/_rels/workbook.xml.rels">
        <xsl:call-template name="work_book_rels"/>
      </pzip:entry>

      <!--theme/theme1.xml-->
      <pzip:entry pzip:target="xl/theme/theme1.xml">
        <xsl:call-template name="theme_copy"/>
      </pzip:entry>
    </pzip:archive>
  </xsl:template>

  <xsl:template match="元:元数据_5200" mode="app">
    <xsl:call-template name="metadataApp"/>
  </xsl:template>
  <xsl:template match="元:元数据_5200" mode="core">
    <xsl:call-template name="metadataCore"/>
  </xsl:template>
  <xsl:template match="元:元数据_5200" mode="custom">
    <xsl:call-template name="metadataCustom"/>
  </xsl:template>

  <!--package relationship-->
  <xsl:template name="package-relationships">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties" Target="docProps/app.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
      <xsl:if test="//元:用户自定义元数据集_520F">
        <Relationship Id="rId4" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/custom-properties" Target="docProps/custom.xml"/>
      </xsl:if>
    </Relationships>
  </xsl:template>

  <!--修改 李杨2012-5-9 content_types-->
  <xsl:template name="ContentTypes">
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default Extension="jpeg" ContentType="image/jpeg"/>
      <Default Extension="jpg" ContentType="image/gif"/>
      
      <Default Extension="wmf" ContentType="image/x-wmf"/>
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
      <Default Extension="gif" ContentType="image/gif"/>
      <!--<Default Extension="jpg" ContentType="image/jpeg"/>-->

      <!--2014-4-15, update by Qihy, chartsheet和chart同时存在时，uof-ooxml转换后需要恢复才能打开， start-->
      <xsl:variable name="temp_count">
        <xsl:choose>
          <xsl:when test="./表:电子表格文档_E826/表:工作表集/表:单工作表/表:图表集合/表:单图表">
            <xsl:value-of select ="count(./表:电子表格文档_E826/表:工作表集/表:单工作表/表:图表集合[表:单图表])"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:for-each select ="./表:电子表格文档_E826/表:图表工作表集/表:图表工作表">
        <xsl:variable name="sheetNo" select="position()"/>
        <xsl:variable name="drawingNo" select="number($temp_count) + number($sheetNo)"/>
        <Override>
          <xsl:attribute name="PartName">
            <xsl:value-of select="concat('/xl/chartsheets/chartsheet',$sheetNo,'.xml')"/>
          </xsl:attribute>
          <xsl:attribute name="ContentType">
            <xsl:value-of select="'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'"/>
          </xsl:attribute>
        </Override>
        <xsl:if test ="./表:工作表_E825/uof:锚点_C644">
          <!--drawing类型声明-->
          <Override>
            <!-- 20130519 update by xuzhenwei 为了匹配C#中的处理的方法以及OO文件命名的规范，修改oo中的drawings下的名称 -->
            <xsl:attribute name="PartName">
              <xsl:value-of select="concat('/xl/drawings/drawing',$drawingNo,'.xml')"/>
            </xsl:attribute>
            <!-- end -->
            <!--2014-4-15 end-->

            <xsl:attribute name="ContentType">
              <xsl:value-of select="'application/vnd.openxmlformats-officedocument.drawing+xml'"/>
            </xsl:attribute>
          </Override>
        </xsl:if>
        <xsl:if test ="./图表:图表_E837">
          <xsl:for-each select ="./图表:图表_E837">
            <xsl:variable name="chartNo">
              <xsl:value-of select="position()"/>
            </xsl:variable>
            <Override>
              <xsl:attribute name="PartName">
                <xsl:value-of select="concat('/xl/charts/chartsheet',$chartNo,'.xml')"/>
              </xsl:attribute>
              <xsl:attribute name="ContentType">
                <xsl:value-of select="'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'"/>
              </xsl:attribute>
            </Override>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
      <!--工作表对应的sheet类型声明-->
      <xsl:for-each select="./表:电子表格文档_E826/表:工作表集/表:单工作表">
        <xsl:variable name="sheetNo" select="position()"/>
        <Override>
          <xsl:attribute name="PartName">
            <xsl:value-of select="concat('/xl/worksheets/sheet',$sheetNo,'.xml')"/>
          </xsl:attribute>
          <xsl:attribute name="ContentType">
            <xsl:value-of select="'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'"/>
          </xsl:attribute>
        </Override>
        <!--待修改  李杨2012-1-5-->
        <xsl:if test="./表:图表集合 or ./uof:锚点集合">
          <!--drawing类型声明-->
          <Override>
            <xsl:attribute name="PartName">
              <xsl:value-of select="concat('/xl/drawings/drawing',$sheetNo,'.xml')"/>
            </xsl:attribute>
            <xsl:attribute name="ContentType">
              <xsl:value-of select="'application/vnd.openxmlformats-officedocument.drawing+xml'"/>
            </xsl:attribute>
          </Override>

          <xsl:if test="./表:图表集合">
            <xsl:for-each select="./表:图表集合">
              <!--图表对应的chart类型声明-->
              <xsl:for-each select="./表:单图表">
                <xsl:variable name="chartNo">
                  <xsl:value-of select="concat('chart',$sheetNo,'_',@uof:chartNo)"/>
                </xsl:variable>
                <Override>
                  <xsl:attribute name="PartName">
                    <xsl:value-of select="concat('/xl/charts/',$chartNo,'.xml')"/>
                  </xsl:attribute>
                  <xsl:attribute name="ContentType">
                    <xsl:value-of select="'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'"/>
                  </xsl:attribute>
                </Override>
              </xsl:for-each>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test="./uof:锚点集合">
            <!--锚点对应的类型声明-->
          </xsl:if>
        </xsl:if>
      </xsl:for-each>
      <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      
      <xsl:if test="//元:用户自定义元数据_5210">
        <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
      </xsl:if>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      <!--<xsl:if test="//uof:其他对象[@uof:公共类型='png']">
        <Default Extension="png" ContentType="image/png"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='bmp']">
        <Default Extension="png" ContentType="image/png"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='gif']">
        <Default Extension="gif" ContentType="image/gif"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='wmf']">
        <Default Extension="wmf" ContentType="image/wmf"/>
      </xsl:if>
      <xsl:if test=".//uof:其他对象[@uof:私有类型='wmf']">
        <Default Extension="wmf" ContentType="image/x-wmf"/>
      </xsl:if>
      <xsl:if test=".//uof:其他对象[@uof:私有类型='emf']">
        <Default Extension="emf" ContentType="image/x-emf"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:私有类型='jpeg' or @uof:公共类型='jpeg' or @uof:公共类型='jpg']">
        <Default Extension="jpeg" ContentType="image/jpeg"/>
      </xsl:if>-->
    </Types>
  </xsl:template>

  <xsl:template name="work_book_rels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <xsl:choose>
        <xsl:when test="表:电子表格文档_E826/表:工作表集/表:单工作表[表:图表]">
          <xsl:for-each select="表:电子表格文档_E826/表:主体/表:全工作表">
            <xsl:variable name="relsheet">
              <xsl:value-of select="./表:工作表/@表:名称"/>
            </xsl:variable>
            <xsl:variable name="sername">
              <xsl:value-of select="./表:工作表/表:图表/表:数据源/表:系列/@表:系列名"/>
            </xsl:variable>
            <xsl:variable name="servalue">
              <xsl:value-of select="./表:工作表/表:图表/表:数据源/表:系列/@表:系列值"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="($sername!='' and not(contains($sername,$relsheet))) or ($servalue!='' and not(contains($servalue,$relsheet)))">
                <xsl:variable name="id" select="@uof:id"/>
                <xsl:variable name="sheet_id">
                  <xsl:value-of select="./following::表:全图表[@uof:sheet_id=$id][1]/@uof:chartid"/>
                </xsl:variable>
                <Relationship>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="concat('rId',$id)"/>
                  </xsl:attribute>
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('chartsheets/sheet',$sheet_id,'.xml')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:when>
              <xsl:otherwise>
                <xsl:variable name="id" select="@uof:id"/>
                <Relationship>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="concat('rId',$id)"/>
                  </xsl:attribute>
                  <xsl:attribute name="Type">
                    <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet'"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('worksheets/sheet',$id,'.xml')"/>
                  </xsl:attribute>
                </Relationship>
              </xsl:otherwise>
            </xsl:choose>
            <!--yx,override,considering more cases above,2010.3.31-->
          </xsl:for-each>
        </xsl:when>   
        <!--有关图表 待改 李杨2012-1-5-->
        <xsl:when test="表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825[not(表:图表)]">
          <xsl:for-each select="表:电子表格文档_E826/表:工作表集/表:单工作表">
            <xsl:variable name="id" select="@uof:sheetNo"/>
            <Relationship>
              <xsl:attribute name="Id">
                <xsl:value-of select="concat('rId',$id)"/>
              </xsl:attribute>
              <xsl:attribute name="Type">
                <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet'"/>
              </xsl:attribute>
              <xsl:attribute name="Target">
                <xsl:value-of select="concat('worksheets/sheet',$id,'.xml')"/>
              </xsl:attribute>
            </Relationship>
          </xsl:for-each>
        </xsl:when>
      </xsl:choose>
      <xsl:variable name="lastId" select="表:电子表格文档_E826/表:工作表集/表:单工作表[position()=last()]/@uof:sheetNo"/>
      <!--图表 待改  李杨2012-1-5-->
      <xsl:if test="/uof:UOF/表:电子表格文档_E826/表:图表工作表集/表:图表工作表[图表:图表_E837]">
        <xsl:for-each select="/uof:UOF/表:电子表格文档_E826/表:图表工作表集/表:图表工作表/图表:图表_E837">
          <Relationship>
            <xsl:attribute name="Id">
              <xsl:value-of select="concat('rIdChartSheet',position())"/>
            </xsl:attribute>
            <xsl:attribute name="Type">
              <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet'"/>
            </xsl:attribute>
            <xsl:attribute name="Target">
              <xsl:value-of select="concat('chartsheets/chartsheet',position(),'.xml')"/>
            </xsl:attribute>
          </Relationship>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test=".//表:数据_E7B3[@类型_E7B6='text' or @类型_E7B6='']">

        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select="concat('rId',$lastId+1)"/>
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings'"/>
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select="'sharedStrings.xml'"/>
          </xsl:attribute>
        </Relationship>
      </xsl:if>
      <Relationship>
        <xsl:attribute name="Id">
          <xsl:value-of select="concat('rId',$lastId+2)"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/styles'"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="'styles.xml'"/>
        </xsl:attribute>
      </Relationship>
      <Relationship>
        <xsl:attribute name="Id">
          <xsl:value-of select="concat('rId',$lastId+3)"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/theme'"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="'theme/theme1.xml'"/>
        </xsl:attribute>
      </Relationship>
    </Relationships>
  </xsl:template>
  <xsl:template name="drawingrels1">
    <xsl:for-each select="表:全图表">
      <xsl:variable name="riid">
        <xsl:value-of select="@uof:chartid"/>
      </xsl:variable>
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:attribute name="Id">
          <!--yx,change,add 后缀，2010.4.14-->
          <xsl:value-of select="concat('rId',$riid,'chart')"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/chart'"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../charts/chart',$riid,'.xml')"/>
        </xsl:attribute>
      </Relationship>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="drawpicture2">
    <xsl:for-each select="uof:锚点[图:图形/@图:其他对象]">
      <xsl:for-each select="图:图形[@图:其他对象]">
        <xsl:variable name="uxux" select="@图:标识符"/>
        <xsl:variable name="tttt">
          <xsl:value-of select="@图:其他对象"/>
        </xsl:variable>
        <xsl:variable name="tttt1">
          <xsl:choose>
            <xsl:when test="contains($uxux,'OBJ0000')">
              <xsl:value-of select="substring-after($uxux,'OBJ0000')"/>
            </xsl:when>
            <xsl:when test="contains($uxux,'OBJ000')">
              <xsl:value-of select="substring-after($uxux,'OBJ000')"/>
            </xsl:when>
            <xsl:when test="contains($uxux,'OBJ00')">
              <xsl:value-of select="substring-after($uxux,'OBJ00')"/>
            </xsl:when>
            <xsl:when test="contains($uxux,'OBJ0')">
              <xsl:value-of select="substring-after($uxux,'OBJ0')"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="ttttt">
          <xsl:value-of select="position()"/>
        </xsl:variable>
        <xsl:variable name="uxux1">
          <xsl:choose>
            <xsl:when test="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:公共类型">
              <xsl:value-of select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:公共类型"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="emf" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tttt]/@uof:私有类型"/>
        <xsl:comment>
          <xsl:value-of select="$emf"/>
        </xsl:comment>
        <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <xsl:attribute name="Id">
            <xsl:value-of select="concat('rId',$tttt1)"/>
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
          </xsl:attribute>
          <xsl:variable name="uxuxt2">
            <xsl:choose>
              <xsl:when test="$uxux1='jpg' or $uxux1='jpeg'">
                <xsl:value-of select="'jpeg'"/>
              </xsl:when>
              <xsl:when test="$uxux1='bmp'">
                <xsl:value-of select="'png'"/>
              </xsl:when>
              <xsl:when test="$uxux1='png'">
                <xsl:value-of select="'png'"/>
              </xsl:when>
              <xsl:when test="$emf!=''">
                <xsl:value-of select="$emf"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$uxux1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat('../media/',$tttt,'.',$uxuxt2)"/>
          </xsl:attribute>
        </Relationship>
      </xsl:for-each>
    </xsl:for-each>

    <xsl:for-each select="uof:锚点/图:图形/图:预定义图形[..//图:图片]">
      <xsl:variable  name="pixxt"    select="parent::图:图形/@图:标识符"/>
      <xsl:variable name="ttttp">
        <xsl:choose>
          <xsl:when test="contains($pixxt,'OBJ0000')">
            <xsl:value-of select="substring-after($pixxt,'OBJ0000')"/>
          </xsl:when>
          <xsl:when test="contains($pixxt,'OBJ000')">
            <xsl:value-of select="substring-after($pixxt,'OBJ000')"/>
          </xsl:when>
          <xsl:when test="contains($pixxt,'OBJ00')">
            <xsl:value-of select="substring-after($pixxt,'OBJ00')"/>
          </xsl:when>
          <xsl:when test="contains($pixxt,'OBJ0')">
            <xsl:value-of select="substring-after($pixxt,'OBJ0')"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="tuyinyong">
        <xsl:value-of select="图:属性/图:填充/图:图片/@图:图形引用"/>
      </xsl:variable>
      <xsl:variable name="uxux1t" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tuyinyong]/@uof:公共类型"/>
      <xsl:variable name="uxux2t" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/图:图形[//图:图片]/@图:标识符"/>
      <xsl:variable name="private" select="ancestor::uof:框架的sheet号/preceding-sibling::uof:对象集/uof:其他对象[@uof:标识符=$tuyinyong]/@uof:私有类型"/>
      <xsl:variable name="biaozhitt">
        <xsl:choose>
          <xsl:when test="$uxux1t='jpg'">
            <xsl:value-of select="'jpeg'"/>
          </xsl:when>
          <xsl:when test="$uxux1t='bmp'">
            <xsl:value-of select="png"/>
          </xsl:when>
          <xsl:when test="$private!=''">
            <xsl:value-of select="$private"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$uxux1t"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <xsl:attribute name="Id">
          <xsl:value-of select="concat('rId',$ttttp)"/>
        </xsl:attribute>
        <xsl:attribute name="Type">
          <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <!--xsl:value-of select="concat('../media/',$tuyinyong,'.',$uxux1t,'e')"/-->
          <xsl:value-of select="concat('../media/',$tuyinyong,'.',$biaozhitt)"/>
        </xsl:attribute>
      </Relationship>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="Comments">
    <comments xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
      <authors>
        <author>
          <xsl:value-of select="'作者'"/>
        </author>
      </authors>
        <commentList>
          <xsl:for-each select=".//表:批注_E7B7">
            <comment>
              <xsl:attribute name="ref">
                <xsl:variable name="rowRef">
                  <xsl:value-of select="./ancestor::表:行_E7F1/@行号_E7F3"/>
                </xsl:variable>
                <xsl:variable name="colNum">
                  <xsl:value-of select="./parent::表:单元格_E7F2/@列号_E7BC"/>
                </xsl:variable>
                <xsl:variable name="colRef">
                  <xsl:call-template name="ColIndex">
                    <xsl:with-param name="colSeq" select="$colNum"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:value-of select="concat($colRef,$rowRef)"/>
              </xsl:attribute>
              <xsl:attribute name="authorId">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:variable name="objRef">
                <xsl:value-of select="./uof:锚点_C644/@图形引用_C62E"/>
              </xsl:variable>
              <xsl:for-each select="./ancestor::uof:UOF/图形:图形集_7C00/图:图形_8062[@标识符_804B = $objRef]">
                <text>
                               
                  <!--zl 20150512-->
                  <xsl:for-each select="./图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D">
                    <r>
                      <rPr>
                        <xsl:if test="./字:句属性_4158/字:是否粗体_4130 = 'true'">
                          <b/>
                        </xsl:if>

                        <xsl:if test="./字:句属性_4158/字:是否斜体_4131 = 'true'">
                          <i/>
                        </xsl:if>

                        <xsl:if test="./字:句属性_4158/字:删除线_4135 = 'true'">
                          <strike/>
                        </xsl:if>
                        
                        <sz>
                          <xsl:attribute name="val">
                            <xsl:choose>
                              <xsl:when test="./字:句属性_4158/字:字体_4128[@字号_412D]">
                                <xsl:value-of select="./字:句属性_4158/字:字体_4128/@字号_412D"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="'9'"/>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:attribute>
                        </sz>
                        <color indexed="81"/>
                        <rFont val="Tahoma"/>
                        <family val="2"/>
                        <charset val="134"/>
                      </rPr>
                      <t>
                        <xsl:value-of select=".//字:文本串_415B"/>
                      </t>
                    </r>
                  </xsl:for-each>
                  <!--zl 20150512-->
                  
                </text>
              </xsl:for-each>
            </comment>
          </xsl:for-each>
        </commentList>
    </comments>
  </xsl:template>

  <xsl:template name="drawingchartsheet">
    <xsl:param name="pos"/>
    <xdr:absoluteAnchor>
      <xdr:pos x="0" y="0"/>
      <xdr:ext cx="9301758" cy="6087070"/>
      <xdr:graphicFrame macro="">
        <xdr:nvGraphicFramePr>
          <xdr:cNvPr id="2" name="Chart 1"/>
          <xdr:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
          </xdr:cNvGraphicFramePr>
        </xdr:nvGraphicFramePr>
        <xdr:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
        </xdr:xfrm>
        <a:graphic>
          <a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart">
            
            <!--<c:chart xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" r:id="rId1"/>-->
            <c:chart xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
              <xsl:attribute name="r:id">
                <xsl:value-of select="concat('rIdChartSheet', $pos)"/>
              </xsl:attribute>
            </c:chart>
          </a:graphicData>
        </a:graphic>
      </xdr:graphicFrame>
      <xdr:clientData/>
    </xdr:absoluteAnchor>
  </xsl:template>

  <xsl:template name="pr">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
      <!--matadata-->
      <xsl:if test ="元:元数据_5200/元:编辑时间_5209|元:元数据_5200/元:创建应用程序_520A|元:元数据_5200/元:文档模板_520C|元:元数据_5200/元:公司名称_5213|元:元数据_5200/元:经理名称_5214|元:元数据_5200/元:页数_5215|元:元数据_5200/元:字数_5216|元:元数据_5200/元:行数_5219|元:元数据_5200/元:段落数_521A">
        <Relationship Id="rId3" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties" Target="docProps/app.xml"/>
      </xsl:if>
      <xsl:if test ="元:元数据_5200/元:标题_5201|元:元数据_5200/元:主题_5202|元:元数据_5200/元:创建者_5203|元:元数据_5200/元:最后作者_5205|元:元数据_5200/元:摘要_5206|元:元数据_5200/元:创建日期_5207|元:元数据_5200/元:编辑次数_5208|元:元数据_5200/元:分类_520B|元:元数据_5200/元:关键字集_520D">
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      </xsl:if>
      
      <!--20130115,gaoyuwei，解决2634BUG"元数据丢失"UOF-OOXML start-->
		<xsl:if test ="元:元数据_5200/元:用户自定义元数据集_520F">
			<Relationship Id="rId4" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/custom-properties" Target="docProps/custom.xml"/>
		</xsl:if>
		<!--end-->
		
    </Relationships>
  </xsl:template>
</xsl:stylesheet>
