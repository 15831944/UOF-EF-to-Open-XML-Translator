<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
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
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" 
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template match="/">
    <!--以UOF的根节点出发-->
    <!--修改预处理式样单  李杨2011-12-28-->
    <uof:UOF>
      <!--uof.xml-->
      <xsl:copy-of select ="document('uof.xml')"/>
      <!--元数据 meta.xml-->
      <xsl:copy-of select="document('_meta/meta.xml')"/>
      <!--公用处理规则 rules.xml-->
      <xsl:copy-of select ="document('rules.xml')"/>

      <!--链接集 hyperlinks.xml-->
      <xsl:if test =".//表:单元格_E7F2/@超链接引用_E7BE">
        <xsl:copy-of select ="document('hyperlinks.xml')"/>
      </xsl:if>

      <!--式样集 styles.xml-->
      <xsl:copy-of select ="document('styles.xml')"/>
      <!--电子表格主体-->
      <!--<xsl:copy-of select ="."/>-->
      
      <!--工作表中背景图片填充-->
      <xsl:if test =".//表:工作表_E825/表:工作表属性_E80D/表:背景填充_E830/图:图片_8005">
        <xsl:copy-of select="document('objectdata.xml')"/>
      </xsl:if>

      <!--图形集 graphics.xml-->
      <xsl:if test =".//uof:锚点_C644/@图形引用_C62E">
        <xsl:copy-of select ="document('graphics.xml')"/>
        
        <!--对象集 objectdata.xml-->
        <!--<xsl:if test="document('objectdata.xml')">-->
        <xsl:if test ="document('graphics.xml')//图:图片数据引用_8037">
          <xsl:copy-of select="document('objectdata.xml')"/>
        </xsl:if>
        <xsl:if test ="document('graphics.xml')//图:填充_804C/图:图片_8005/@图形引用_8007">
          <xsl:copy-of select="document('objectdata.xml')"/>
        </xsl:if>
		  
        <!--图表 chart.xml-->
        <!--<xsl:if test ="document('chart.xml')">-->
        <xsl:if test ="document('graphics.xml')//图:图表数据引用_8065">
          <xsl:copy-of select ="document('chart.xml')"/>
          <!--20130122 gaoyuwei 2642 标题集图片填充的填充内容丢失 start-->
			<xsl:if test ="document('chart.xml')//图表:填充_E746/图:图片_8005/@图形引用_8007">
				<xsl:copy-of select="document('objectdata.xml')"/>
			</xsl:if>
			<!--end-->
        </xsl:if>
      </xsl:if>

      <!--扩展区 extend.xml-->
      <!--<xsl:if test ="document('extend.xml')">
        <xsl:copy-of select ="document('extend.xml')"/>
      </xsl:if>-->
     
      <!--电子表格主体-->
      <表:电子表格文档_E826>
        <!--所有工作表集合-->
        <xsl:for-each select="表:电子表格文档_E826">
          <xsl:if test="./表:工作表_E825[表:工作表内容_E80E]">
            <表:工作表集>
              
              <xsl:for-each select="./表:工作表_E825[表:工作表内容_E80E]">
                <xsl:variable name="sheetNo">
                  <xsl:value-of select="position()"/>
                </xsl:variable>
                <表:单工作表>
                  <xsl:attribute name="uof:sheetNo">
                    <xsl:value-of select="$sheetNo"/>
                  </xsl:attribute>
                  <xsl:copy-of select="."/>
                  <xsl:variable name="sheetID" select="@标识符_E7AC"/>
                  <xsl:variable name="sheetName" select="@名称_E822"/>
                  <xsl:for-each select="document('rules.xml')/规则:公用处理规则_B665">

                    <!--规则:数据有效性集_B618  李杨修改 2012-1-4-->
                    <!--Modified by LDM in 2010/12/13-->
                    <!--此处需要修改，增添了区域集，可以表示同一条件控制下的多个不同区域-->
                    <!--默认区域集中描述的区域属于同一工作表-->
                    <xsl:if test="./规则:电子表格_B66C/规则:数据有效性集_B618">
                      <xsl:for-each select="./规则:电子表格_B66C/规则:数据有效性集_B618/规则:数据有效性_B619">
                        <xsl:variable name="sqref_DV_sub">
                          <xsl:value-of select="规则:区域集_B61A/规则:区域_B61B"/>
                        </xsl:variable>
                        <xsl:variable name="effectSheetName">
                          <xsl:value-of select='translate(substring-before($sqref_DV_sub,"!"),"&apos;","")'/>
                        </xsl:variable>
                        <xsl:if test="$effectSheetName = $sheetName">
                          <!--<表:数据有效性>-->
                          <xsl:copy-of select="."/>
                          <!--</表:数据有效性>-->
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:if>

                    <!--规则:条件格式化集_B628  李杨修改 2012-1-4-->
                    <!--Modified by LDM in 2010/11/30-->
                    <!--表:提取出条件格式化集中的各条件格式设置到对应的工作表中-->
                    <!--UOF1.1中增添了区域集，可以表示同一条件控制下的多个不同区域-->
                    <!--默认区域集中描述的区域属于同一工作表-->
                    <xsl:if test="./规则:电子表格_B66C/规则:条件格式化集_B628">
                      <xsl:for-each select="./规则:电子表格_B66C/规则:条件格式化集_B628">
                        <xsl:for-each select="./规则:条件格式化_B629">
                          <xsl:variable name="sqref_CF_sub">
                            <xsl:value-of select="./规则:区域集_B61A/规则:区域_B61B"/>
                          </xsl:variable>
                          <xsl:variable name="effectSheetName">
                            <!--如何输出单引号？-->
                            <!--Is it right?-->
                            <!--Yes it's a right method to process the single quots-->
                            <xsl:value-of select='translate(substring-before($sqref_CF_sub,"!"),"&apos;","")'/>
                          </xsl:variable>
                          <xsl:if test="$effectSheetName = $sheetName">
                            <xsl:variable name="conditional_count">
                              <xsl:if test="./规则:条件_B62B">
                                <xsl:value-of select="count(./规则:条件_B62B)"/>
                              </xsl:if>
                            </xsl:variable>
                            <xsl:variable name="conditionalFormatID">
                              <xsl:value-of select="position()"/>
                            </xsl:variable>
                            <表:单条件格式化>
                              <xsl:attribute name="effectSheetName">
                                <xsl:value-of select="$effectSheetName"/>
                              </xsl:attribute>
                              <xsl:attribute name="conditionID">
                                <xsl:value-of select="$conditionalFormatID"/>
                              </xsl:attribute>
                              <xsl:attribute name="conditionCount">
                                <xsl:value-of select="$conditional_count"/>
                              </xsl:attribute>
                              <xsl:copy-of select="."/>
                            </表:单条件格式化>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:for-each>
                    </xsl:if>
                  </xsl:for-each>

                  <!--图表集合 修改 李杨2012-2-10-->
                  <xsl:if test =".//uof:锚点_C644[@图形引用_C62E]">
                    <表:图表集合>
                      <xsl:for-each select =".//uof:锚点_C644">
                        <xsl:variable name ="graphID1">
                          <xsl:value-of select ="./@图形引用_C62E"/>
                        </xsl:variable>
                        <xsl:if test ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphID1]/图:图表数据引用_8065">
                          <表:单图表>
                            <xsl:variable name="chartNo" select="position()"/>
                            <xsl:attribute name="uof:chartNo">
                              <xsl:value-of select="$chartNo"/>
                            </xsl:attribute>
                            <xsl:variable name ="chartID">
                              <xsl:value-of select ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphID1]/图:图表数据引用_8065"/>
                            </xsl:variable>

                            <xsl:for-each select ="document('chart.xml')/图表:图表集_E836/图表:图表_E837[@标识符_E828=$chartID]">
                              <xsl:copy-of select ="."/>
                            </xsl:for-each>
                          </表:单图表>
                        </xsl:if>
                      </xsl:for-each>
                    </表:图表集合>
                  </xsl:if>
                 
                  <!--<xsl:if test =".//uof:锚点_C644[@图形引用_C62E]">
                    <表:图表集合>
                      <xsl:for-each select =".//uof:锚点_C644">
                        <xsl:variable name ="graphID1">
                          <xsl:value-of select ="./@图形引用_C62E"/>
                        </xsl:variable>
                        -->
                  <!--图表 chart.xml--><!--
                        <xsl:for-each select ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[图:图表数据引用_8065]">
                          <xsl:variable name="chartNo" select="position()"/>
                          <xsl:if test ="./@标识符_804B=$graphID1">
                            <xsl:variable name ="chartID">
                              <xsl:value-of select ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphID1]/图:图表数据引用_8065"/>
                            </xsl:variable>

                            <xsl:attribute name="uof:sheetNo">
                              <xsl:value-of select="$sheetNo"/>
                            </xsl:attribute>
                            <xsl:if test ="document('chart.xml')/图表:图表集_E836/图表:图表_E837[@标识符_E828=$chartID]">
                              <xsl:for-each select ="document('chart.xml')/图表:图表集_E836/图表:图表_E837[@标识符_E828=$chartID]">
                                <表:单图表>
                                  <xsl:attribute name="uof:chartNo">
                                    <xsl:value-of select="$chartNo"/>
                                  </xsl:attribute>
                                  <xsl:copy-of select="."/>
                                </表:单图表>
                              </xsl:for-each>
                            </xsl:if>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:for-each>
                    </表:图表集合>
                  </xsl:if>-->
                  
                  <!--锚点集合-->
                  <xsl:if test=".//uof:锚点_C644">
                    <uof:锚点集合>
                      <xsl:attribute name="uofsheetNo">
                        <xsl:value-of select="$sheetNo"/>
                      </xsl:attribute>
                      <xsl:for-each select=".//uof:锚点_C644">
                        <xsl:variable name="refID" select="position()"/>
                        <uof:单锚点>
                          <xsl:attribute name="uof:refID">
                            <xsl:value-of select="$refID"/>
                          </xsl:attribute>
                          <xsl:copy-of select="."/>
                        </uof:单锚点>
                      </xsl:for-each>
                    </uof:锚点集合>
                  </xsl:if>

                  <!--估计需要修改，UOF2.0中的批注不在content.xml中 李杨2012-1-4-->
                  <!--<xsl:if test=".//表:批注_E7B7">
                    <xsl:copy-of select=".//表:批注_E7B7"/>
                  </xsl:if>-->
                </表:单工作表>
              </xsl:for-each>
            </表:工作表集>
          </xsl:if>
          <!--chartSheet 修改 李杨2012-5-20-->
          <xsl:if test="./表:工作表_E825[uof:锚点_C644 and not(表:工作表内容_E80E)]">
            <表:图表工作表集>
              <xsl:if test ="./表:工作表_E825/uof:锚点_C644[@图形引用_C62E]">
                  <xsl:for-each select ="./表:工作表_E825/uof:锚点_C644">
                    <xsl:variable name ="graphID1">
                      <xsl:value-of select ="./@图形引用_C62E"/>
                    </xsl:variable>
                    <xsl:if test ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphID1]/图:图表数据引用_8065">
                      <表:图表工作表>
                        <xsl:variable name="chartNo" select="position()"/>
                        <xsl:attribute name="uof:chartNo">
                          <xsl:value-of select="$chartNo"/>
                        </xsl:attribute>
                        <xsl:variable name ="chartID">
                          <xsl:value-of select ="document('graphics.xml')/图形:图形集_7C00/图:图形_8062[@标识符_804B=$graphID1]/图:图表数据引用_8065"/>
                        </xsl:variable>
                        <xsl:copy-of select="ancestor::表:工作表_E825"/>

                        <xsl:for-each select ="document('chart.xml')/图表:图表集_E836/图表:图表_E837[@标识符_E828=$chartID]">
                          <xsl:copy-of select ="."/>
                        </xsl:for-each>
                      </表:图表工作表>
                    </xsl:if>
                  </xsl:for-each>
              </xsl:if>
              <!--<xsl:for-each select="./表:工作表[表:图表]">
                <表:图表工作表>
                  <xsl:for-each select="./@*">
                    <xsl:copy-of select="."/>
                  </xsl:for-each>
                  <xsl:variable name="sheetNo">
                    <xsl:value-of select="translate(@表:标识符,'SHEET_','')"/>
                  </xsl:variable>
                  <xsl:copy-of select="./表:工作表属性"/>
                  <xsl:copy-of select="./表:图表"/>
                </表:图表工作表>
              </xsl:for-each>-->
            </表:图表工作表集>
          </xsl:if>
        </xsl:for-each>
      </表:电子表格文档_E826>

      <xsl:choose>
        <xsl:when test=".//表:数据_E7B3[@类型_E7B6='text' or @类型_E7B6='']">
          <!--修改 李杨2012-1-4-->
          <uof:单元格内容集合>
            <!--把数据格式类型为文本类型的字串提取出来，以便于后期和OOXML对应转换，相当于其中的sharedStrings，其它的格式类型如number类型不单独提取-->
            <xsl:for-each select="//表:数据_E7B3[@类型_E7B6='text' or @类型_E7B6='']">
              <xsl:variable name="refID" select="position()"/>
              <xsl:variable name="rowID" select="ancestor::表:行_E7F1/@行号_E7F3"/>
              <xsl:variable name="colID" select="ancestor::表:单元格_E7F2/@列号_E7BC"/>
              <xsl:variable name="cellID" select="concat($rowID,'_',$colID)"/>
              <xsl:variable name ="sheetId" select="ancestor::表:工作表_E825/@标识符_E7AC"/>
              <uof:单元格内容>
                <xsl:attribute name="uof:cellID">
                  <xsl:value-of select="$cellID"/>
                </xsl:attribute>
                <xsl:attribute name="uof:sheetId">
                  <xsl:value-of select="$sheetId"/>
                </xsl:attribute>
                <xsl:attribute name="uof:refID">
                  <xsl:value-of select="$refID"/>
                </xsl:attribute>
                <xsl:copy-of select=".//字:句_419D/字:文本串_415B"/>
              </uof:单元格内容>
            </xsl:for-each>
          </uof:单元格内容集合>
        </xsl:when>
        <xsl:otherwise>
          <uof:单元格内容集合></uof:单元格内容集合>
        </xsl:otherwise>
      </xsl:choose>
    </uof:UOF>
  </xsl:template>
</xsl:stylesheet>
