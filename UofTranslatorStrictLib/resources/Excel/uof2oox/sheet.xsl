<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing">
  <xsl:import href="hyperlink.xsl"/>

  <!--Not Checked-->
  <!--Marked by LDM in 2011/01/07-->
  <xsl:template name="ChartSheet">
    <xsl:param name="seq"/>
    <chartsheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
      <sheetViews>
        <sheetView tabSelected="1" zoomScale="75" workbookViewId="0" zoomToFit="1"/>
      </sheetViews>
      <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
      <drawing>
        <xsl:attribute name="r:id">
          <xsl:value-of select="concat('rIdChartSheet',$seq)"/>
        </xsl:attribute>
      </drawing>
    </chartsheet>
  </xsl:template>
  <xsl:template name="reldrawing">
    <xsl:param name="goal"/>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/drawing">
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../drawings/drawing',$goal,'.xml')"/>
        </xsl:attribute>
      </Relationship>
    </Relationships>
  </xsl:template>

  <!--模板功能：将列号的数字表示转换成字母表示-->
  <!--最多只能转换以两个字母表示的列，如AF,D等-->
  <!--Modified by LDM in 2010/12/18-->
  <!--ColIndex-->
  <!--2014-3-19，注释掉，hyperlink.xsl中已经定义，此处无需定义， start-->
  <!--<xsl:template name="ColIndex">
    <xsl:param name="colSeq"/>
    <xsl:choose>
      <xsl:when test="$colSeq &lt; 27">
        <xsl:choose>
          <xsl:when test="$colSeq &lt; 10">
            <xsl:value-of select="translate($colSeq,'123456789','ABCDEFGHI')"/>
          </xsl:when>
          <xsl:when test="($colSeq &lt;19) and ($colSeq &gt; 9)">
            <xsl:value-of select="translate($colSeq - 9,'123456789','JKLMNOPQR')"/>
          </xsl:when>
          <xsl:when test="($colSeq &lt;27) and ($colSeq &gt; 18)">
            <xsl:value-of select="translate($colSeq - 18,'12345678','STUVWXYZ')"/>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>-->
  <!--2014-3-19 end-->
  
  <!--工作表模板-->
  <xsl:template name="Sheet">
    <xsl:variable name="sheetNo">
      <xsl:value-of select="./@uof:sheetNo"/>
    </xsl:variable>
    <worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <!--工作表属性-->
      <xsl:if test=".//表:工作表属性_E80D">
        <!--标签背景色-->
        <xsl:if test=".//表:工作表属性_E80D/表:标签背景色_E7C0">
          <xsl:variable name="bgColor">
            <xsl:value-of select=".//表:工作表属性_E80D/表:标签背景色_E7C0"/>
          </xsl:variable>
          <xsl:variable name="tabColor">
            <xsl:value-of select="concat('FF',substring-after($bgColor,'#'))"/>
          </xsl:variable>
          <sheetPr>
            <tabColor>
              <xsl:attribute name="rgb">
                <xsl:value-of select="$tabColor"/>
              </xsl:attribute>
            </tabColor>
          </sheetPr>
        </xsl:if>
		 <!--20130117 gaoyuwei bug2647 “明细在数据下方”，“明细在数据右侧”选中和不选中转换后统一转换为选中 start-->
		  <!--分组集-->
		  <xsl:if test=".//表:分组集_E7F6/@标记是否在数据下方_E838 or .//表:分组集_E7F6/@标记是否在数据下方_E838 ">
			  <sheetPr>

				  <outlinePr>
				  <xsl:if test=".//表:分组集_E7F6/@标记是否在数据下方_E838='false' ">
					  <xsl:attribute name="summaryBelow">
						  <xsl:value-of select="'0'"/>
					  </xsl:attribute>
			  </xsl:if>
			  <xsl:if test=".//表:分组集_E7F6/@标记是否在数据右侧_E839='false' ">
				  <xsl:attribute name="summaryRight">
					  <xsl:value-of select="'0'"/>
				  </xsl:attribute>			
			  </xsl:if>
				  </outlinePr>
			</sheetPr>
		  </xsl:if>
		  
		  <!--end-->        
        
        <!--视图-->
        <xsl:if test=".//表:工作表属性_E80D/表:视图_E7D5">
          <sheetViews>
            <xsl:for-each select=".//表:工作表属性_E80D/表:视图_E7D5">
              <sheetView>
                <xsl:if test ="表:最上行_E7DB and 表:最左列_E7DC">
                  
                  <!--2014-3-26, update by Qihy, topLeftCell取值不正确，修复BUG3153， start-->
                  <!--<xsl:attribute name ="topLeftCell">
                    <xsl:value-of select ="concat(表:最左列_E7DC,表:最上行_E7DB)"/>
                  </xsl:attribute>-->
                  <xsl:variable name="colSeq">
                    <xsl:call-template name="ColIndex">
                      <xsl:with-param name="colSeq" select="表:最左列_E7DC"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:attribute name ="topLeftCell">
                    <xsl:value-of select ="concat($colSeq,表:最上行_E7DB)"/>
                  </xsl:attribute>
                  <!--2014-3-26 end-->
                  
                </xsl:if>
                <xsl:if test="表:是否选中_E7D6='true'">
                  <xsl:attribute name="tabSelected">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:attribute name="workbookViewId">
                  <xsl:value-of select="@窗口标识符_E7E5"/>
                </xsl:attribute>
                <xsl:if test="表:是否显示公式_E7DE='true'">
                  <xsl:attribute name="showFormulas">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--添加 是否显示行号列标。李杨2012-3-18-->
                <xsl:if test ="表:是否显示行号列标_E7E3='false'">
                  <xsl:attribute name ="showRowColHeaders">0</xsl:attribute>
                </xsl:if>
                <xsl:if test="表:是否显示网格_E7DF='true'">
                  <xsl:attribute name="showGridLines">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="表:是否显示网格_E7DF='false'">
                  <xsl:attribute name="showGridLines">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--Modified by LDM in 2011/01/23-->
                <!--网格线颜色-->
                <xsl:if test="表:网格颜色_E7E0">
                  <xsl:attribute name="defaultGridColor">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                  <xsl:attribute name="colorId">
                    <xsl:if test ="表:网格颜色_E7E0='#ff0000'">
                      <xsl:value-of select="'10'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#008000'">
                      <xsl:value-of select ="'11'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#ffff00'">
                      <xsl:value-of select ="'13'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#0000ff'">
                      <xsl:value-of select ="'12'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#cc99ff'">
                      <xsl:value-of select ="'46'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#666699'">
                      <xsl:value-of select ="'54'"/>
                    </xsl:if>
                    <xsl:if test ="表:网格颜色_E7E0='#ff99cc'">
                      <xsl:value-of select ="'45'"/>
                    </xsl:if>
                  </xsl:attribute>
                </xsl:if>

                <xsl:if test="表:当前视图类型_E7DD">
                  <xsl:attribute name="view">
                    <xsl:choose>
                      <xsl:when test="表:当前视图类型_E7DD= 'page'">
                        <xsl:value-of select="'pageBreakPreview'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'normal'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>
                <!--视图的最上行与最左列-->
                <!--Modified by LDM in 2010/12/18
                <xsl:if test="表:最上行 and 表:最左列">
                  <xsl:attribute name="topLeftCell">
                    <xsl:variable name="topRow">
                      <xsl:value-of select="表:最上行"/>
                    </xsl:variable>
                    <xsl:variable name="leftCol">
                      <xsl:value-of select="表:最左列"/>
                    </xsl:variable>
                    <xsl:variable name="leftCol_Alp">
                      <xsl:call-template name="Dec2Ts">
                        <xsl:with-param name="colSeq" select="$leftCol"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select="concat($leftCol_Alp,$topRow)"/>
                  </xsl:attribute>
                </xsl:if>
-->
                <!--Modified by LDM-->
                <!--UOF1.1中缩放为百分比类型-->
                <xsl:if test="表:缩放_E7C4">
                  <xsl:attribute name="zoomScale">
                    <xsl:choose>
                      <xsl:when test="表:当前视图类型_E7DD='page'">
                        <xsl:if test="表:分页缩放_E7E1">
                          <xsl:value-of select="表:分页缩放_E7E1"/>
                        </xsl:if>
                        <xsl:if test="not(表:分页缩放_E7E1)">
                          <xsl:value-of select="表:缩放_E7C4"/>
                        </xsl:if>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="表:缩放_E7C4"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="表:分页缩放_E7E1">
                  <xsl:attribute name="zoomScalePageLayoutView">
                    <xsl:value-of select="表:分页缩放_E7E1"/>
                  </xsl:attribute>
                </xsl:if>
                <!--Need Consideration Codes-->
                <!--Marked by LDM in 2010/12/18-->
                <xsl:if test="表:拆分_E7D7">
                  <pane>
                    <xsl:variable name="width">
                      <xsl:value-of select="表:拆分_E7D7/@宽_C605"/>
                    </xsl:variable>
                    <xsl:variable name="height">
                      <xsl:value-of select="表:拆分_E7D7/@长_C604"/>
                    </xsl:variable>
                    <xsl:variable name="colSeq">
                      <xsl:value-of select="round(($width - 22) div 40)"/>
                    </xsl:variable>
                    <xsl:variable name="rowSeq">
                      <xsl:value-of select="round(($height - 10 )div  11)"/>
                    </xsl:variable>
                    <xsl:variable name="colSeq_Alp">
                      <xsl:call-template name="ColIndex">
                        <xsl:with-param name="colSeq" select="$colSeq"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:attribute name="xSplit">
                      <xsl:value-of select="20*(24 + $colSeq  * 56)"/>
                    </xsl:attribute>
                    <xsl:attribute name="ySplit">
                      <xsl:value-of select="20*(13 + $rowSeq * 14 )"/>
                    </xsl:attribute>
                    <xsl:attribute name="topLeftCell">
                      <xsl:value-of select="concat($colSeq_Alp,$rowSeq + 1)"/>
                    </xsl:attribute>
                  </pane>
                </xsl:if>
                <xsl:if test="表:冻结_E7D8">
                  <xsl:variable name="colSeq">
                    <xsl:value-of select="表:冻结_E7D8/@列号_E7DA"/>
                  </xsl:variable>
                  <xsl:variable name="rowSeq">
                    <xsl:value-of select="表:冻结_E7D8/@行号_E7D9"/>
                  </xsl:variable>
                  
                  <!--2014-5-5, update by Qihy, bug3303, 冻结测试转换后打开需要恢复， start-->
                  <xsl:variable name="colSeq1" select="表:冻结_E7D8/@最左列_E83E"/>
                  <xsl:variable name="colSeq_Alp">
                    <xsl:call-template name="ColIndex">
                      <xsl:with-param name="colSeq" select="$colSeq1"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <!--2014-5-5 end-->
                  
                  <xsl:if test="表:冻结_E7D8[@列号_E7DA and  @行号_E7D9]">
                    <pane>
                      <xsl:attribute name="xSplit">
                        <xsl:value-of select="$colSeq"/>
                      </xsl:attribute>
                      <xsl:attribute name="ySplit">
                        <xsl:value-of select="$rowSeq"/>
                      </xsl:attribute>
                      <xsl:attribute name="topLeftCell">
                        <xsl:value-of select="concat($colSeq_Alp,$rowSeq + 1)"/>
                      </xsl:attribute>
                      <xsl:attribute name="state">
                        <xsl:value-of select="'frozen'"/>
                      </xsl:attribute>
                    </pane>
                  </xsl:if>
                  <xsl:if test="表:冻结_E7D8[@列号_E7DA and not(@行号_E7D9)]">
                    <pane>
                      <xsl:attribute name="xSplit">
                        <xsl:value-of select="$colSeq"/>
                      </xsl:attribute>
                      <xsl:attribute name="topLeftCell">
                        <xsl:value-of select="concat($colSeq_Alp, 1)"/>
                      </xsl:attribute>
                      <xsl:attribute name="state">
                        <xsl:value-of select="'frozen'"/>
                      </xsl:attribute>
                    </pane>
                  </xsl:if>
                  <xsl:if test="表:冻结_E7D8[@行号_E7D9 and not(@列号_E7DA)]">
                    <pane>
                      <xsl:attribute name="ySplit">
                        
                        <!--2014-3-26 update by Qihy, ySplit取值不正确, start-->
                      <!--20130205 gaoyuwei  bug_2706 FreezeRow1and2冻结丢失 start--><!--
						  <xsl:value-of select="$rowSeq"/>
						  --><!--end-->
                        <xsl:choose>
                          <xsl:when test ="表:最上行_E7DB and 表:最上行_E7DB != '1'">
                            <xsl:value-of select="$rowSeq - 表:最上行_E7DB  + 1"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$rowSeq"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <!--2014-3-26 end-->
                      
                      <xsl:attribute name="topLeftCell">
                        <xsl:value-of select="concat('A',$rowSeq + 1)"/>
                      </xsl:attribute>
                      <xsl:attribute name="state">
                        <xsl:value-of select="'frozen'"/>
                      </xsl:attribute>
                    </pane>
                  </xsl:if>
                  <!--<xsl:if test="表:选中区域_E7E2">
                    <xsl:variable name="selectedCell">
                      <xsl:value-of select="表:选中区域_E7E2"/>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="contains($selectedCell,':')">
                        <xsl:variable name="sCStartTemp">
                          <xsl:value-of select="substring-before($selectedCell,':')"/>
                        </xsl:variable>
                        <xsl:variable name="sCEndTemp">
                          <xsl:value-of select="substring-after($selectedCell,':')"/>
                        </xsl:variable>
                        <xsl:variable name="sCTemp">
                          <xsl:value-of select="translate($selectedCell,'$','')"/>
                        </xsl:variable>
                        <xsl:variable name="sCEnd">
                          <xsl:value-of select="translate($sCEndTemp,'$','')"/>
                        </xsl:variable>
                        <selection>
                          <xsl:attribute name="activeCell">
                            <xsl:value-of select="$sCEnd"/>
                          </xsl:attribute>
                          <xsl:attribute name="sqref">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                        </selection>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:variable name="sCTemp">
                          <xsl:value-of select="translate($selectedCell,'$','')"/>
                        </xsl:variable>
                        <selection>
                          <xsl:attribute name="activeCell">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                          <xsl:attribute name="sqref">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                        </selection>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>-->
                </xsl:if>
                <!--添加 是否显示零值 属性。李杨2012-3-18-->
                <xsl:if test ="表:是否显示零值_E7E4='false'">
                  <xsl:attribute name ="showZeros">0</xsl:attribute>
                </xsl:if>
                <!--添加 选中区域。李杨2012-3-18-->
                <xsl:if test="表:选中区域_E7E2">
                    <xsl:variable name="selectedCell">
                      <xsl:value-of select="表:选中区域_E7E2"/>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="contains($selectedCell,':')">
                        <xsl:variable name="sCStartTemp">
                          <xsl:value-of select="substring-before($selectedCell,':')"/>
                        </xsl:variable>
                        <xsl:variable name="sCEndTemp">
                          <xsl:value-of select="substring-after($selectedCell,':')"/>
                        </xsl:variable>
                        <xsl:variable name="sCTemp">
                          <xsl:value-of select="translate($selectedCell,'$','')"/>
                        </xsl:variable>
                        <xsl:variable name="sCEnd">
                          <xsl:value-of select="translate($sCEndTemp,'$','')"/>
                        </xsl:variable>
                        <selection>
                          <xsl:attribute name="activeCell">
                            <xsl:value-of select="$sCEnd"/>
                          </xsl:attribute>
                          
                          <!--yanghaojie-temp-start-->
                          <!--
                          <xsl:if test="../../../表:工作表_E825/@名称_E822='工作表设置' and ../../../表:工作表_E825/@标识符_E7AC='SHEET_4'">
                            <xsl:attribute name="activeCell">
                              <xsl:value-of select="concat('A',$sCEnd)"/>
                            </xsl:attribute>
                          </xsl:if>
                          -->
                          <!--yanghaojie-temp-end-->
                          <xsl:attribute name="sqref">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                        </selection>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:variable name="sCTemp">
                          <xsl:value-of select="translate($selectedCell,'$','')"/>
                        </xsl:variable>
                        <selection>
                          <xsl:attribute name="activeCell">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                          <!--yanghaojie-temp-start-->
                          <!--
                          <xsl:if test="../../../表:工作表_E825/@名称_E822='工作表设置' and ../../../表:工作表_E825/@标识符_E7AC='SHEET_4'">
                            <xsl:attribute name="activeCell">
                              <xsl:value-of select="concat('A',$sCTemp)"/>
                            </xsl:attribute>
                          </xsl:if>-->
                          <!--yanghaojie-temp-end-->
                          <xsl:attribute name="sqref">
                            <xsl:value-of select="$sCTemp"/>
                          </xsl:attribute>
                        </selection>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                <!--<xsl:if test ="表:选中区域_E7E2">
                  <selection>
                    <xsl:variable name ="v">
                      <xsl:value-of select ="表:选中区域_E7E2"/>
                    </xsl:variable>
                    <xsl:if test ="contains($v,':')">
                      --><!--选中的开始单元格--><!--
                      <xsl:variable name ="be">
                        <xsl:value-of select ="substring-before($v,':')"/>
                      </xsl:variable>
                      <xsl:variable name ="be2" select ="substring-after($be,'R')"/>
                      <xsl:variable name ="beR" select ="substring-before($be2,'C')"/>
                      <xsl:variable name ="beC2" select ="substring-after($be2,'C')"/>
                      <xsl:variable name ="beC">
                        <xsl:call-template name ="ColumnTran">
                          <xsl:with-param name ="Column" select ="$beC2"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:variable name ="before" select ="concat($beC,$beR)"/>
                      --><!--选中的结束单元格--><!--
                      <xsl:variable name="af">
                        <xsl:value-of select ="substring-after($v,':')"/>
                      </xsl:variable>
                      <xsl:variable name ="af2" select ="substring-after($af,'R')"/>
                      <xsl:variable name ="afR" select ="substring-before($af2,'C')"/>
                      <xsl:variable name ="afC2" select ="substring-after($af2,'C')"/>
                      <xsl:variable name ="afC">
                        <xsl:call-template name ="ColumnTran">
                          <xsl:with-param name ="Column" select ="$afC2"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:variable name ="after" select ="concat($afC,$afR)"/>

                      <xsl:attribute name ="activeCell">
                        <xsl:value-of select ="$before"/>
                      </xsl:attribute>
                      <xsl:attribute name ="sqref">
                        <xsl:value-of select ="concat($before,':',$after)"/>
                      </xsl:attribute>
                    </xsl:if>

                    <xsl:if test ="not(contains($v,':'))">
                      <xsl:variable name ="vb" select ="substring-after($v,'R')"/>
                      <xsl:variable name ="vbR" select ="substring-before($vb,'C')"/>
                      <xsl:variable name ="vbC2" select ="substring-after($vb,'C')"/>
                      <xsl:variable name ="vbC">
                        <xsl:call-template name ="ColumnTran">
                          <xsl:with-param name ="Column" select ="$vbC2"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:attribute name ="activeCell">
                        <xsl:value-of select ="concat($vbC,$vbR)"/>
                      </xsl:attribute>
                      <xsl:attribute name ="sqref">
                        <xsl:value-of select ="concat($vbC,$vbR)"/>
                      </xsl:attribute>
                    </xsl:if>
                  </selection>
                </xsl:if>-->
              </sheetView>
            </xsl:for-each>
          </sheetViews>
        </xsl:if>
      </xsl:if>

      <xsl:if test=".//表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省行高_E7EA or 表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省列宽_E7EB">
        <sheetFormatPr>
          <xsl:variable name="defaultRowHeight">
            <xsl:value-of select=".//表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省行高_E7EA"/>
          </xsl:variable>
          <xsl:attribute name="defaultRowHeight">
            <xsl:value-of select="$defaultRowHeight"/>
          </xsl:attribute>
          <xsl:attribute name="customHeight">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
          <xsl:variable name="defaultColWidth">
            <xsl:value-of select=".//表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省列宽_E7EB"/>
          </xsl:variable>
          <xsl:attribute name="defaultColWidth">
            <!--
            <xsl:value-of select="$defaultColWidth"/>
            -->
            <!--
            <xsl:value-of select="$defaultColWidth div 54 * 9 + 1.7"/>
            -->
            <xsl:value-of select="$defaultColWidth div 54 * 9"/>

          </xsl:attribute>
        </sheetFormatPr>
      </xsl:if>
      <!--列设置-->
      <xsl:if test=".//表:工作表内容_E80E[表:列_E7EC] or .//表:工作表内容_E80E[列]">
        <cols>
          <xsl:for-each select=".//表:工作表内容_E80E/表:列_E7EC">
            <col>
              <xsl:if test="@列号_E7ED">
                <xsl:variable name="colNo">
                  <xsl:if test="@列号_E7ED">
                    <xsl:value-of select="@列号_E7ED"/>
                  </xsl:if>
                  <!--<xsl:if test="@列号">
                    <xsl:value-of select="@列号"/>
                  </xsl:if>-->
                </xsl:variable>

                <xsl:attribute name="min">
                  <xsl:value-of select="$colNo"/>
                </xsl:attribute>
                <xsl:if test ="not(@跨度_E7EF)">
                  <xsl:attribute name="max">
                    <xsl:value-of select="$colNo"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test ="@跨度_E7EF">
                  <xsl:attribute name ="max">
                    <xsl:value-of select ="$colNo + @跨度_E7EF"/>
                  </xsl:attribute>
                </xsl:if>
                
                <!--添加分组集  李杨2012-2-27-->
                <xsl:if test ="./@outlineLevel">
                  <xsl:attribute name ="outlineLevel">
                    <xsl:value-of select ="./@outlineLevel"/>
                  </xsl:attribute>
                </xsl:if>

                <!--<xsl:if test="./@level and ./@level!='0'">
                  <xsl:attribute name="outlineLevel">
                    <xsl:value-of select="./@level"/>
                  </xsl:attribute>
                </xsl:if>-->
              </xsl:if>
              <xsl:if test="./@是否隐藏_E73C='true'">
                <xsl:attribute name="hidden">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@列宽_E7EE">
                <xsl:variable name="width" select="@列宽_E7EE"/>
                <!--zl 2015-1-6 start-->
                <xsl:attribute name="width">
                  <xsl:choose>
                    <xsl:when test="@列宽_E7EE &gt;= 7 and @列宽_E7EE &lt; 8">
                      <xsl:value-of select="'1.83203125'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 39 and @列宽_E7EE &lt; 40">
                      <xsl:value-of select="'7.1640625'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 41 and @列宽_E7EE &lt; 42">
                      <xsl:value-of select="'7.2'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 45 and @列宽_E7EE &lt; 46">
                      <xsl:value-of select="'7.375'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 51 and @列宽_E7EE &lt; 52">
                      <xsl:value-of select="'11.6640625'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 72 and @列宽_E7EE &lt; 73">
                      <xsl:value-of select="'9'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 76 and @列宽_E7EE &lt; 77">
                      <xsl:value-of select="'12.75'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 80 and @列宽_E7EE &lt; 81">
                      <xsl:value-of select="'13.375'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 93 and @列宽_E7EE &lt; 94">
                      <xsl:value-of select="'15.625'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 105 and @列宽_E7EE &lt; 106">
                      <xsl:value-of select="'26.1640625'"/>
                    </xsl:when>
                    <xsl:when test="@列宽_E7EE &gt;= 245 and @列宽_E7EE &lt; 246">
                      <xsl:value-of select="'39.375'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="@列宽_E7EE div 0.254 div 54 * 1.9"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <!--zl end-->
                <xsl:attribute name="customWidth">
                  <xsl:value-of select="1"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(@列宽_E7EE)">
                <xsl:if test="./parent::表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省列宽_E7EB">
                  <xsl:attribute name="width">
                    <xsl:value-of select="./parent::表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省列宽_E7EB div 6"/>
                  </xsl:attribute>
                </xsl:if>

              </xsl:if>
              
              <!--2014-3-26, update by Qihy, bestFit取值不正确， start-->
              <!--<xsl:if test="./@是否预定义图形蓝色填充-liyang.uos_E7F0='true'">-->
              <xsl:if test="@是否自适应列宽_E7F0 = 'true'">
              <!--2014-3-26 end-->
                
                <xsl:attribute name="bestFit">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>

              <!--yanghaojie-temp-start-->
              <!--2014-6-9, add by Qihy, 修复单元格拉伸时单元格式样不正确导致单元格宽度变化和图表图形变化等问题， start-->
              <!--
              <xsl:if test="@式样引用_E7BD">
                <xsl:attribute name="style">
                  <xsl:value-of select="substring-after(@式样引用_E7BD,'CELLSTYLE_')"/>
                </xsl:attribute>
              </xsl:if>-->
              <!--2014-6-9 end-->
              <!--yanghaojie-temp-end-->
              
            </col>
          </xsl:for-each>
        </cols>
      </xsl:if>
      <!--工作表内容-->
      <sheetData>
        <xsl:for-each select=".//表:工作表内容_E80E/表:行_E7F1">
          <row>
            <xsl:if test="@行号_E7F3">
              <xsl:attribute name="r">
                <xsl:value-of select="@行号_E7F3"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@行高_E7F4">
              <xsl:attribute name="ht">
                <xsl:value-of select="@行高_E7F4"/>
              </xsl:attribute>
              <xsl:attribute name="customHeight">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(@行高_E7F4)">
              <xsl:attribute name="ht">
                <xsl:value-of select="ancestor::表:工作表内容_E80E/表:缺省行高列宽_E7E9/@缺省行高_E7EA"/>
              </xsl:attribute>
              <xsl:attribute name="customHeight">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@是否隐藏_E73C='true'">
              <xsl:attribute name="hidden">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <!--李杨添加分组集行分组代码  2012-2-28-->
            <xsl:if test ="./@outlineLevel">
              <xsl:attribute name ="outlineLevel">
                <xsl:value-of select ="./@outlineLevel"/>
              </xsl:attribute>
            </xsl:if>
            <!--<xsl:if test="@level and not(@level='0')">
              <xsl:attribute name="outlineLevel">
                <xsl:value-of select="./@level"/>
              </xsl:attribute>
            </xsl:if>-->
            <xsl:apply-templates select="./表:单元格_E7F2"/>
          </row>
        </xsl:for-each>
      </sheetData>
      <!-- add xuzhenwei 2012-12-10 新增功能点：用户方案集 start -->
      <xsl:if test=".//表:工作表内容_E80E/表:方案集_E7F8">
        <scenarios>
          <xsl:for-each select=".//表:工作表内容_E80E/表:方案集_E7F8/表:方案_E7F9">
            <scenario>
              <xsl:if test="@名称_E774">
                <xsl:attribute name="name">
                  <xsl:value-of select="@名称_E774"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@是否防止更改_E800">
                <xsl:attribute name="locked">
                  <xsl:choose>
                    <xsl:when test="@是否防止更改_E800">
                      <xsl:value-of select="'1'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@是否隐藏_E801">
                <xsl:attribute name="hidden">
                  <xsl:choose>
                    <xsl:when test="@是否隐藏_E801">
                      <xsl:value-of select="'1'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="@备注_E7FE">
                <xsl:attribute name="comment">
                  <xsl:value-of select="@备注_E7FE"/>
                </xsl:attribute>
              </xsl:if>
              <!-- OOXML的单元格内容 -->
              <xsl:if test="表:可变单元格_E7FA">
                <xsl:for-each select="表:可变单元格_E7FA">
                  <xsl:variable name="cell_id" select="@行号_E7FB"/>
                  <inputCells>
                    <!--设置单元格属性-->
                    <xsl:attribute name="val">
                      <xsl:value-of select="@值_E7FD"/>
                    </xsl:attribute>
                    <xsl:if test="@列号_E7FC">
                      <xsl:variable name="colNo">
                        <xsl:value-of select="@列号_E7FC +1"/>
                      </xsl:variable>
                      <!-- 数字列号转换字母列号。2012-2-10-->
                      <xsl:variable name="colSeq">
                        <xsl:call-template name="ColIndex">
                          <xsl:with-param name="colSeq" select="$colNo"/>
                        </xsl:call-template>
                      </xsl:variable>
                      <xsl:variable name="rowNo" select="@行号_E7FB +1"/>
                      <xsl:attribute name="r">
                        <xsl:value-of select="concat($colSeq,$rowNo)"/>
                      </xsl:attribute>
                    </xsl:if>
                  </inputCells>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test=".//表:批注_E7B7">
                <Relationship>
                  <xsl:attribute name="Id">
                    <xsl:value-of select ="'rId_Comment1'"/>
                    <!--<xsl:value-of select="concat('rId_Comment',$sheetNo)"/>-->
                  </xsl:attribute>
                  <xsl:attribute name="Type">
                    <xsl:value-of select ="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
                    <!--<xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/comments'"/>-->
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:value-of select ="concat('../drawings/drawing',$sheetNo,'.xml')"/>
                    <!--<xsl:value-of select="concat('../comments',$sheetNo,'.xml')"/>-->
                  </xsl:attribute>
                </Relationship>
              </xsl:if>
            </scenario>
          </xsl:for-each>
        </scenarios>
      </xsl:if>
      <!-- end -->
      <!--Not Finished-->
      <!--筛选-->
      <!--修改 李杨2012-2-24-->
      <xsl:if test="./表:工作表_E825/表:筛选集_E83A/表:筛选_E80F">
        <xsl:for-each select="./表:工作表_E825/表:筛选集_E83A/表:筛选_E80F">
          <autoFilter>
            <xsl:variable name="att" select="./表:范围_E810"/>
            <xsl:variable name="att2">
              <xsl:value-of select="substring-after($att,'!')"/>
            </xsl:variable>
            <xsl:variable name="att3">
              <xsl:value-of select="translate($att2,'$','')"/>
            </xsl:variable>
            <xsl:attribute name="ref">
              <xsl:choose>
                <xsl:when test ="contains($att,'!')">
                  <xsl:value-of select="$att3"/>
                </xsl:when>
                <xsl:otherwise >
                  <xsl:value-of select ="$att"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:for-each select="./表:条件_E811">
              <filterColumn>
                <xsl:attribute name="colId">
                  <xsl:value-of select="number(./@列号_E819)-1"/>
                  <!--@表:列号-->
                </xsl:attribute>
                <!-- 20130518 add by xuzhenwei 修改数据过滤，start -->
                <xsl:choose>
                  <xsl:when test="表:自定义_E814[@类型_E83C] and 表:自定义_E814/表:操作条件_E815">
                    <customFilters>
                      <xsl:if test="表:自定义_E814[@类型_E83C='and']">
                        <xsl:attribute name="and">
                          <xsl:value-of select="'1'"/>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:for-each select="表:自定义_E814/表:操作条件_E815">
                        <customFilter>
                          <xsl:variable name="v" select="表:值_E817"/>
                          <xsl:if test="表:操作码_E816 = 'equal-to'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'equal'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'greater-than'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'greaterThan'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'greater-than-or-equal-to'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'greaterThanOrEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'less-than'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'lessThan'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'less-than-or-equal-to'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'lessThanOrEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'not-equal-to'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'notEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="$v"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'contain'">
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat('*',$v,'*')"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'not-contain'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'notEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat('*',$v,'*')"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'start-with'">
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat($v,'*')"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'not-start-with'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'notEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat($v,'*')"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'end-with'">
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat('*',$v)"/>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="表:操作码_E816 = 'not-end-with'">
                            <xsl:attribute name="operator">
                              <xsl:value-of select="'notEqual'"/>
                            </xsl:attribute>
                            <xsl:attribute name="val">
                              <xsl:value-of select="concat('*',$v)"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--待修改  李杨2012-2-24-->
                          <!--<xsl:if test="表:操作码_E816 = 'between'">
                          <xsl:attribute name="operator">
                            <xsl:value-of select="'between'"/>
                          </xsl:attribute>
                          <xsl:attribute name="val">
                            <xsl:value-of select="$v"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="表:操作码_E816 = 'not-between'">
                        <xsl:attribute name="operator">
                          <xsl:value-of select="'between'"/>
                        </xsl:attribute>
                        <xsl:attribute name="val">
                          <xsl:value-of select="$v"/>
                        </xsl:attribute>
                      </xsl:if>-->
                        </customFilter>
                      </xsl:for-each>
                    </customFilters>
                  </xsl:when>
                  <xsl:when test="表:普通_E812[@类型_E7B6='value']">
                    <filters>
                      <xsl:for-each select="表:普通_E812">
                        <filter>
                          <xsl:attribute name="val">
                            <xsl:value-of select="@值_E813"/>
                          </xsl:attribute>
                        </filter>
                      </xsl:for-each>
                    </filters>
                  </xsl:when>
                  <xsl:when test="表:普通_E812[@类型_E7B6='top-item']">
                    <top10>
                      <xsl:attribute name="val">
                        <xsl:value-of select="表:普通_E812/@值_E813"/>
                      </xsl:attribute>
                    </top10>
                  </xsl:when>
                  <!--可能会有问题  李杨2012-2-24-->
                  <xsl:when test="表:颜色_E818">
                    <colorFilter dxfId="0">
                      <xsl:attribute name="dxfId">
                        <xsl:value-of select ="number(count(/uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:条件格式化集_B628/规则:条件格式化_B629/规则:条件_B62B)) + number(count(./ancestor/表:筛选_E80F[表:条件_E811/表:颜色_E818]))"/>
                        <!--<xsl:value-of select="number(count(preceding::规则:条件格式化_B629/规则:条件_B62B)) + number(count(./ancestor/表:筛选_E80F[表:条件_E811/表:颜色_E818]))"/>-->
                      </xsl:attribute>
                    </colorFilter>
                  </xsl:when>
                  <xsl:otherwise></xsl:otherwise>
                </xsl:choose>
                <!-- end -->
              </filterColumn>
            </xsl:for-each>
          </autoFilter>
        </xsl:for-each>
      </xsl:if>
      <!--合并-->
      <xsl:if test=".//表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2/表:合并_E7AF[@列数_E7B0] or 表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2/表:合并_E7AF[@行数_E7B1]">
        <mergeCells>
          <xsl:attribute name="count">
            <xsl:value-of select="count(.//表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2/表:合并_E7AF[@列数_E7B0 or @行数_E7B1])"/>
          </xsl:attribute>
          <xsl:for-each select=".//表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2/表:合并_E7AF[@列数_E7B0 or @行数_E7B1]">
            <mergeCell>
              <xsl:variable name="rowStart">
                <xsl:value-of select="../../@行号_E7F3"/>
              </xsl:variable>
              
              <xsl:variable name="rowInc">
                <xsl:if test="./@行数_E7B1">
                  <!--<xsl:value-of select="./@行数_E7B1"/>-->
                  <xsl:value-of select="./@行数_E7B1"/>
                </xsl:if>
                <xsl:if test="not(./@行数_E7B1)">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:variable>
              
              <xsl:variable name="rowEnd">
                <xsl:value-of select="number($rowStart) + number($rowInc)"/>
              </xsl:variable>

              <xsl:variable name="colStart_temp">
                <xsl:value-of select="../@列号_E7BC"/>
              </xsl:variable>
              <xsl:variable name="colInc">
                <xsl:if test ="./@列数_E7B0">
                  <!--<xsl:value-of select="./@列数_E7B0"/>-->
                  <xsl:value-of select="./@列数_E7B0"/>
                </xsl:if>
                <xsl:if test ="not(./@列数_E7B0)">
                  <xsl:value-of select ="'0'"/>
                </xsl:if>
              </xsl:variable>
              <xsl:variable name="colEnd_temp">
                <xsl:value-of select="number($colStart_temp) + number($colInc)"/>
              </xsl:variable>

              <xsl:variable name="colStart">
                <xsl:call-template name="ColIndex">
                  <xsl:with-param name="colSeq" select="$colStart_temp"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:variable name="colEnd">
                <xsl:call-template name="ColIndex">
                  <xsl:with-param name="colSeq" select="$colEnd_temp"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:variable name ="cellsRef">
                <xsl:value-of select="concat($colStart,$rowStart,':',$colEnd,$rowEnd)"/>
              </xsl:variable>
              <xsl:attribute name="ref">
                <xsl:value-of select="$cellsRef"/>
              </xsl:attribute>
            </mergeCell>
          </xsl:for-each>
        </mergeCells>
      </xsl:if>
      <!--fenye-->
      <!--Not Finished-->
      <!--Not Finished-->
      <!--条件格式化-->
      <xsl:if test="表:单条件格式化">
        <xsl:for-each select="表:单条件格式化">
          
          <!--2014-3-27, update by Qihy, dxfId取值不正确， start -->
          <xsl:variable name="count1">
            <xsl:call-template name="computeCountN">
              <xsl:with-param name="sum" select ="0"/>
              <xsl:with-param name="id" select="@conditionID"/>
            </xsl:call-template>
          </xsl:variable>
          <!--2014-3-27 end-->

          <!--20130408 gaoyuwei bug_2665_1 格式化条件丢失 start-->
			<!--<xsl:variable name="position">
				<xsl:value-of select="./@conditionID - 1"/>
			</xsl:variable>-->			
			<!--bug_2665_1 end-->	
          
          <!--2014-3-27， add by Qihy, 修复bug3151， start-->
          
			    <!--yanghaojie-temp-start-->
          <xsl:for-each select="规则:条件格式化_B629">
          <!--yanghaojie-temp-end-->
            <conditionalFormatting>
              <xsl:variable name="refArea">
                <xsl:choose>
                  <xsl:when test="./ancestor::规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否RC引用_B634='false'">
                    <xsl:value-of select="translate(substring-after(./规则:区域集_B61A/规则:区域_B61B,'!'),'$','')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="./规则:区域集_B61A/规则:区域_B61B"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>

              <xsl:variable name ="n">
                <xsl:apply-templates select="preceding-sibling::规则:条件格式化_B629"/>
              </xsl:variable>

              <xsl:attribute name="sqref">
                <xsl:value-of select="translate(substring-after($refArea,'!'),'$','')"/>
              </xsl:attribute>

              <xsl:for-each select="规则:条件_B62B">
                <xsl:variable name="count2">
                  <xsl:value-of select="position()"/>
                </xsl:variable>
                <cfRule>
                  <xsl:if test="@类型_B673">
                    <xsl:attribute name="type">
                      <xsl:choose>
                        <xsl:when test="@类型_B673='cell-value'">
                          <xsl:value-of select="'cellIs'"/>
                        </xsl:when>
                        <xsl:when test="@类型_B673='formula'">
                          <xsl:value-of select="'expression'"/>
                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="规则:格式_B62D">

                    <xsl:attribute name="dxfId">

                      <!--2014-3-27, update by Qihy, dxfId取值不正确， start -->
                      <!--20130408 gaoyuwei bug_2665_2 格式化条件丢失 start-->
                      <!--<xsl:value-of select="$position"/>-->
                      <!--end-->
                      <xsl:value-of select="$count1 + $count2 - 1"/>
                      <!--2014-3-27 end-->

                    </xsl:attribute>
                    <xsl:attribute name="priority">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="@类型_B673='cell-value' or  @类型_B673='containsText'">
                    <xsl:if test="规则:操作码_B62C">
                      <xsl:attribute name="operator">
                        <xsl:call-template name="opera2"/>
                      </xsl:attribute>
                    </xsl:if>
                  </xsl:if>
                  <xsl:if test="@类型_B673='containsText'">
                    <xsl:attribute name="text">
                      <xsl:value-of select="'Hesse'"/>
                    </xsl:attribute>
                  </xsl:if>

                  <!--2014-3-24, update by Qihy, 修复BUG3110中实例文档MS-CDEPLOY_RequirementSpecification.xlsx互操作需要修复的问题-->
                  <!--<xsl:if test="规则:第一操作数_B61E">
                    <formula>
                      <xsl:value-of select="translate(规则:第一操作数_B61E,'=','')"/>
                    </formula>
                  </xsl:if>
                  <xsl:if test="规则:第二操作数_B61F">
                    <xsl:if test="not(contains(规则:第二操作数_B61F,'='))">
                      <formula>
                        <xsl:value-of select="规则:第二操作数_B61F"/>
                      </formula>
                    </xsl:if>
                    <xsl:if test="contains(规则:第二操作数_B61F,'=')">
                    
                      -->
                  <!--20130408 gaoyuwei bug_2665_3 格式化条件丢失 start-->
                  <!--
						          <formula>
							          <xsl:value-of select="translate(规则:第二操作数_B61F,'=','')"/>
						          </formula>
					            -->
                  <!--end-->
                  <!--
						
                    </xsl:if>-->

                  <!--<xsl:if test="规则:第一操作数_B61E">
                    <formula>
                      <xsl:value-of select="substring-after(规则:第一操作数_B61E,'=')"/>
                    </formula>
                  </xsl:if>
                  <xsl:if test="规则:第二操作数_B61F">
                    <xsl:if test="not(contains(规则:第二操作数_B61F,'='))">
                      <formula>
                        <xsl:value-of select="规则:第二操作数_B61F"/>
                      </formula>
                    </xsl:if>
                    <xsl:if test="contains(规则:第二操作数_B61F,'=')">
                      <formula>
                        <xsl:value-of select="substring-after(规则:第二操作数_B61F,'=')"/>
                      </formula>
                    </xsl:if>-->
                  <xsl:if test="规则:第一操作数_B61E">
                    <formula>
                      <xsl:choose>
                        <xsl:when test="contains(规则:第一操作数_B61E,'=')">
                          <xsl:value-of select="substring-after(规则:第一操作数_B61E,'=')"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="规则:第一操作数_B61E"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </formula>
                  </xsl:if>
                  <xsl:if test="规则:第二操作数_B61F">
                    <!--zl start-->
                    <formula>
                    <xsl:choose>
                      <xsl:when test="contains(规则:第二操作数_B61F,'=')">
                        <xsl:value-of select="substring-after(规则:第二操作数_B61F,'=')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="规则:第二操作数_B61F"/>
                      </xsl:otherwise>
                    </xsl:choose>
                    </formula>
                    <!--end zl-->
                    <!--2014-3-24 end-->

                  </xsl:if>
                </cfRule>
              </xsl:for-each>
            </conditionalFormatting>
          </xsl:for-each>
          
        </xsl:for-each>
      </xsl:if>
      <!---->
      <!--dataValidations-->
      <xsl:if test=".//规则:数据有效性_B619 and not(.//规则:数据有效性_B619/规则:校验类型_B61C='custom')">
        <dataValidations>
          <xsl:variable name="c" select="count(.//规则:数据有效性_B619)"/>
          <xsl:attribute name="count">
            <xsl:value-of select="$c"/>
          </xsl:attribute>
          <xsl:for-each select="./规则:数据有效性_B619">
            <dataValidation>
              <!--type-->
              <xsl:if test="规则:校验类型_B61C">
                <xsl:variable name="validationType" select="规则:校验类型_B61C"/>
                <xsl:attribute name="type">
                  <xsl:choose>
                    <xsl:when test="$validationType='any-value'">
                      <xsl:value-of select="'none'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='whole-number'">
                      <xsl:value-of select="'whole'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='decimal'">
                      <xsl:value-of select="'decimal'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='list'">
                      <xsl:value-of select="'list'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='date'">
                      <xsl:value-of select="'date'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='time'">
                      <xsl:value-of select="'time'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='text-length'">
                      <xsl:value-of select="'textLength'"/>
                    </xsl:when>
                    <xsl:when test="$validationType='custom'">
                      <xsl:value-of select="'custom'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'none'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </xsl:if>
              <!--allowBlank-->
              <xsl:if test="规则:是否忽略空格_B620">
                <xsl:if test="规则:是否忽略空格_B620='true'">
                  <xsl:attribute name="allowBlank">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <xsl:if test="规则:错误提示_B626">
                <!--errorStyle-->
                <xsl:if test="规则:错误提示_B626[@类型_B627]">
                  <xsl:attribute name="errorStyle">
                    <xsl:value-of select="规则:错误提示_B626/@类型_B627"/>
                  </xsl:attribute>
                </xsl:if>
                <!--errorTitle-->
                <xsl:if test="规则:错误提示_B626[@标题_B624]">
                  <xsl:attribute name="errorTitle">
                    <xsl:value-of select="规则:错误提示_B626/@标题_B624"/>
                  </xsl:attribute>
                </xsl:if>
                <!--error-->
                <xsl:if test="规则:错误提示_B626[@内容_B625]">
                  <xsl:attribute name="error">
                    <xsl:value-of select="规则:错误提示_B626/@内容_B625"/>
                  </xsl:attribute>
                </xsl:if>
                <!--showErrorMessage-->
                <xsl:if test="规则:错误提示_B626[@是否显示_B623] and 规则:错误提示_B626[@是否显示_B623='true']">
                  <xsl:attribute name="showErrorMessage">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <!--showDropDown-->
              <xsl:if test="规则:是否显示下拉箭头_B621">
                <xsl:if test="规则:是否显示下拉箭头_B621='true'">
                  <xsl:attribute name="showDropDown">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="规则:是否显示下拉箭头_B621='false'">
                  <xsl:attribute name="showDropDown">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <xsl:if test="规则:输入提示_B622">
                <!--showInputMessage-->
                <xsl:if test="规则:输入提示_B622[@是否显示_B623] and 规则:输入提示_B622[@是否显示_B623 = 'true']">
                  <xsl:attribute name="showInputMessage">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--promptTitle-->
                <xsl:if test="规则:输入提示_B622[@标题_B624]">
                  <xsl:attribute name="promptTitle">
                    <xsl:value-of select="规则:输入提示_B622/@标题_B624"/>
                  </xsl:attribute>
                </xsl:if>
                <!--prompt-->
                <xsl:if test="规则:输入提示_B622[@内容_B625]">
                  <xsl:attribute name="prompt">
                    <xsl:value-of select="规则:输入提示_B622/@内容_B625"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <xsl:if test="规则:操作码_B61D">
                <xsl:attribute name="operator">
                  <xsl:call-template name="opera"/>
                </xsl:attribute>
              </xsl:if>

              <!--Modified by LDM in 2010/11/29-->
              <!--UOF1.1中增加了区域集元素，需组合各区域值：区域引用 和 单个单元格引用-->
              <xsl:if test="规则:区域集_B61A">
                <xsl:variable name="count" select="count(./规则:区域集_B61A/规则:区域_B61B)"/>
                
                <!--2014-3-18, update by Qihy, 增加数据有效性-选中R1C1引用样式时的转换， start-->
                <!--<xsl:variable name="sqref_var">
                  <xsl:call-template name="combineRef">
                    <xsl:with-param name="count" select="$count"/>
                    <xsl:with-param name="seq" select="'1'"/>
                  </xsl:call-template>
                </xsl:variable>-->
                <xsl:variable name="sqref_var">
                  <xsl:choose>
                    
                    <!--2014-3-24, update by Qihy, 修复bug_3158	(Strict)OOX->UOF->OOX 日期数据丢失,打开需要修复， start-->
                    <xsl:when test ="./ancestor::uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否RC引用_B634 = 'true'">
                    <!--2014-3-24 end-->
                      
                      <xsl:call-template name ="combineRef1">
                        <xsl:with-param name="count" select="$count"/>
                        <xsl:with-param name="seq" select="'1'"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="combineRef">
                        <xsl:with-param name="count" select="$count"/>
                        <xsl:with-param name="seq" select="'1'"/>
                      </xsl:call-template>
                    </xsl:otherwise>
                  </xsl:choose>
                  
                </xsl:variable>
                <!--2014-3-18 end-->
                
                <xsl:attribute name="sqref">
                  <xsl:value-of select="$sqref_var"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(./规则:校验类型_B61C = 'list')">
                <xsl:if test="规则:第一操作数_B61E">
                  <formula1>
                    <xsl:variable name="formula1">
                      <xsl:value-of select="规则:第一操作数_B61E"/>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="contains($formula1,'=')">
                        <xsl:value-of select="substring-after($formula1,'=')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$formula1"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </formula1>
                </xsl:if>
              </xsl:if>
              <xsl:if test="./规则:校验类型_B61C = 'list'">
                <xsl:if test="规则:第一操作数_B61E">
                  <formula1>
                    <xsl:variable name="refArea">
                      <xsl:value-of select="规则:第一操作数_B61E"/>
                      <!--./parent::表:数据有效性/表:区域集/表:区域-->
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="contains($refArea,'=')">
                        <xsl:variable name="formula1">
                          <xsl:value-of select="substring-after($refArea,'=')"/>
                        </xsl:variable>
                        <xsl:value-of select="$formula1"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="concat('&quot;',$refArea,'&quot;')"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </formula1>
                </xsl:if>
              </xsl:if>
              <xsl:if test="规则:第二操作数_B61F and not(./规则:校验类型_B61C = 'list')">
                <formula2>
                  <xsl:variable name="formula2">
                    <xsl:value-of select="规则:第二操作数_B61F"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="contains($formula2,'=')">
                      <xsl:value-of select="substring-after($formula2,'=')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$formula2"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </formula2>
              </xsl:if>
            </dataValidation>
          </xsl:for-each>
        </dataValidations>
      </xsl:if>
      <!--Hyperlink-->
      <xsl:if test=".//表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2[@超链接引用_E7BE]">
        <hyperlinks>
          <xsl:for-each select=".//表:工作表内容_E80E/表:行_E7F1/表:单元格_E7F2[@超链接引用_E7BE]">
            <xsl:call-template name="HyperLink"/>
          </xsl:for-each>
        </hyperlinks>
      </xsl:if>
      <xsl:if test=".//表:工作表属性_E80D">
        <xsl:if test=".//表:工作表属性_E80D[表:页面设置_E7C1]">
          <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:打印_E7CA]">
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否带网格线_E7CB or @是否带行号列标_E7CC]">
              <printOptions>
                <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:垂直对齐方式_E701='center'">
                  <xsl:attribute name="verticalCentered">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:水平对齐方式_E700='center'">
                  <xsl:attribute name="horizontalCentered">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                
               <!--2014-4-28, delete by Qihy, 页面设置-页边距-对齐方式不正确， start-->
                <!--<xsl:attribute name="verticalCentered">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>-->
                <!--2014-4-28 end-->
                
                <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否带行号列标_E7CC='true']">
                  <xsl:attribute name="headings">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否带网格线_E7CB='true']">
                  <xsl:attribute name="gridLines">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否带网格线_E7CB='false']">
                  <xsl:attribute name="gridLines">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                </xsl:if>
              </printOptions>
            </xsl:if>
          </xsl:if>
          <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:页边距_E7C5]">
            <pageMargins>
              <xsl:for-each select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:页边距_E7C5">
                <!--pageMargins-->
                <xsl:if test="@左_C608">
                  <xsl:variable name="l" select="@左_C608"/>
                  <xsl:attribute name="left">
                    <xsl:value-of select="number($l * 0.0138)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="@右_C60A">
                  <xsl:variable name="r" select="@右_C60A"/>
                  <xsl:attribute name="right">
                    <xsl:value-of select="number($r * 0.0138)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="@上_C609">
                  <xsl:variable name="t" select="@上_C609"/>
                  <xsl:attribute name="top">
                    <xsl:value-of select="number($t * 0.0138)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="@下_C60B">
                  <xsl:variable name="b" select="@下_C60B"/>
                  <xsl:attribute name="bottom">
                    <xsl:value-of select="number($b * 0.0138)"/>
                  </xsl:attribute>
                </xsl:if>
                <!--UOF1.0中的uof:页边距类型/@页眉、@页脚已经在UOF1.1中删除-->
                <!--UOF1.1中的页面设置类型中有一个单独的页眉页脚元素-->
                <!--UOF1.1中无法设置页眉页脚位置，采用默认设置，分别默认为1.27cm,1.27cm-->
                <!--UOF2.0中无法设置页眉页脚位置，schema中无定义，分别默认为1.3,1.3。李杨2012-3-16-->
                <xsl:attribute name="header">
                  <xsl:value-of select ="'0.5'"/>
                  <!--<xsl:value-of select="number(1.27 * 0.0138)"/>-->
                </xsl:attribute>
                <xsl:attribute name="footer">
                  <xsl:value-of select ="'0.5'"/>
                  <!--<xsl:value-of select="number(1.27 * 0.0138)"/>-->
                </xsl:attribute>
              </xsl:for-each>
            </pageMargins>
          </xsl:if>
          <pageSetup>
            <xsl:variable name="pageWidth">
              <xsl:value-of select="floor(.//表:工作表属性_E80D/表:页面设置_E7C1/表:纸张_E7C2/@宽_C605)"/>
            </xsl:variable>
            <xsl:variable name="pageHeight">
              <xsl:value-of select="floor(.//表:工作表属性_E80D/表:页面设置_E7C1/表:纸张_E7C2/@长_C604)"/>
            </xsl:variable>
            <xsl:attribute name="paperSize">
              <xsl:choose>
                <xsl:when test="$pageWidth = '728' and $pageHeight='1031'">
                  <xsl:value-of select="'12'"/>
                </xsl:when>
                <xsl:when test="$pageWidth = '515' and $pageHeight='728'">
                  <xsl:value-of select="'13'"/>
                </xsl:when>
                <!--A3-->
                <xsl:when test="$pageWidth = '842' and $pageHeight='1190'">
                  <xsl:value-of select="'8'"/>
                </xsl:when>
                <!--A4-->
                <xsl:when test="$pageWidth = '595' and $pageHeight='842'">
                  <xsl:value-of select="'9'"/>
                </xsl:when>
                <!--A5-->
                <xsl:when test="$pageWidth = '419' and $pageHeight='595'">
                  <xsl:value-of select="'11'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'9'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <!--
            <xsl:if test=".//表:工作表属性/表:页面设置[表:纸张 and 表:纸张/@uof:纸型 &gt; 0]">
              <xsl:variable name="ps" select=".//表:工作表属性/表:页面设置/表:纸张/@uof:纸型"/>
              <xsl:attribute name="paperSize">
                <xsl:value-of select="$ps"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性/表:页面设置[表:纸张 and 表:纸张/@uof:纸型 &lt; 0]">
              <xsl:variable name="ps" select=".//表:工作表属性/表:页面设置/表:纸张/@uof:纸型"/>
              <xsl:attribute name="paperSize">
                <xsl:value-of select="'9'"/>
              </xsl:attribute>
            </xsl:if>
            -->
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:缩放_E7C4]">
              <xsl:attribute name="scale">
                <xsl:value-of select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:缩放_E7C4"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:调整_E7D1/@页宽倍数_E7D3]">
              <xsl:attribute name="fitToWidth">
                <xsl:value-of select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:调整_E7D1/@页宽倍数_E7D3"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:调整_E7D1/@页高倍数_E7D2]">
              <xsl:attribute name="fitToHeight">
                <xsl:value-of select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:调整_E7D1/@页高倍数_E7D2"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:打印_E7CA/@是否先列后行_E7CE]">
              <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否先列后行_E7CE='false']">
                <xsl:attribute name="pageOrder">
                  <xsl:value-of select="'overThenDown'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA[@是否先列后行_E7CE='true']">
                <xsl:attribute name="pageOrder">
                  <xsl:value-of select="'downThenOver'"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:纸张方向_E7C3]">
              <xsl:attribute name="orientation">
                <xsl:variable name="fx" select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:纸张方向_E7C3"/>
                <xsl:if test="$fx='landscape'">
                  <xsl:value-of select="'landscape'"/>
                </xsl:if>
                <xsl:if test="$fx='portrait'">
                  <xsl:value-of select="'portrait'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:打印_E7CA/@是否按草稿方式_E7CD]">
              <xsl:attribute name="draft">
                <xsl:value-of select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:打印_E7CA/@是否按草稿方式_E7CD"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1/表:批注打印方式_E7CF='sheet-end'">
              <xsl:attribute name="cellComments">
                <xsl:value-of select="'atEnd'"/>
              </xsl:attribute>
            </xsl:if>
            <!--下面的内容必须有才能显示出图表-->
            <xsl:if test=".//表:工作表属性_E80D/表:页面设置_E7C1[表:错误单元格打印方式_E7D0]">
              <xsl:variable name="sty" select=".//表:工作表属性_E80D/表:页面设置_E7C1/表:错误单元格打印方式_E7D0"/>
              <xsl:attribute name="errors">
                <xsl:choose>
                  <xsl:when test="$sty='blank'">
                    <xsl:value-of select="'blank'"/>
                  </xsl:when>
                  <xsl:when test="$sty='dash'">
                    <xsl:value-of select="'dash'"/>
                  </xsl:when>
                  <xsl:when test="$sty='na'">
                    <xsl:value-of select="'NA'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'displayed'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </pageSetup>
        </xsl:if>
      </xsl:if>
      <!--Not Checked-->
      <!--Marked by LDM in 2010/12/19-->
      <xsl:if test="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6">
        <headerFooter>
          <!--页眉-->
          <xsl:if test="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='header-left' or @位置_E7C9='header-center' or @位置_E7C9='header-right']">
            <oddHeader>
              <!--左-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='header-left']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:value-of select ="'&amp;L'"/>
                <xsl:for-each select="*">
                  <!--2015-03-20 update by sunxinwei start -->
                  <!--<xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:value-of select ="'&amp;L'"/>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:if test ="字:文本串_415B='sheetName'">
                      <xsl:value-of select ="'&amp;A'"/>
                    </xsl:if>
                    <xsl:if test="字:文本串_415B='numpages'">
                      <xsl:value-of select="'&amp;N'"/>
                    </xsl:if>
                    <xsl:if test="字:文本串_415B='page'">
                      <xsl:value-of select="'&amp;P'"/>
                    </xsl:if>
                    <xsl:if test="字:文本串_415B='date'">
                      <xsl:value-of select="'&amp;D'"/>
                    </xsl:if>
                    <xsl:if test="字:文本串_415B='filename'">
                      <xsl:value-of select="'&amp;F'"/>
                    </xsl:if>
                    <xsl:if test="字:文本串_415B='time'">
                      <xsl:value-of select="'&amp;T'"/>
                    </xsl:if>
                    <xsl:if test ="字:文本串_415B='Filename'">
                      <xsl:value-of select ="'&amp;Z'"/>
                    </xsl:if>
                  </xsl:if> -->

                  <!--<xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="字:文本串_415B"/>                  
                  </xsl:if>
               </xsl:for-each>
                  <xsl:if test="name()='字:空格符_4161'">&#160;</xsl:if>-->
                  <xsl:variable name="index" select="position()"></xsl:variable>
                  <xsl:if test="name()='字:句_419D'">
                    <xsl:for-each select="*">
                      <xsl:if test="name()='字:空格符_4161'">&#160;</xsl:if>
                      <xsl:if test="name()='字:文本串_415B'">
                        <xsl:if test="not(name(../../node()[$index+1])='字:域结束_41A0')">
                          <xsl:value-of select="."/>
                        </xsl:if>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test="name()='字:域开始_419E'">
                    <xsl:variable name="version" select="@类型_416E"/>
                    <xsl:variable name ="val">
                      <xsl:call-template name ="headerandfoot">
                        <xsl:with-param name ="leixing" select ="$version"/>
                      </xsl:call-template>
                    </xsl:variable>
                    <xsl:value-of select ="concat('&amp;',$val)"/>
                  </xsl:if>
                  <!--  </xsl:for-each>
               <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="$val"/>-->
                </xsl:for-each>
              </xsl:for-each>
              <!--中-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='header-center']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:choose >
                      <xsl:when test ="contains(字:文本串_415B,'sheetName')">
                        <xsl:value-of select ="'&amp;C&amp;A'"/>
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:value-of select="concat('&amp;C',字:文本串_415B)"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="concat('&amp;C',字:文本串_415B)"/>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="concat('&amp;C',$val)"/>
                </xsl:for-each>
              </xsl:for-each>
              <!--右-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='header-right']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:choose>
                      <xsl:when test ="contains(字:文本串_415B,'sheetName')">
                        <xsl:value-of select ="'&amp;R&amp;A'"/>
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:value-of select="concat('&amp;R',字:文本串_415B)"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="concat('&amp;R',字:文本串_415B)"/>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="concat('&amp;R',$val)"/>
                </xsl:for-each>
              </xsl:for-each>
            </oddHeader>
          </xsl:if>
          <!--页脚-->
          <xsl:if test="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='footer-left' or @位置_E7C9='footer-center' or @位置_E7C9='footer-right']">
            <oddFooter>
              <!--左-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='footer-left']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:choose>
                      <xsl:when test ="contains(字:文本串_415B,'sheetName')">
                        <xsl:value-of select ="'&amp;L&amp;A'"/>
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:value-of select="concat('&amp;L',字:文本串_415B)"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:if>
                  <xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="concat('&amp;L',字:文本串_415B)"/>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="concat('&amp;L',$val)"/>

                </xsl:for-each>
              </xsl:for-each>
              <!--中-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='footer-center']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:choose>
                      <xsl:when test ="contains(字:文本串_415B,'sheetName')">
                        <xsl:value-of select ="'&amp;C&amp;A'"/>
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:value-of select="concat('&amp;C',字:文本串_415B)"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:if>
                  <xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="concat('&amp;C',字:文本串_415B)"/>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="concat('&amp;C',$val)"/>
                </xsl:for-each>
              </xsl:for-each>
              <!--右-->
              <xsl:for-each select="./表:工作表_E825/表:工作表属性_E80D/表:页面设置_E7C1/表:页眉页脚_E7C6[@位置_E7C9='footer-right']/字:段落_416B">
                <xsl:if test="字:句_419D/字:句属性_4158/字:字体_4128[@颜色_412F]">
                  <xsl:variable name="ziticolor" select="translate(字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F,'#','')"/>
                  <xsl:variable name="ziticolor1" select="concat('K',$ziticolor)"/>
                  <xsl:variable name="zichuan" select="字:句_419D/字:文本串_415B"/>
                  <xsl:value-of select="concat('&amp;',$ziticolor1,$zichuan)"/>
                </xsl:if>
                <xsl:for-each select="字:句_419D">
                  <xsl:if test="字:句属性_4158 and not(preceding-sibling::字:域代码_419F)">
                    <xsl:choose>
                      <xsl:when test ="contains(字:文本串_415B,'sheetName')">
                        <xsl:value-of select ="'&amp;R&amp;A'"/>
                      </xsl:when>
                      <xsl:otherwise >
                        <xsl:value-of select="concat('&amp;R',字:文本串_415B)"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:if>
                  <xsl:if test="not(字:句属性_4158)">
                    <xsl:value-of select="concat('&amp;R',字:文本串_415B)"/>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="字:域开始_419E">
                  <xsl:variable name="version" select="@类型_416E"/>
                  <xsl:variable name ="val">
                    <xsl:call-template name ="headerandfoot">
                      <xsl:with-param name ="leixing" select ="$version"/>
                    </xsl:call-template>
                  </xsl:variable>
                  <xsl:value-of select ="concat('&amp;R',$val)"/>
                </xsl:for-each>
              </xsl:for-each>
            </oddFooter>
          </xsl:if>
        </headerFooter>
        <!-- sunxinwei 20150331 end-->
      </xsl:if>
      <!--update by sunxinwei 20150408 strat-->
      <xsl:if test="./表:工作表_E825/表:分页符集_E81E">
        <rowBreaks>
          <xsl:variable name="cnt" select="count(./表:工作表_E825/表:分页符集_E81E/表:分页符_E81F[@行号_E820])"/>
          <xsl:attribute name="count">
            <xsl:value-of select="$cnt"/>
          </xsl:attribute>
          <xsl:attribute name="manualBreakCount">
            <xsl:value-of select="$cnt"/>
          </xsl:attribute>
          <xsl:for-each select="./表:工作表_E825/表:分页符集_E81E/表:分页符_E81F">
            <xsl:if test="@行号_E820">
              <xsl:variable name="rid" select="@行号_E820"/>
              <brk>
                <xsl:attribute name="id">
                  <xsl:value-of select="$rid"/>
                </xsl:attribute>
                <xsl:attribute name="max">
                  <xsl:value-of select="'16383'"/>
                </xsl:attribute>
                <xsl:attribute name="man">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </brk>
            </xsl:if>
          </xsl:for-each>
        </rowBreaks>
        <colBreaks>
          <xsl:variable name="cnt2" select="count(./表:工作表_E825/表:分页符集_E81E/表:分页符_E81F[@列号_E821])"/>
          <xsl:attribute name="count">
            <xsl:value-of select="$cnt2"/>
          </xsl:attribute>
          <xsl:attribute name="manualBreakCount">
            <xsl:value-of select="$cnt2"/>
          </xsl:attribute>
          <xsl:for-each select="./表:工作表_E825/分页符集_E81E/表:分页符_E81F">
            <xsl:if test="@列号_E821">
              <xsl:variable name="rid" select="@列号_E821"/>
              <brk>
                <xsl:attribute name="id">
                  <xsl:value-of select="$rid"/>
                </xsl:attribute>
                <xsl:attribute name="max">
                  <xsl:value-of select="'1048575'"/>
                </xsl:attribute>
                <xsl:attribute name="man">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </brk>
            </xsl:if>
          </xsl:for-each>
        </colBreaks>
        <!--sunxinwei 20150408 end-->
      </xsl:if>
      <!--表：图表 修改  李杨2012-2-21-->
      <xsl:if test="./表:图表集合/表:单图表/图表:图表_E837 or .//表:工作表内容_E80E/uof:锚点_C644">
        <drawing r:id="rId1"/>
      </xsl:if>

      <xsl:if test=".//表:工作表属性_E80D/表:背景填充_E830/图:图片_8005">
        <picture r:id="rId100"/>
      </xsl:if>
      
      <xsl:if test =".//表:工作表内容_E80E//表:批注_E7B7 and not(./表:图表集合/表:单图表/图表:图表_E837 or .//表:工作表内容_E80E/uof:锚点_C644)">
        <drawing r:id="rId_Comment1"/>
      </xsl:if>
    </worksheet>
  </xsl:template>
  <xsl:template name ="headerandfoot">
    <xsl:param name ="leixing"/>
    <xsl:choose>
      <xsl:when test="$leixing='numpages'">
        <xsl:value-of select="'&amp;N'"/>
      </xsl:when>
      <xsl:when test="$leixing='page'">
        <xsl:value-of select="'&amp;P'"/>
      </xsl:when>
      <xsl:when test="$leixing='date'">
        <xsl:value-of select="'&amp;D'"/>
      </xsl:when>
      <xsl:when test="$leixing='filename'">
        <xsl:value-of select="'&amp;F'"/>
      </xsl:when>
      <xsl:when test="$leixing='time'">
        <xsl:value-of select="'&amp;T'"/>
      </xsl:when>
      <xsl:when test ="$leixing='Filename'">
        <xsl:value-of select ="'&amp;Z'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template match ="规则:条件格式化_B629">
    <xsl:variable name ="n2">
      <xsl:value-of select ="count(./规则:条件_B62B)"/>
    </xsl:variable>
    <xsl:value-of select ="$n2"/>
  </xsl:template>

  <!--添加列转换。李杨2012-3-18-->
  <xsl:template name ="ColumnTran">
    <xsl:param name ="Column"/>
    <xsl:variable name ="c" select ="$Column"/>
    <xsl:choose >
      <xsl:when test ="$c &lt; 10">
        <xsl:variable name ="cv">
          <xsl:value-of select ="translate($c,'123456789','ABCDEFGHI')"/>
        </xsl:variable>
        <xsl:value-of select ="$cv"/>
      </xsl:when>
      <xsl:when test ="($c &lt; 21)and($c &gt; 9)">
        <xsl:if test ="$c = 10">
          <xsl:value-of select ="'J'"/>
        </xsl:if>
        <xsl:if test ="$c = 11">
          <xsl:value-of select ="'K'"/>
        </xsl:if>
        <xsl:if test ="$c = 12">
          <xsl:value-of select ="'L'"/>
        </xsl:if>
        <xsl:if test ="$c = 13">
          <xsl:value-of select ="'M'"/>
        </xsl:if>
        <xsl:if test ="$c = 14">
          <xsl:value-of select ="'N'"/>
        </xsl:if>
        <xsl:if test ="$c = 15">
          <xsl:value-of select ="'O'"/>
        </xsl:if>
        <xsl:if test ="$c = 16">
          <xsl:value-of select ="'P'"/>
        </xsl:if>
        <xsl:if test ="$c = 17">
          <xsl:value-of select ="'Q'"/>
        </xsl:if>]
        <xsl:if test ="$c = 18">
          <xsl:value-of select ="'R'"/>
        </xsl:if>
        <xsl:if test ="$c = 19">
          <xsl:value-of select ="'S'"/>
        </xsl:if>
        <xsl:if test ="$c = 20">
          <xsl:value-of select ="'T'"/>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <!--添加公式模板。李杨2012-3-19-->
  <xsl:template name ="formula1">
    <xsl:param name ="vbe"/>
    <xsl:variable name ="be2" select ="substring-after($vbe,'R')"/>
    <xsl:variable name ="beR" select ="4 + substring-before($be2,'C')"/>
    <xsl:variable name ="beC2" select ="3 + substring-after($be2,'C')"/>
    <xsl:variable name ="beC">
      <xsl:call-template name ="ColumnTran">
        <xsl:with-param name ="Column" select ="$beC2"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select ="concat($beC,$beR)"/>
  </xsl:template>
  <!--Modified by LDM in 2010/12/19-->
  <!--单元格数据模板-->
  <xsl:template match="表:数据_E7B3" name ="CellFormat">
    <xsl:variable name="sheetNo">
      <xsl:value-of select="ancestor::表:单工作表/@uof:sheetNo"/>
    </xsl:variable>
    <xsl:variable name="sheetId">
      <xsl:value-of select ="ancestor::表:工作表_E825/@标识符_E7AC"/>
    </xsl:variable>
    <!--数据类型属性-->
    <xsl:attribute name="t">
      <xsl:choose>
        <!--boolean?-->
        <xsl:when test="./@类型_E7B6='boolean'">
          <xsl:value-of select="'b'"/>
        </xsl:when>
        <!--error?-->
        <xsl:when test="./@类型_E7B6='error'">
          <xsl:value-of select="'e'"/>
        </xsl:when>
        <xsl:when test="./@类型_E7B6='date'">
          <xsl:value-of select="'d'"/>
        </xsl:when>
        
        <!--zl 20150512-->
        <xsl:when test="./@类型_E7B6='number' and not(contains(./表:公式_E7B5,'DATE')) and not(contains(./表:公式_E7B5,'TODAY'))">
          <xsl:value-of select="'n'"/>
        </xsl:when>
        
        <xsl:when test="./@类型_E7B6='number' and contains(./表:公式_E7B5,'DATE')">
          <xsl:value-of select="'d'"/>
        </xsl:when>

        <xsl:when test="./@类型_E7B6='number' and contains(./表:公式_E7B5,'TODAY')">
          <xsl:value-of select="'d'"/>
        </xsl:when>
		
		    <xsl:when test="./@类型_E7B6='date'">
          <xsl:value-of select="'d'"/>
        </xsl:when>
        <!--zl 20150512-->
        
        <xsl:when test="./@类型_E7B6='text'">
          <xsl:value-of select="'s'"/>
          <!--公式？-->
          <!--
                    <xsl:choose>
                      <xsl:when test="./表:公式">
                        <xsl:value-of select="'str'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'s'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                    -->
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'s'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <!--单元格引用式样属性-->
    <!--Modified by LDM in 2011/01/07-->
    <!--zl 2015-4-21-->
    <xsl:if test="ancestor::表:单元格_E7F2/表:数据_E7B3/@类型_E7B6 != 'number'">
      <xsl:attribute name="s">
        <xsl:variable name="styleRef" select="ancestor::表:单元格_E7F2/@式样引用_E7BD"/>
        <xsl:variable name="position">
          <xsl:for-each select="/uof:UOF/式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[@标识符_E7AC = $styleRef]">
            <xsl:value-of select="count(./preceding-sibling::式样:单元格式样_9916)"/>
          </xsl:for-each>
        </xsl:variable>

        <xsl:value-of select="$position"/>
      </xsl:attribute>
    </xsl:if>
    <!--zl 2015-4-21-->
    
    <!--zl 2015-5-17-->
    <xsl:if test="ancestor::表:单元格_E7F2/表:数据_E7B3/@类型_E7B6 = 'number' and not(ancestor::表:单元格_E7F2/表:数据_E7B3/表:公式_E7B5)">
      <xsl:attribute name="s">
        <xsl:value-of select="substring-after(ancestor::表:单元格_E7F2/@式样引用_E7BD,'_')"/>
      </xsl:attribute>
    </xsl:if>
    <!--zl 2015-5-17-->

    <!--zl 2015-5-17-->
    <xsl:if test="ancestor::表:单元格_E7F2/表:数据_E7B3/@类型_E7B6 = 'number' and contains(ancestor::表:单元格_E7F2/表:数据_E7B3/表:公式_E7B5,'DATE') or contains(ancestor::表:单元格_E7F2/表:数据_E7B3/表:公式_E7B5,'TODAY')">
      <xsl:attribute name="s">
        <xsl:value-of select="substring-after(ancestor::表:单元格_E7F2/@式样引用_E7BD,'_')"/>
      </xsl:attribute>
    </xsl:if>
    <!--zl 2015-5-17-->

    <!--zl 2015-5-17-->
    <xsl:if test="ancestor::表:单元格_E7F2/表:数据_E7B3/@类型_E7B6 = 'number' and contains(ancestor::表:单元格_E7F2/表:数据_E7B3/表:公式_E7B5,'SUM') or contains(ancestor::表:单元格_E7F2/表:数据_E7B3/表:公式_E7B5,'COUNTIF')">
      <xsl:attribute name="s">
        <xsl:value-of select="substring-after(ancestor::表:单元格_E7F2/@式样引用_E7BD,'_')"/>
      </xsl:attribute>
    </xsl:if>
    <!--zl 2015-5-17-->
    
    <!--
      <xsl:choose>
        <xsl:when test="$styleRef = 'DEFSTYLE' or $styleRef = ''">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="translate($styleRef,'CELLSTYLE_','')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
     -->
    <!--添加公式转换。李杨2012-3-19-->
    <xsl:if test="./表:公式_E7B5 and ./表:公式_E7B5!='' and ./@类型_E7B6!='str'">
      <f>
        <!--<xsl:variable name="funcTemp" select="./表:公式_E7B5"/>
        <xsl:value-of select="substring-after($funcTemp,'=')"/>-->

        <xsl:variable name ="v" select ="substring-after(./表:公式_E7B5,'=')"/>
        <xsl:variable name ="ss" select ="substring-before($v,'(')"/>
        <xsl:variable name ="ss2" select ="substring-after($v,'(')"/>
        <xsl:variable name ="kuohao" select ="translate($ss2,')','')"/>
        <xsl:choose >
          <xsl:when test ="contains($v,'[')">
            
            <!--公式中单元格引用以冒号（：）分割时-->
            <xsl:if test ="contains($kuohao,':')">
              <xsl:variable name ="be" select ="substring-before($kuohao,':')"/>
              <xsl:variable name ="af" select ="substring-after($kuohao,':')"/>
              <!--R[-3]C[-2]类型-->
              <xsl:if test ="contains(substring-before($be,'C'),']') and contains(substring-before($af,'C'),']')">
                <xsl:variable name ="fkh" select ="translate(translate($be,'[',''),']','')"/>
                <xsl:variable name ="before">
                  <xsl:call-template name ="formula1">
                    <xsl:with-param name ="vbe" select ="$fkh"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="fkh2" select ="translate(translate($af,'[',''),']','')"/>
                <xsl:variable name ="after">
                  <xsl:call-template name ="formula1">
                    <xsl:with-param name ="vbe" select ="$fkh2"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:value-of select ="concat($ss,'(',$before,':',$after,')')"/>
              </xsl:if>
              <!--<xsl:if test ="contains(substring-before($af,'C'),']')">
            <xsl:variable name ="fkh" select ="translate(translate($af,'[',''),']','')"/>R-3C-2
            <xsl:variable name ="after">
              <xsl:call-template name ="formula1">
                <xsl:with-param name ="vbe" select ="$fkh"/>
              </xsl:call-template>
            </xsl:variable>
          </xsl:if>-->
              <!--RC[-2]类型-->
              <xsl:if test ="not(contains(substring-before($be,'C'),']')) and not(contains(substring-before($af,'C'),']'))">
                <xsl:variable name ="fkh" select ="translate(translate($be,'[',''),']','')"/>
                <xsl:variable name ="vb" select ="3 + substring-after($fkh,'C')"/>
                <xsl:variable name ="col">
                  <xsl:call-template name="ColIndex">
                    <xsl:with-param name="colSeq" select="$vb"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="row" select ="../../@行号_E7F3"/>
                <xsl:variable name ="before" select ="concat($col,$row)"/>

                <xsl:variable name ="fkh2" select ="translate(translate($af,'[',''),']','')"/>
                <xsl:variable name ="va" select ="3 + substring-after($fkh2,'C')"/>
                <xsl:variable name ="col2">
                  <xsl:call-template name="ColIndex">
                    <xsl:with-param name="colSeq" select="$va"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="after" select ="concat($col2,$row)"/>
                <xsl:value-of select ="concat($ss,'(',$before,':',$after,')')"/>
              </xsl:if>
              <!--<xsl:if test ="not(contains(substring-before($af,'C'),']'))">
            <xsl:variable name ="fkh" select ="translate(translate($af,'[',''),']','')"/>
            <xsl:variable name ="va" select ="substring-after($fkh,'C')"/>
            <xsl:variable name ="col">
              <xsl:call-template name="ColIndex">
                <xsl:with-param name="colSeq" select="$va"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="row" select ="../../@行号_E7F3"/>
            <xsl:variable name ="after" select ="concat($col,$row)"/>
          </xsl:if>-->
            </xsl:if>
            <!--公式中单元格引用以逗号（，）分割时-->
            <xsl:if test ="contains($kuohao,',')">
              <xsl:variable name ="be" select ="substring-before($kuohao,',')"/>
              <xsl:variable name ="af" select ="substring-after($kuohao,',')"/>
              <!--R[-3]C[-2]类型-->
              <xsl:if test ="contains(substring-before($be,'C'),']') and contains(substring-before($af,'C'),']')">
                <xsl:variable name ="fkh" select ="translate(translate($be,'[',''),']','')"/>
                <xsl:variable name ="before">
                  <xsl:call-template name ="formula1">
                    <xsl:with-param name ="vbe" select ="$fkh"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="fkh2" select ="translate(translate($af,'[',''),']','')"/>
                <xsl:variable name ="after">
                  <xsl:call-template name ="formula1">
                    <xsl:with-param name ="vbe" select ="$fkh2"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:value-of select ="concat($ss,'(',$before,':',$after,')')"/>
              </xsl:if>
              <!--RC[-2]类型-->
              <xsl:if test ="not(contains(substring-before($be,'C'),']')) and not(contains(substring-before($af,'C'),']'))">
                <xsl:variable name ="fkh" select ="translate(translate($be,'[',''),']','')"/>
                <xsl:variable name ="vb" select ="3 + substring-after($fkh,'C')"/>
                <xsl:variable name ="col">
                  <xsl:call-template name="ColIndex">
                    <xsl:with-param name="colSeq" select="$vb"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="row" select ="../../@行号_E7F3"/>
                <xsl:variable name ="before" select ="concat($col,$row)"/>

                <xsl:variable name ="fkh2" select ="translate(translate($af,'[',''),']','')"/>
                <xsl:variable name ="va" select ="3 + substring-after($fkh2,'C')"/>
                <xsl:variable name ="col2">
                  <xsl:call-template name="ColIndex">
                    <xsl:with-param name="colSeq" select="$va"/>
                  </xsl:call-template>
                </xsl:variable>
                <xsl:variable name ="after" select ="concat($col2,$row)"/>
                <xsl:value-of select ="concat($ss,'(',$before,':',$after,')')"/>
              </xsl:if>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="$v"/>
          </xsl:otherwise>
        </xsl:choose>
      </f>
    </xsl:if>
    <xsl:if test="./@类型_E7B6='text'">
      <v>
        <xsl:variable name="rowID" select="ancestor::表:行_E7F1/@行号_E7F3"/>
        <xsl:variable name="colID" select="ancestor::表:单元格_E7F2/@列号_E7BC"/>
        <xsl:variable name="cellID" select="concat($rowID,'_',$colID)"/>
        <xsl:variable name="sharedStrSeq">
          <xsl:value-of select="ancestor::uof:UOF/uof:单元格内容集合/uof:单元格内容[@uof:sheetId=$sheetId and @uof:cellID=$cellID]/@uof:refID"/>
        </xsl:variable>
        <xsl:value-of select="$sharedStrSeq - 1"/>
      </v>
    </xsl:if>


	
    <xsl:if test="./@类型_E7B6='number'">
      <v>
        <xsl:value-of select=".//字:文本串_415B"/>
      </v>
    </xsl:if>
    <!--yanghaojie-start-->
    <xsl:if test="./@类型_E7B6='date'">
      <v>
        <xsl:value-of select=".//字:文本串_415B"/>
      </v>
    </xsl:if>
    <!--yanghaojie-end-->
    
    <xsl:if test ="./@类型_E7B6='error'">
      <v>
        <!--yanghaojie-start-->
        <xsl:variable name="temp" select=".//字:文本串_415B"/>
        <xsl:value-of select ="translate($temp,'*','#')"/>
        <!--<xsl:value-of select =".//字:文本串_415B"/>-->
        <!--yanghaojie-end-->
      </v>
    </xsl:if>

  </xsl:template>
  <!--Modified by LDM in 2010/12/19-->
  <!--单元格模板-->
  <xsl:template match="表:单元格_E7F2">
    
    <!--2014-4-8, update by Qihy, 修改批注颜色填充修复后打开， start-->

    <!--2014-4-27, update by Qihy, 修复加批注后单元格内容丢失问题， start-->
    <!--<xsl:if test="not(./表:批注_E7B7)">-->
    
    <!--2014-4-28, update by Qihy, 修复单元格样式丢失问题， start-->
    <xsl:if test="((./表:数据_E7B3 or ./表:合并_E7AF) and ./表:批注_E7B7) or not(./表:批注_E7B7)">
    <!--2014-4-28 end-->
      
    <!--2014-4-27 end-->  
      
    <c>
      <xsl:if test="@列号_E7BC">
        <xsl:variable name="colNo">
          <xsl:value-of select="@列号_E7BC"/>
        </xsl:variable>
        <!--李杨添加，数字列号转换字母列号。2012-2-10-->
        <xsl:variable name="colSeq">
          <xsl:call-template name="ColIndex">
            <xsl:with-param name="colSeq" select="$colNo"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="rowNo" select="ancestor::表:行_E7F1/@行号_E7F3"/>
        <xsl:attribute name="r">
          <xsl:value-of select="concat($colSeq,$rowNo)"/>
        </xsl:attribute>
        <!--yanghaojie-start-->
        <xsl:attribute name="s">
          <xsl:variable name="style" select="./@式样引用_E7BD"/>
          <xsl:variable name="style1" select="substring-after($style,'_')"/>
          <xsl:value-of select="$style1"/>
        </xsl:attribute>
        <!--yanghaojie-end-->
      </xsl:if>

      <!--单元格的数据类型-->
      <xsl:if test="./表:数据_E7B3">
        <xsl:apply-templates select="./表:数据_E7B3"/>
      </xsl:if>
      <xsl:if test="not(./表:数据_E7B3)">
        <xsl:attribute name="s">
          <xsl:variable name="styleRef" select="./@式样引用_E7BD"/>
          <xsl:variable name="position">
            <xsl:for-each select="/uof:UOF/式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[@标识符_E7AC = $styleRef]">
              <xsl:value-of select="count(./preceding-sibling::式样:单元格式样_9916)"/>
            </xsl:for-each>
          </xsl:variable>
          <xsl:value-of select="$position"/>
          <!--
          <xsl:variable name="styleRef" select="./@表:式样引用"/>
          <xsl:choose>
            <xsl:when test="$styleRef = 'DEFSTYLE' or $styleRef = ''">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="translate($styleRef,'CELLSTYLE_','')"/>
            </xsl:otherwise>
          </xsl:choose>
          -->
        </xsl:attribute>
      </xsl:if>
    </c>
    </xsl:if>
    <!--2014-4-8 end-->
  </xsl:template>


  <xsl:template name="HyperLink">
    <hyperlink>
      <xsl:variable name="colSeqTemp" select="./@列号_E7BC"/>
      <xsl:variable name="rowSeq" select="./parent::表:行_E7F1/@行号_E7F3"/>
      <xsl:variable name="colSeq">
        <xsl:call-template name="ColIndex">
          <xsl:with-param name="colSeq" select="$colSeqTemp"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:attribute name="ref">
        <xsl:value-of select="concat($colSeq,$rowSeq)"/>
      </xsl:attribute>
      <xsl:variable name="pos" select="translate(@超链接引用_E7BE,'ID','')"/>
      <xsl:variable name="refId" select="@超链接引用_E7BE"/>
      <xsl:if test="ancestor::表:电子表格文档_E826/preceding-sibling::超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$refId]">
        <xsl:variable name="tar" select="ancestor::表:电子表格文档_E826/preceding-sibling::超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$refId]/超链:目标_AA01"/>
        <xsl:choose>
          <xsl:when test="contains($tar,'!')">
            <xsl:attribute name="location">
              <xsl:value-of select="$tar"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="not(contains($tar,'!'))">
            <xsl:attribute name="r:id">
              <xsl:value-of select="concat('rId_HyperLink_',$pos)"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:if test="ancestor::表:电子表格文档_E826/preceding-sibling::超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$refId]/超链:提示_AA05">
          <xsl:attribute name="tooltip">
            <xsl:value-of select="ancestor::表:电子表格文档_E826/preceding-sibling::超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$refId]/超链:提示_AA05"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
    </hyperlink>
  </xsl:template>



  <xsl:template name="SheetRels">
    <xsl:param name="relsSheet"/>
    <xsl:variable name="sheetNo">
      <xsl:value-of select="./@uof:sheetNo"/>
    </xsl:variable>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--表：图表 修改  李杨2012-2-->
      <xsl:if test="./表:图表集合/表:单图表/图表:图表_E837 or .//表:工作表内容_E80E/uof:锚点_C644">
        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select="'rId1'"/>
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat('../drawings/drawing',$sheetNo,'.xml')"/>
          </xsl:attribute>
        </Relationship>
      </xsl:if>
      <!--hyperLinks-->
      <xsl:if test="//表:单元格_E7F2[@超链接引用_E7BE]">
        <xsl:for-each select="//表:单元格_E7F2[@超链接引用_E7BE]">
          <xsl:variable name="hyperId" select="@超链接引用_E7BE"/>
          <xsl:variable name="hyperId2" select="translate($hyperId,'ID','')"/>
          <xsl:if test="not(contains(ancestor::uof:UOF/超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$hyperId]/超链:目标_AA01,'!'))">
            <Relationship>
              <xsl:attribute name="Id">
                <xsl:value-of select="concat('rId_HyperLink_',$hyperId2)"/>
              </xsl:attribute>
              <xsl:attribute name="Type">
                <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink'"/>
              </xsl:attribute>
              <xsl:attribute name="Target">
                <xsl:variable name="goal">
                  <xsl:value-of select="ancestor::uof:UOF/超链:链接集_AA0B/超链:超级链接_AA0C[@标识符_AA0A=$hyperId]/超链:目标_AA01"/>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="contains($goal,':') and string-length(substring-before($goal,':'))=1">
                    <xsl:value-of select="concat('file:///',$goal)"/>
                  </xsl:when>
                  
                  <!--20130110 gaoyuwei bug_2663 "链接：电子邮件地址多了一个mailto"  start-->
					<xsl:when test="contains($goal,'mailto:mailto')">
						 <xsl:value-of select="substring-after( $goal, 'mailto:')"/>
					</xsl:when>
					<!--end-->
					
                  <xsl:otherwise>
                    <xsl:value-of select="$goal"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="TargetMode">
                <xsl:value-of select="'External'"/>
              </xsl:attribute>
              <!--uof中没有元素是表示内部链接还是外部链接，所以这里都转为外部-->
            </Relationship>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>

      <!--背景填充-->
      <!--背景图片填充  修改 李杨2012-2-21-->
      <!--Modified by LDM in 2011/01/23-->
      <xsl:if test=".//表:背景填充_E830/图:图片_8005">
        <xsl:variable name="imageId">
          <xsl:value-of select=".//表:背景填充_E830/图:图片_8005/@图形引用_8007"/>
        </xsl:variable>
        <!--<xsl:variable name="imageType">
          <xsl:value-of select="/uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$imageId]/@公共类型_D706"/>
        </xsl:variable>-->
        <xsl:variable name ="picName1">
          <xsl:value-of select ="/uof:UOF/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$imageId]/对象:路径_D703"/>
        </xsl:variable>
        <xsl:variable name ="picName2">
      
			<!-- 20130319 gaoyuwei bug_2741 背景图片填充的填充内容丢失 20130116 -->
			<xsl:choose>
				<xsl:when test="contains($picName1,'.\data\')">
					<xsl:value-of select="substring-after($picName1,'.\data\')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="substring-after($picName1,'/data/')"/>
				</xsl:otherwise>
			</xsl:choose>
			<!-- end -->      
      
        </xsl:variable>
        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select="'rId100'"/>
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/image'"/>
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select ="concat('../media/',$picName2)"/>
            <!--<xsl:choose>
              <xsl:when test="$imageType='jpg'">
                <xsl:value-of select="concat('../media/',$imageId,'.','jpeg')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('../media/',$imageId,'.',$imageType)"/>
              </xsl:otherwise>
            </xsl:choose>-->
          </xsl:attribute>
        </Relationship>
      </xsl:if>

      <xsl:if test=".//表:批注_E7B7">
        <Relationship>
          <xsl:attribute name="Id">
            <xsl:value-of select ="'rId_Comment1'"/>
            <!--<xsl:value-of select="concat('rId_Comment',$sheetNo)"/>-->
          </xsl:attribute>
          <xsl:attribute name="Type">
            <xsl:value-of select ="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
            <!--<xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/comments'"/>-->
          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select ="concat('../drawings/drawing',$sheetNo,'.xml')"/>
            <!--<xsl:value-of select="concat('../comments',$sheetNo,'.xml')"/>-->
          </xsl:attribute>
        </Relationship>
      </xsl:if>
    </Relationships>
  </xsl:template>
  <xsl:template name="sty">
    <xsl:param name="styid"/>
    <xsl:for-each select="ancestor::式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[@标识符_E7AC=$styid]">
      <xsl:value-of select="position()"/>
    </xsl:for-each>
  </xsl:template>
  <!--数据有效性-->
  <xsl:template name="opera">
    <xsl:choose>
      <xsl:when test="规则:操作码_B61D = 'between'">
        <xsl:value-of select="'between'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'equal-to'">
        <xsl:value-of select="'equal'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'greater-than'">
        <xsl:value-of select="'greaterThan'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'greater-than-or-equal-to'">
        <xsl:value-of select="'greaterThanOrEqual'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D='less-than'">
        <xsl:value-of select="'lessThan'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'less-than-or-equal-to'">
        <xsl:value-of select="'lessThanOrEqual'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'not-between'">
        <xsl:value-of select="'notBetween'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'not-equal-to'">
        <xsl:value-of select="'notEqual'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'start-with'">
        <xsl:value-of select="'beginsWith'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'contain'">
        <xsl:value-of select="'containsText'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B61D = 'end-with'">
        <xsl:value-of select="'endsWith'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--条件格式化-->
  <xsl:template name="opera2">
    <xsl:choose>
      <xsl:when test="规则:操作码_B62C = 'between'">
        <xsl:value-of select="'between'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C = 'equal-to'">
        <xsl:value-of select="'equal'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C = 'greater-than'">
        <xsl:value-of select="'greaterThan'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C = 'greater-than-or-equal-to'">
        <xsl:value-of select="'greaterThanOrEqual'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C='less-than'">
        <xsl:value-of select="'lessThan'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C = 'less-than-or-equal-to'">
        <xsl:value-of select="'lessThanOrEqual'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C = 'not-between'">
        <xsl:value-of select="'notBetween'"/>
      </xsl:when>
      <xsl:when test="规则:操作码_B62C='not-equal-to'">
        <xsl:value-of select="'notEqual'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--待处理 李杨2012-1-5-->
  <xsl:template match="uof:si的编号">
    <xsl:param name="cid"/>
    <xsl:param name="sheetid"/>
    <xsl:if test="uof:单元格里的句si[@uof:cellid=$cid]">
      <xsl:if test="uof:单元格里的句si[@uof:sheetid=$sheetid]">
        <xsl:variable name="sss" select="@uof:id"/>
        <xsl:value-of select="$sss - 1"/>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!--Modified by LDM in 2010/11/29-->
  <!--组合区域引用值-->
  <xsl:template name="combineRef">
    <xsl:param name="count"/>
    <xsl:param name="seq"/>
    <xsl:param name="sqref"/>
    <xsl:if test="$count='0'">
      <xsl:value-of select="$sqref"/>
    </xsl:if>
    <xsl:if test="not($count='0')">
      <xsl:variable name="subRef" select="规则:区域集_B61A/规则:区域_B61B[number($count)]"/>
      <!--区域引用 和 单个单元格引用-->
      <xsl:variable name="sqref_subVar_part">
        <xsl:value-of select="translate(substring-after($subRef,'!'),'$','')"/>
      </xsl:variable>
      <xsl:variable name="sqref_subVar">
        <xsl:choose>
          <xsl:when test="$seq=1">
            <xsl:value-of select="$sqref_subVar_part"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat($sqref,' ',$sqref_subVar_part)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:call-template name="combineRef">
        <xsl:with-param name="count" select="number($count) - 1"/>
        <xsl:with-param name="seq" select="number($seq) + 1"/>
        <xsl:with-param name="sqref" select="$sqref_subVar"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!--2014-3-18, add by Qihy, 增加数据有效性-选中R1C1引用样式时的转换， start-->
  <xsl:template name="combineRef1">
    <xsl:param name="count"/>
    <xsl:param name="seq"/>
    <xsl:param name="sqref"/>
    <xsl:if test="$count='0'">
      <xsl:value-of select="$sqref"/>
    </xsl:if>
    <xsl:if test="not($count='0')">
      <xsl:variable name="subRef" select="规则:区域集_B61A/规则:区域_B61B[number($count)]"/>
      <!--区域引用 和 单个单元格引用-->
      <xsl:choose>
        <xsl:when test="contains(substring-after($subRef,'!'),':')">
          <xsl:variable name ="subRef_before" select ="substring-before(substring-after($subRef,'!'), ':')"/>
          <xsl:variable name ="subRef_after" select ="substring-after(substring-after($subRef,'!'), ':')"/>
          <xsl:variable name ="colSeq1" select ="substring-after($subRef_before,'C')"/>
          <xsl:variable name="sqref_subVar_col1">

            <!--2014-3-19, update by Qihy, 增加两个字母表示的列的转换，三个字母的暂时未考虑 start-->
            <!--<xsl:choose>
              <xsl:when test="$colSeq1 &lt; 10">
                <xsl:value-of select="translate($colSeq1,'123456789','ABCDEFGHI')"/>
              </xsl:when>
              <xsl:when test="($colSeq1 &lt;19) and ($colSeq1 &gt; 9)">
                <xsl:value-of select="translate($colSeq1 - 9,'123456789','JKLMNOPQR')"/>
              </xsl:when>
              <xsl:when test="($colSeq1 &lt;27) and ($colSeq1 &gt; 18)">
                <xsl:value-of select="translate($colSeq1 - 18,'12345678','STUVWXYZ')"/>
              </xsl:when>
            </xsl:choose>-->
            <xsl:call-template name ="ColIndex1">
              <xsl:with-param name ="colSeq" select ="$colSeq1"/>
            </xsl:call-template>
            
          </xsl:variable>
          <xsl:variable name="sqref_subVar_row1">
            <xsl:value-of select="translate(substring-before($subRef_before,'C'),'R', '')"/>
          </xsl:variable>
          <xsl:variable name ="colSeq2" select ="substring-after($subRef_after,'C')"/>
          <xsl:variable name="sqref_subVar_col2">
            <!--<xsl:choose>
              <xsl:when test="$colSeq2 &lt; 10">
                <xsl:value-of select="translate($colSeq2,'123456789','ABCDEFGHI')"/>
              </xsl:when>
              <xsl:when test="($colSeq2 &lt;19) and ($colSeq2 &gt; 9)">
                <xsl:value-of select="translate($colSeq2 - 9,'123456789','JKLMNOPQR')"/>
              </xsl:when>
              <xsl:when test="($colSeq2 &lt;27) and ($colSeq2 &gt; 18)">
                <xsl:value-of select="translate($colSeq2 - 18,'12345678','STUVWXYZ')"/>
              </xsl:when>
            </xsl:choose>-->
            <xsl:call-template name ="ColIndex1">
              <xsl:with-param name ="colSeq" select ="$colSeq2"/>
            </xsl:call-template>
            
          </xsl:variable>
          <xsl:variable name="sqref_subVar_row2">
            <xsl:value-of select="translate(substring-before($subRef_after,'C'),'R', '')"/>
          </xsl:variable>    
          <xsl:variable name="sqref_subVar_part">
            <xsl:value-of select="concat($sqref_subVar_col1, $sqref_subVar_row1,':',$sqref_subVar_col2, $sqref_subVar_row2)"/>
          </xsl:variable>
          <xsl:variable name="sqref_subVar">
            <xsl:choose>
              <xsl:when test="$seq=1">
                <xsl:value-of select="$sqref_subVar_part"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat($sqref,' ',$sqref_subVar_part)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:call-template name="combineRef1">
            <xsl:with-param name="count" select="number($count) - 1"/>
            <xsl:with-param name="seq" select="number($seq) + 1"/>
            <xsl:with-param name="sqref" select="$sqref_subVar"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name ="colSeq" select ="substring-after(substring-after($subRef,'!'),'C')"/>
          <xsl:variable name="sqref_subVar_col">
            
            <!--<xsl:choose>
              <xsl:when test="$colSeq &lt; 10">
                <xsl:value-of select="translate($colSeq,'123456789','ABCDEFGHI')"/>
              </xsl:when>
              <xsl:when test="($colSeq &lt;19) and ($colSeq &gt; 9)">
                <xsl:value-of select="translate($colSeq - 9,'123456789','JKLMNOPQR')"/>
              </xsl:when>
              <xsl:when test="($colSeq &lt;27) and ($colSeq &gt; 18)">
                <xsl:value-of select="translate($colSeq - 18,'12345678','STUVWXYZ')"/>
              </xsl:when>
            </xsl:choose>-->
            <xsl:call-template name ="ColIndex1">
              <xsl:with-param name ="colSeq" select ="$colSeq"/>
            </xsl:call-template>
            <!--2014-3-19-->
          </xsl:variable>
          <xsl:variable name="sqref_subVar_row">
            <xsl:value-of select="translate(substring-before(substring-after($subRef,'!'),'C'),'R', '')"/>
          </xsl:variable>
          <xsl:variable name="sqref_subVar_part">
            <xsl:value-of select="concat($sqref_subVar_col, $sqref_subVar_row)"/>
          </xsl:variable>
          <xsl:variable name="sqref_subVar">
            <xsl:choose>
              <xsl:when test="$seq=1">
                <xsl:value-of select="$sqref_subVar_part"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat($sqref,' ',$sqref_subVar_part)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:call-template name="combineRef1">
            <xsl:with-param name="count" select="number($count) - 1"/>
            <xsl:with-param name="seq" select="number($seq) + 1"/>
            <xsl:with-param name="sqref" select="$sqref_subVar"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
      
    </xsl:if>
  </xsl:template>
  <!--2014-3-18 end-->

  <!--2014-3-27, update by Qihy, dxfId取值不正确， start -->
  <xsl:template name="computeCountN">
    <xsl:param name="sum"/>
    <xsl:param name="id"/>
    <xsl:choose>
      <xsl:when test="$id = 1">
        <xsl:value-of select="$sum"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="computeCountN">
          <xsl:with-param name="sum" select ="$sum + number(parent::表:单工作表/表:单条件格式化[@conditionID = $id - 1]/@conditionCount)"/>
          <xsl:with-param name="id" select="$id - 1"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--2014-3-27 end-->
  
  </xsl:stylesheet>  

<!--
				<xsl:if test=".//表:图表">
          <xsl:variable name="sheetNo">
            <xsl:value-of select="./@uof:sheetNo"/>
          </xsl:variable>
					<Relationship>
						<xsl:attribute name="Id">
							<xsl:value-of select="'rId1'"/>
						</xsl:attribute>
						<xsl:attribute name="Type">
							<xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
						</xsl:attribute>
						<xsl:attribute name="Target">
							<xsl:value-of select="concat('../drawings/drawing',$sheetNo,'.xml')"/>
						</xsl:attribute>
					</Relationship>
				</xsl:if>
				<xsl:if test="表:工作表内容/uof:锚点 and 表:工作表内容[not(表:图表)]">
          <Relationship>
						<xsl:attribute name="Id">
							<xsl:value-of select="'rId1'"/>
						</xsl:attribute>
						<xsl:attribute name="Type">
							<xsl:value-of select="'http://purl.oclc.org/ooxml/officeDocument/relationships/drawing'"/>
						</xsl:attribute>
						<xsl:attribute name="Target">
							<xsl:value-of select="concat('../drawings/drawing',$relsSheet,'.xml')"/>
						</xsl:attribute>
					</Relationship>
				</xsl:if>
        -->


<!--what does this means?-->
<!--Need Consideration Codes-->
<!--Marked by LDM in 2010/12/19-->
<!--
      <xsl:if test="./字:句 and not(./@表:数据类型) or ./@表:数据类型!='text'or ./@表:数据类型='error'">
        <v>
          <xsl:choose>
            <xsl:when test="表:数据/@表:数据类型='boolean'">
              <xsl:if test="表:数据/字:句/字:文本串='TRUE'">
                <xsl:value-of select="'1'"/>
              </xsl:if>
              <xsl:if test="表:数据/字:句/字:文本串='FALSE'">
                <xsl:value-of select="'0'"/>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="表:数据/字:句/字:文本串"/>
            </xsl:otherwise>
          </xsl:choose>
        </v>
      </xsl:if>
      -->