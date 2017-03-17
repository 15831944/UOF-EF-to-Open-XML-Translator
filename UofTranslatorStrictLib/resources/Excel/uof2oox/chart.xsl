<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:uof="http://schemas.uof.org/cn/2009/uof" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:图="http://schemas.uof.org/cn/2009/graph" xmlns:字="http://schemas.uof.org/cn/2009/wordproc" xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet" xmlns:图表="http://schemas.uof.org/cn/2009/chart" xmlns:xsd="http://www.ord.com" xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<xsl:import href="chartArea.xsl"/>
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--ChartSheet-->
	<xsl:template name="chartsheet">
		<xsl:apply-templates select="../图表:图表_E837"/>
	</xsl:template>
	<!--图表主体-->
	<!--Modified by LDM in 2010/12/16-->
	<xsl:template name="chart">
		<xsl:apply-templates select="./图表:图表_E837"/>
	</xsl:template>
	<!--Modified by LDM in 2010/12/19-->
	<!--图表模板-->
	<xsl:template match="图表:图表_E837" name="ChartTrans">
		<!--<xsl:variable name="chartNo">
      <xsl:value-of select="ancestor::表:单图表/@uof:chartNo"/>
    </xsl:variable>-->
		<c:chartSpace>
			<c:roundedCorners val="0"/>
			<c:chart>
				<xsl:if test="./图表:标题_E736">
					<c:title>
						<xsl:for-each select="图表:标题_E736">
							<c:tx>
								<c:rich>
									<a:bodyPr>
										<xsl:if test="图表:对齐_E726">
											<xsl:call-template name="Alignment"/>
										</xsl:if>
									</a:bodyPr>
									<a:p>
										<a:r>
											<a:rPr>
												<xsl:for-each select="./ancestor::图表:图表_E837/图表:标题_E736">
													<xsl:call-template name="Font"/>
												</xsl:for-each>
											</a:rPr>
											<xsl:if test="@名称_E742">
												<a:t>
													<xsl:value-of select="./@名称_E742"/>
												</a:t>
											</xsl:if>
										</a:r>
									</a:p>
								</c:rich>
							</c:tx>
							<xsl:if test="图表:填充_E746 or 图表:边框线_4226">
								<c:spPr>
									<xsl:if test="图表:填充_E746">
										<xsl:call-template name="Fill"/>
									</xsl:if>
									<xsl:if test="图表:边框线_4226">
										<xsl:call-template name="Border"/>
									</xsl:if>
								</c:spPr>
							</xsl:if>
						</xsl:for-each>
						<c:overlay val="0"/>
					</c:title>
				</xsl:if>
				<xsl:call-template name="ChartType"/>
				<!--floor and backWall-->
				<xsl:for-each select="./图表:基底_E7A4">
					<c:floor>
						<!--填充-->
						<c:spPr>
							<xsl:if test="图表:填充_E746">
								<xsl:call-template name="Fill"/>
							</xsl:if>
							<!--边框-->
							<xsl:if test="图表:边框线_4226">
								<xsl:call-template name="Border"/>
							</xsl:if>
						</c:spPr>
					</c:floor>
				</xsl:for-each>
				<xsl:for-each select="./图表:背景墙_E7A1">
					<c:backWall>
						<!--填充-->
						<c:spPr>
							<xsl:if test="图表:填充_E746">
								<xsl:call-template name="Fill"/>
							</xsl:if>
							<!--边框-->
							<xsl:if test="图表:边框线_4226">
								<xsl:call-template name="Border"/>
							</xsl:if>
						</c:spPr>
					</c:backWall>
				</xsl:for-each>
				<xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F">
					<xsl:call-template name="PlotArea"/>
				</xsl:if>
				<xsl:if test="./图表:图例_E794">
					<xsl:apply-templates select="./图表:图例_E794"/>
				</xsl:if>
			</c:chart>
			<xsl:apply-templates select="./图表:图表区_E743"/>
		</c:chartSpace>
	</xsl:template>
	<!--Need Consideration Codes==========================================================================================================-->
	<!--Modified by LDM in 2010/12/17-->
	<!--
  <xsl:template name="chartTrans">
    <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
      <xsl:if test="表:数据系列集/表:数据系列">
        <xsl:for-each select="表:数据系列集/表:数据系列">
          <xsl:variable name="ser" select="@表:系列"/>
          <c:ser>
            <c:marker>
              <c:symbol val="none"/>
            </c:marker>
            <xsl:if test="@表:系列 ">
              <xsl:call-template name="shujuxilieji_xiliehao">
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="ancestor::表:图表">
              <xsl:for-each select="ancestor::表:图表/表:数据源/表:系列">

                <xsl:variable name="bianhao3">
                  <xsl:number count="表:系列" level="single"/>
                </xsl:variable>

                <xsl:if test="$bianhao3=$ser">
                  <xsl:call-template name="shujuyuan_xilieming">
                  </xsl:call-template>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="表:填充 or 表:边框">
              <c:spPr>
                <xsl:call-template name="shujuxilieji_tcandbk">
                </xsl:call-template>
              </c:spPr>
            </xsl:if>
            <xsl:if test="ancestor::表:图表">
              <xsl:for-each select="ancestor::表:图表/表:数据点集/表:数据点">
                <xsl:if test="@表:系列=$ser">
                  <c:dPt>
                    <xsl:call-template name="shujudianji_dian">
                    </xsl:call-template>
                    <xsl:if test="表:填充 or 表:边框">
                      <c:spPr>
                        <xsl:call-template name="shujudianji_tcandbk">
                        </xsl:call-template>
                      </c:spPr>
                    </xsl:if>
                  </c:dPt>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="表:显示标志">
              <c:dLbls>
                <c:txPr>
                  <a:bodyPr/>

                  <a:p>
                    <a:pPr>
                      <a:defRPr sz="1200">

                        <a:latin typeface="Times New Roman"/>
                        <a:ea typeface="永中宋体"/>
                      </a:defRPr>
                    </a:pPr>

                  </a:p>
                </c:txPr>

                <xsl:if test="ancestor::表:图表/表:数据点集/表:数据点">
                  <xsl:for-each select="ancestor::表:图表/表:数据点集/表:数据点">
                    <xsl:if test="@表:系列=$ser">
                      <xsl:call-template name="shujudianji_xianshi">
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>

                <xsl:call-template name="shujuxilieji_xianshi">
                </xsl:call-template>
              </c:dLbls>
              <xsl:if test="ancestor::表:图表/表:数据系列集/表:数据系列/表:显示标志">
                <c:dLbls>
                  <c:txPr>
                    <a:bodyPr/>

                    <a:p>
                      <a:pPr>
                        <a:defRPr sz="1200">

                          <a:latin typeface="Times New Roman"/>
                          <a:ea typeface="永中宋体"/>
                        </a:defRPr>
                      </a:pPr>

                    </a:p>
                  </c:txPr>

                  <xsl:for-each select="ancestor::表:图表/表:数据系列集/表:数据系列/表:显示标志">
                    <xsl:if test="@表:系列=$ser">
                      <xsl:call-template name="shujudianji_xianshi">
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:for-each>

                  <xsl:call-template name="shujuxilieji_xianshi">
                  </xsl:call-template>
                </c:dLbls>
              </xsl:if>
            </xsl:if>

            <xsl:for-each select="ancestor::表:图表/表:数据源/表:系列">
              <xsl:variable name="bianhao3">
                <xsl:number count="表:系列" level="single"/>
              </xsl:variable>
              <xsl:if test="$bianhao3=$ser">
                <xsl:call-template name="shujuyuan_flmandxlz">
                </xsl:call-template>
              </xsl:if>
            </xsl:for-each>
          </c:ser>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  -->
	<!--图例项？-->
	<!--
      <xsl:if test="./表:图例项[@表:系列]">
        <xsl:for-each select="表:图例/表:图例项[@表:系列]">
          <xsl:variable name="legend" select="@表:系列"/>
          <c:legendEntry>
            <xsl:if test="@表:系列">
              <xsl:call-template name="tuli_tlxhao">
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="表:字体">
              <c:txPr>
                <xsl:call-template name="tuli_tlxziti">
                </xsl:call-template>
              </c:txPr>
            </xsl:if >
          </c:legendEntry>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="表:图例/表:图例项[not(@表:系列)]">
        <xsl:for-each select="表:图例/表:图例项[not(@表:系列)]">
          <xsl:variable name="legend" select="@表:系列"/>
          <xsl:if test="表:字体">
            <c:txPr>
              <xsl:call-template name="tuli_tlxziti">
              </xsl:call-template>
            </c:txPr>
          </xsl:if >
        </xsl:for-each>
      </xsl:if>
      -->
	<!--Obsoluted Codes-->
	<!--obsoluted-->
	<!--已合并-->
	<!--
  <xsl:if test="@表:类型!='pie'">
    <xsl:if test="表:网格线集/表:网格线[@表:位置 = 'category axis'] or 表:分类轴">
      <c:catAx>
        <xsl:if test="表:分类轴">
          <c:axId val="00000000"/>
        </xsl:if>
        <xsl:if test="表:分类轴/表:刻度">
          <xsl:call-template name="fenleizhou_kedu">
          </xsl:call-template>
        </xsl:if>

        <xsl:if test="表:网格线集/表:网格线[@表:位置 = 'category axis']">
          <xsl:for-each select="表:网格线集/表:网格线[@表:位置 = 'category axis']">
            <xsl:call-template name="wanggexian">
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="表:标题集/表:标题[@表:位置='category axis']">
          <c:title>
            <xsl:for-each select="表:标题集/表:标题">

              <xsl:if test="@表:位置='category axis'">

                <xsl:choose>
                  <xsl:when test="@表:名称 or 表:字体 or 表:对齐">
                    <c:tx>
                      <c:rich>
                        <a:bodyPr>

                          <xsl:if test="表:对齐"	>
                            <xsl:call-template name="bt_duiqi">
                            </xsl:call-template>
                          </xsl:if>
                        </a:bodyPr>
                        <a:p>
                          <xsl:choose>
                            <xsl:when test="表:字体/字:字体[@字:西文字体引用 or @字:中文字体引用]">
                              <xsl:call-template name="defaultParaProp"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <a:pPr>
                                <a:defRPr>





                                  <a:latin>
                                    <xsl:attribute name="typeface">
                                      <xsl:value-of select="'Times New Roman'"/>
                                    </xsl:attribute>
                                  </a:latin>


                                  <a:ea>

                                    <xsl:attribute name="typeface">
                                      <xsl:value-of select="'永中宋体'"/>
                                    </xsl:attribute>
                                  </a:ea>

                                </a:defRPr>
                              </a:pPr>
                            </xsl:otherwise>
                          </xsl:choose>
                          <a:r>


                            <a:rPr>
                              <xsl:call-template name="bt_ziti">
                              </xsl:call-template>
                            </a:rPr>


                            <xsl:if test="@表:名称">
                              <a:t>
                                <xsl:call-template name="bt_mingcheng">
                                </xsl:call-template>
                              </a:t>
                            </xsl:if>
                          </a:r>
                        </a:p>
                      </c:rich>
                    </c:tx>
                  </xsl:when>

                  <xsl:when test="not(表:字体)">
                    <c:tx>
                      <a:bodyPr/>

                      <a:p>
                        <a:pPr>
                          <a:defRPr>
                            <xsl:attribute name="sz">

                              <xsl:value-of select="'1200'"/>

                            </xsl:attribute>





                            <a:latin>
                              <xsl:attribute name="typeface">
                                <xsl:value-of select="'Times New Roman'"/>
                              </xsl:attribute>
                            </a:latin>


                            <a:ea>

                              <xsl:attribute name="typeface">
                                <xsl:value-of select="'永中宋体'"/>
                              </xsl:attribute>
                            </a:ea>

                          </a:defRPr>
                        </a:pPr>
                      </a:p>
                    </c:tx>
                  </xsl:when>

                </xsl:choose>

                <xsl:if test="表:边框 or 表:填充">
                  <c:spPr>
                    <xsl:if test="表:填充">
                      <xsl:call-template name="bt_tianchong">
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="表:边框">
                      <xsl:call-template name="bt_biankuang">
                      </xsl:call-template>
                    </xsl:if>
                  </c:spPr>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </c:title>

        </xsl:if>


        <xsl:if test="表:分类轴[@表:主刻度类型 or @表:次刻度类型 or @表:刻度线标志]">
          <xsl:call-template name="fenleizhou_shuxing">
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="表:分类轴/表:线型">
          <c:spPr>
            <xsl:call-template name="fenleizhou_xianxing">
            </xsl:call-template>
          </c:spPr>
        </xsl:if>


        <xsl:if test="(表:分类轴/表:字体) or 表:分类轴[表:对齐/表:文字方向|表:对齐/表:旋转角度]">
          <c:txPr>
            <xsl:if test="表:分类轴/表:对齐">
              <xsl:call-template name="cataxvert"/>
            </xsl:if>
            <xsl:if test="not(表:分类轴/表:对齐)">
              <a:bodyPr/>
            </xsl:if>
            <xsl:if test="表:分类轴/表:字体">
              <xsl:call-template name="fenleizhoucharacter"/>
            </xsl:if>

          </c:txPr>
        </xsl:if>
        <c:crossAx val="11111111"/>
        <xsl:if test="表:分类轴/表:对齐[表:偏移量]">
          <xsl:call-template name="fenleizhou_pianyi">
          </xsl:call-template>
        </xsl:if>
      </c:catAx>
    </xsl:if>
    <xsl:if  test="not(表:分类轴) and @表:子类型!='doughnut_exploded' and @表:子类型!='doughnut_standard' and @表:子类型!='radar_standard' and @表:子类型!='radar_marker' and @表:子类型!='radar_filled'">
      <c:dateAx>
        <c:axId val="00000000"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="b"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="11111111"/>
        <c:crosses val="autoZero"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
    </xsl:if>

    <xsl:if test="表:网格线集/表:网格线[@表:位置='value axis'] or 表:数值轴">
      <c:valAx>
        <c:axId val="11111111"/>
        <xsl:if test="表:数值轴/表:刻度">
          <xsl:call-template name="shuzhizhou_kedu1">
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="表:网格线集/表:网格线[@表:位置 = 'value axis']">
          <xsl:for-each select="表:网格线集/表:网格线[@表:位置 = 'value axis']">
            <xsl:call-template name="wanggexian">
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="表:标题集/表:标题[@表:位置='value axis']">
          <c:title>
            <xsl:for-each select="表:标题集/表:标题">
              <xsl:if test="@表:位置='value axis'">
                <xsl:choose>
                  <xsl:when test="@表:名称 or 表:字体 or 表:对齐">
                    <c:tx>
                      <c:rich>
                        <a:bodyPr>
                          <xsl:if test="表:对齐"	>
                            <xsl:call-template name="Alignment">
                            </xsl:call-template>
                          </xsl:if>
                        </a:bodyPr>
                        <a:p>
                          <xsl:choose>
                            <xsl:when test="表:字体/字:字体[@字:西文字体引用 or @字:中文字体引用]">
                              <xsl:call-template name="defaultParaProp"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <a:pPr>
                                <a:defRPr>
                                  <a:latin>
                                    <xsl:attribute name="typeface">
                                      <xsl:value-of select="'Times New Roman'"/>
                                    </xsl:attribute>
                                  </a:latin>


                                  <a:ea>

                                    <xsl:attribute name="typeface">
                                      <xsl:value-of select="'永中宋体'"/>
                                    </xsl:attribute>
                                  </a:ea>

                                </a:defRPr>
                              </a:pPr>
                            </xsl:otherwise>
                          </xsl:choose>
                          <a:r>

                            <xsl:if test="表:字体"	>
                              <a:rPr>
                                <xsl:call-template name="bt_ziti">
                                </xsl:call-template>
                              </a:rPr>
                            </xsl:if>

                            <xsl:if test="@表:名称">
                              <a:t>
                                <xsl:call-template name="bt_mingcheng">
                                </xsl:call-template>
                              </a:t>
                            </xsl:if>
                          </a:r>
                        </a:p>
                      </c:rich>
                    </c:tx>
                  </xsl:when>
                  <xsl:when test="not(表:字体)">
                    <c:tx>
                      <a:bodyPr/>
                      <a:p>
                        <a:pPr>
                          <a:defRPr>
                            <xsl:attribute name="sz">
                              <xsl:value-of select="'1200'"/>
                            </xsl:attribute>
                            <a:latin>
                              <xsl:attribute name="typeface">
                                <xsl:value-of select="'Times New Roman'"/>
                              </xsl:attribute>
                            </a:latin>
                            <a:ea>
                              <xsl:attribute name="typeface">
                                <xsl:value-of select="'永中宋体'"/>
                              </xsl:attribute>
                            </a:ea>
                          </a:defRPr>
                        </a:pPr>
                      </a:p>
                    </c:tx>
                  </xsl:when>

                </xsl:choose>

                <xsl:if test="表:边框 or 表:填充">
                  <c:spPr>
                    <xsl:if test="表:填充">
                      <xsl:call-template name="bt_tianchong">
                      </xsl:call-template>
                    </xsl:if>
                    <xsl:if test="表:边框">
                      <xsl:call-template name="bt_biankuang">
                      </xsl:call-template>
                    </xsl:if>
                  </c:spPr>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </c:title>
        </xsl:if>
        <xsl:if test="表:数值轴[@表:主刻度类型 or @表:次刻度类型 or @表:刻度线标志]">
          <xsl:call-template name="shuzhizhou_shuxing">
          </xsl:call-template>
        </xsl:if>
        <xsl:if test="表:数值轴/表:线型">
          <c:spPr>
            <xsl:call-template name="shuzhizhou_xianxing">
            </xsl:call-template>
          </c:spPr>
        </xsl:if>

        <xsl:if test="表:数值轴/表:字体 or 表:数值轴/表:对齐">
          <c:txPr>
            <xsl:call-template name="shuzhizhou_ztanddq">
            </xsl:call-template>
          </c:txPr>
        </xsl:if>
        <c:crossAx val="00000000"/>
        <xsl:if test="表:数值轴/表:刻度">
          <xsl:call-template name="shuzhizhou_kedu2">
          </xsl:call-template>
        </xsl:if>
      </c:valAx>
    </xsl:if>
  </xsl:if>
  -->
	<!--
  <xsl:template name="adjust4">

    <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
      <xsl:if test="表:数据系列集/表:数据系列">
        <xsl:for-each select="表:数据系列集/表:数据系列">
          <xsl:variable name="ser" select="@表:系列"/>
          <c:ser>
            <c:explosion val="25"/>
            <xsl:if test="@表:系列 ">
              <xsl:call-template name="shujuxilieji_xiliehao">
              </xsl:call-template>
            </xsl:if>
            <xsl:if test="ancestor::表:图表">
              <xsl:for-each select="ancestor::表:图表/表:数据源/表:系列">

                <xsl:variable name="bianhao3">
                  <xsl:number count="表:系列" level="single"/>
                </xsl:variable>

                <xsl:if test="$bianhao3=$ser">
                  <xsl:call-template name="shujuyuan_xilieming">
                  </xsl:call-template>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="表:填充 or 表:边框">
              <c:spPr>
                <xsl:call-template name="shujuxilieji_tcandbk">
                </xsl:call-template>
              </c:spPr>
            </xsl:if>
            <xsl:if test="ancestor::表:图表">
              <xsl:for-each select="ancestor::表:图表/表:数据点集/表:数据点">
                <xsl:if test="@表:系列=$ser">
                  <c:dPt>
                    <xsl:call-template name="shujudianji_dian">
                    </xsl:call-template>
                    <xsl:if test="表:填充 or 表:边框">
                      <c:spPr>
                        <xsl:call-template name="shujudianji_tcandbk">
                        </xsl:call-template>
                      </c:spPr>
                    </xsl:if>
                  </c:dPt>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="表:显示标志">
              <c:dLbls>
                <c:txPr>
                  <a:bodyPr/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr sz="1200">
                        <a:latin typeface="Times New Roman"/>
                        <a:ea typeface="永中宋体"/>
                      </a:defRPr>
                    </a:pPr>

                  </a:p>
                </c:txPr>
                <xsl:if test="ancestor::表:图表/表:数据点集/表:数据点">
                  <xsl:for-each select="ancestor::表:图表/表:数据点集/表:数据点">
                    <xsl:if test="@表:系列=$ser">
                      <xsl:call-template name="shujudianji_xianshi">
                      </xsl:call-template>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>
                <xsl:call-template name="shujuxilieji_xianshi">
                </xsl:call-template>
              </c:dLbls>
            </xsl:if>
            <xsl:if test="ancestor::表:图表">
              <xsl:for-each select="ancestor::表:图表/表:数据源/表:系列">
                <xsl:variable name="bianhao3">
                  <xsl:number count="表:系列" level="single"/>
                </xsl:variable>
                <xsl:if test="$bianhao3=$ser">
                  <xsl:call-template name="shujuyuan_flmandxlz">
                  </xsl:call-template>
                </xsl:if>
              </xsl:for-each>
            </xsl:if>
          </c:ser>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  -->
	<!--Marked by LDM in 2010/12/19-->
	<!--
  <xsl:template name="back_plotArea">
  <c:plotArea>
    <xsl:if test="表:绘图区">
      <xsl:apply-templates select="表:绘图区"/>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="@表:类型='column' and @表:子类型='column_clustered'">
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <xsl:call-template name="chartTrans"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='column_stacked'">
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='column_100%_stacked'">
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='column_clustered_3D'">
        <c:bar3DChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='column_100%_3D'">
        <c:bar3DChart>
          <c:barDir val="col"/>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:shape val="box"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
          <c:axId val="0"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='column_stacked_3D'">
        <c:bar3DChart>
          <c:barDir val="col"/>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
          <c:axId val="0"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_standard'">
        <c:lineChart>
          <c:grouping val="standard"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_stacked'">
        <c:lineChart>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_100%_stacked'">
        <c:lineChart>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_standard_marker'">
        <c:lineChart>
          <c:grouping val="standard"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_stcked_marker'">
        <c:lineChart>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='line' and @表:子类型='line_100%_stacked_marker'">
        <c:lineChart>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:lineChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_clustered'">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="clustered"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_stacked'">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_100%_stacked'">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_clustered_3D'">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="clustered"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_stacked_3D'">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='bar' and @表:子类型='bar_100%_3D'">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="percentStacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:shape val="box"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
          <c:axId val="0"/>
        </c:bar3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_standard'">
        <c:pieChart>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:overlap val="100"/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:pieChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_exploded'">
        <c:pieChart>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:pieChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_3D'">
        <c:pie3DChart>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:pie3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_exploded_3D'">
        <c:pie3DChart>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:pie3DChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_of-pie'">
        <c:ofPieChart>
          <c:ofPieType val="pie"/>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:ofPieChart>
      </xsl:when>
      <xsl:when test="@表:类型='pie' and @表:子类型='pie_of-bar'">
        <c:ofPieChart>
          <c:ofPieType val="bar"/>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:ofPieChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='area_standard'">
        <c:areaChart>
          <c:grouping val="standard"/>

          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:areaChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='area_stacked'">
        <c:areaChart>
          <c:grouping val="stacked"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:areaChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='area_100%_stacked'">
        <c:areaChart>
          <c:grouping val="percentStacked"/>

          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:areaChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='doughnut_exploded'">
        <c:doughnutChart>
          <c:varyColors val="1"/>

          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:doughnutChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='doughnut_standard'">
        <c:doughnutChart>
          <c:varyColors val="1"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:doughnutChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='radar_standard'">
        <c:radarChart>
          <c:radarStyle val="marker"/>

          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:radarChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='radar_marker'">
        <c:radarChart>
          <c:radarStyle val="marker"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:radarChart>
      </xsl:when>
      <xsl:when test="@表:类型='column' and @表:子类型='radar_filled'">
        <c:radarChart>
          <c:radarStyle val="filled"/>

          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:serLines/>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:radarChart>
      </xsl:when>
      <xsl:otherwise>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <xsl:if test="表:数据系列集 or 表:数据点集 or 表:数据源">
            <xsl:call-template name="chartTrans"/>
          </xsl:if>
          <c:axId val="00000000"/>
          <c:axId val="11111111"/>
        </c:barChart>
      </xsl:otherwise>
    </xsl:choose>
-->
	<!--饼图没有网格线、坐标轴和数据表-->
	<!--Modified by LDM in 2010/12/18-->
	<!--
    <xsl:if test="@表:类型!='pie'">
      <c:catAx>
        <xsl:if test ="表:分类轴">
          <xsl:apply-templates select="./表:分类轴"/>
        </xsl:if>
        <xsl:if test="./表:网格线集/表:网格线[@表:位置='category axis']">
          <xsl:apply-templates select="./表:网格线集/表:网格线[@表:位置='category axis']"/>
        </xsl:if>
        <xsl:if test="./表:标题集/表:标题[@表:位置='category axis']">
          <xsl:apply-templates select="./表:标题集/表:标题[@表:位置='category axis']"/>
        </xsl:if>
      </c:catAx>
      <c:valAx>
        <xsl:if test ="./表:数值轴">
          <xsl:apply-templates select="./表:数值轴"/>
        </xsl:if>
        <xsl:if test="./表:网格线集/表:网格线[@表:位置='value axis']">
          <xsl:apply-templates select="./表:网格线集/表:网格线[@表:位置='value axis']"/>
        </xsl:if>
        <xsl:if test="./表:标题集/表:标题[@表:位置='value axis']">
          <xsl:apply-templates select="./表:标题集/表:标题[@表:位置='value axis']"/>
        </xsl:if>
      </c:valAx>
    </xsl:if>
    -->
	<!--
  </c:plotArea>
  </xsl:template>
  -->
</xsl:stylesheet>
