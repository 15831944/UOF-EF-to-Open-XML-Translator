<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:fo="http://www.w3.org/1999/XSL/Format" 
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
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
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                >
  <xsl:template name="Style">
		<styleSheet  
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      mc:Ignorable="x14ac"
      xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
			<xsl:if test="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916">
				<xsl:variable name="CellStyNum" select="count(式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916)"/>
        <!--yanghaojie-start-注释-->
        <!--
				<numFmts>
					<xsl:attribute name="count">
						<xsl:value-of select="$CellStyNum"/>
					</xsl:attribute>
					<xsl:for-each select="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[not(./@标识符_E7AC='DEFSTYLE')]">
						<xsl:if test="表:数字格式_E7A9">
							<numFmt>
								<xsl:attribute name="numFmtId">
									<xsl:value-of select="position()"/>
								</xsl:attribute>
								<xsl:attribute name="formatCode">
									<xsl:value-of select="表:数字格式_E7A9/@格式码_E73F"/>
								</xsl:attribute>
							</numFmt>
						</xsl:if>
            
            
            <xsl:if test ="not(表:数字格式_E7A9)">
              <numFmt>
                <xsl:attribute name="numFmtId">
                  <xsl:value-of select="position()"/>
                </xsl:attribute>
                <xsl:attribute name="formatCode">
                  <xsl:value-of select="'general'"/>
                </xsl:attribute>
              </numFmt>
            </xsl:if>
					</xsl:for-each>
				</numFmts>-->
        <!--yanghaojie-end-注释-->
				<fonts>
          <xsl:attribute name="count">
            <xsl:value-of select="$CellStyNum + 1"/>
          </xsl:attribute>
          
          <!--添加默认的两种类型-->
          <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="宋体"/>
            <family val="2"/>
            <charset val="134"/>
            <scheme val="minor"/>
          </font>
					<!--字体集-->
					<xsl:for-each select="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[not(@标识符_E7AC='DEFSTYLE')]">
						<xsl:if test="表:字体格式_E7A7">
							<font>
                <!--bold-->
                <xsl:if test="表:字体格式_E7A7/字:是否粗体_4130='true'">
                  <b/>
                </xsl:if>
                <!--italic-->
                <xsl:if test="表:字体格式_E7A7/字:是否斜体_4131='true'">
                  <i/>
                </xsl:if>
                <!--strike-->
                <xsl:if test="表:字体格式_E7A7/字:删除线_4135 != 'none'">
                  <strike/>
                </xsl:if>
                <!--空心-->
                <xsl:if test="表:字体格式_E7A7/字:是否空心_413E='true'">
                  <outline/>
                </xsl:if>
                <!--阴影-->
                <xsl:if test="表:字体格式_E7A7/字:是否阴影_4140='true'">
                  <shadow/>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7[字:下划线_4136]">
                  <u>
                    <xsl:variable name="lineType" select="表:字体格式_E7A7/字:下划线_4136/@线型_4137"/>
                    <xsl:attribute name="val">
                      <xsl:choose>
                        <xsl:when test="$lineType='none'">
                          <xsl:value-of select="'none'"/>
                        </xsl:when>
                        <xsl:when test="$lineType='single'">
                          <xsl:value-of select="'single'"/>
                        </xsl:when>
                        <xsl:when test="$lineType='double'">
                          <xsl:value-of select="'double'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="'single'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:attribute>
                  </u>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7[字:上下标类型_4143]">
                  <xsl:if test="表:字体格式_E7A7/字:上下标类型_4143='sub'">
                    <vertAlign val="subscript"/>
                  </xsl:if>
                  <xsl:if test="表:字体格式_E7A7/字:上下标类型_4143='sup'">
                    <vertAlign val="superscript"/>
                  </xsl:if>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7/字:字体_4128[@字号_412D]">
                  <sz>
                    <xsl:attribute name="val">
                      <xsl:value-of select="./表:字体格式_E7A7/字:字体_4128/@字号_412D"/>
                    </xsl:attribute>
                  </sz>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@字号_412D)]">
                  <sz val="11"/>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F]">
                  <xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F='auto']">
                    <color theme="1"/>
                  </xsl:if>
                  <xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F!='auto']">
                    <xsl:variable name="fontColor" select="表:字体格式_E7A7/字:字体_4128/@颜色_412F"/>
                    <color>
                      
                      <!--2014-4-28, update by Qihy, 超链接点击链接，颜色不正确， start-->
                      <xsl:choose>
                        <xsl:when test="$fontColor = '#0000ff' or $fontColor = '#0000FF'">
                          <xsl:attribute name="theme">
                            <xsl:value-of select="10"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="rgb">
                            <xsl:value-of select="concat('FF',substring-after($fontColor,'#'))"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                      <!--2014-4-28 end-->
                      
                    </color>
                  </xsl:if>
                </xsl:if>
                <xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@颜色_412F)]">
                  <color theme="1"/>
                </xsl:if>
                
								<xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@中文字体引用_412A)]">
									<name val="宋体"/>
								</xsl:if>
                <xsl:if test="表:字体格式_E7A7/字:字体_4128[@中文字体引用_412A]">
                  <name>
                    <xsl:attribute name ="val">
                      <xsl:variable name="esFontRef">
                        <xsl:value-of select="表:字体格式_E7A7/字:字体_4128/@中文字体引用_412A"/>
                      </xsl:variable>
                      
                     	<!--20130319 gaoyuwei Bug_2750 uof-oo 功能测试B19-B24，字体样式不正确-->

						<xsl:variable name="LatinFontRef">
							<xsl:value-of select="表:字体格式_E7A7/字:字体_4128/@西文字体引用_4129"/>
						</xsl:variable>
						<xsl:choose>
						<xsl:when test="contains($esFontRef,$LatinFontRef)">
							<xsl:value-of select="translate(/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902 = $esFontRef]/@名称_9903,'永中','')"/>
						</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="translate(/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902 = $LatinFontRef]/@名称_9903,'永中','')"/>
							</xsl:otherwise>
							</xsl:choose>
						
						<!--end-->

                    </xsl:attribute>
                  </name>
                </xsl:if>
								<charset>
									<xsl:attribute name="val">
										<xsl:value-of select="'134'"/>
									</xsl:attribute>
								</charset>
                
                <!--
                <xsl:if test="表:字体格式/字:下划线[@字:颜色]">
                  <xsl:if test="not(表:字体格式/字:下划线/@字:颜色 = 'auto')">
                  <color>
                    <xsl:attribute name="rgb">
                      <xsl:value-of select="substring-after(表:字体格式/字:下划线/@字:颜色,'#')"/>
                    </xsl:attribute>
                  </color>
                </xsl:if>
                </xsl:if>
                -->
								
							</font>
						</xsl:if>
						<xsl:if test="not(表:字体格式_E7A7)">
							<font/>
						</xsl:if>
					</xsl:for-each>
				</fonts>
        
        <!--Not Finished-->
        <!--Marked by LDM in 2011/01/07-->
				<fills>
					<xsl:attribute name="count">
						<xsl:value-of select="$CellStyNum + 1"/>
					</xsl:attribute>
					<xsl:for-each select="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916">
						<xsl:if test="表:填充_E7A3">
							<fill>
								<xsl:if test="表:填充_E7A3[图:颜色_8004] or 表:填充_E7A3[图:图案_800A]">
									<patternFill>
										<xsl:if test="表:填充_E7A3[图:颜色_8004 and 图:颜色_8004!='']">
											<xsl:attribute name="patternType">
												<xsl:value-of select="'solid'"/>
											</xsl:attribute>
										</xsl:if>
                    <!--图案填充-->
                    <!--Not Finished-->
                    <!--Marked by LDM in 2011/01/10-->
										<xsl:if test="表:填充_E7A3[图:图案_800A]">
											<xsl:attribute name="patternType">
                        <xsl:call-template name="PatternTypeTransfer_U2O"/>
											</xsl:attribute>
										</xsl:if>

                    <!--颜色填充-->
										<xsl:if test="表:填充_E7A3/图:颜色_8004">
											<xsl:variable name="fillColor" select="表:填充_E7A3/图:颜色_8004"/>
											<fgColor>
												<xsl:attribute name="rgb">
                          <!-- 20130409 update by xuzhenwei BUG_2811:互操作测试 oo-uof（编辑）-oo A16:A20单元格无填充，转换后为蓝色纯色填充 start -->
                          <xsl:choose>
                            <xsl:when test="$fillColor = 'auto'">
                              <xsl:value-of select="'FFFFFFFF'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="concat('FF',substring-after($fillColor,'#'))"/>
                            </xsl:otherwise>
                          </xsl:choose>
                          <!-- end -->
												</xsl:attribute>
											</fgColor>
											<bgColor indexed="64"/>
										</xsl:if>
										<xsl:if test="表:填充_E7A3/图:图案_800A">
                      <!--2014-4-28, update by Qihy， 图案填充不正确， start-->
                      
                        <!--<xsl:variable name="fgColor" select="表:填充_E7A3/图:图案_800A/@前景色_800B"/>
                        <xsl:variable name="bgColor" select="表:填充_E7A3/图:图案_800A/@背景色_800C"/>
                        <fgColor>
                          <xsl:attribute name="rgb">
                            <xsl:value-of select="concat('FF',substring-after($fgColor,'#'))"/>
                          </xsl:attribute>
                        </fgColor>
                        <bgColor>
                          <xsl:attribute name="rgb">
                            <xsl:value-of select="concat('FF',substring-after($bgColor,'#'))"/>
                          </xsl:attribute>
                        </bgColor>-->
                        <xsl:if test="表:填充_E7A3/图:图案_800A[@前景色_800B]">
                          <fgColor>
                            <xsl:variable name="color2" select="表:填充_E7A3/图:图案_800A/@前景色_800B"/>
                            <xsl:attribute name="rgb">
                              <xsl:value-of select="concat('FF',translate($color2,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
                            </xsl:attribute>
                          </fgColor>
                        </xsl:if>
                        <xsl:if test="表:填充_E7A3/图:图案_800A[@背景色_800C]">
                          <bgColor>
                            <xsl:variable name="color3" select="表:填充_E7A3/图:图案_800A/@背景色_800C"/>
                            <xsl:if test="$color3='auto'">
                              <xsl:attribute name="auto">
                              <xsl:value-of select="1"/>
                              </xsl:attribute>
                            </xsl:if>
                              <xsl:if test="not($color3='auto')">
                                <xsl:attribute name="rgb">
                              <xsl:value-of select="concat('FF',translate($color3,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
                                </xsl:attribute>
                              </xsl:if>
                                
                          </bgColor>
                        </xsl:if>
                      </xsl:if>
                    <!--2014-4-28 end-->
									</patternFill>
								</xsl:if>

                <!--渐变颜色填充-->
								<xsl:if test="表:填充_E7A3[图:渐变_800D]">
									<xsl:variable name="startColor" select="concat('FF',substring-after(表:填充_E7A3/图:渐变_800D/@起始色_800E,'#'))"/>
									<xsl:variable name="endColor" select="concat('FF',substring-after(表:填充_E7A3/图:渐变_800D/@终止色_800F,'#'))"/>
									<xsl:if test="表:填充_E7A3/图:渐变_800D[@种子类型_8010='linear' or @种子类型_8010='radar']">
										<gradientFill>
											<xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013]">
												<xsl:if test="表:填充_E7A3/图:渐变_800D/@渐变方向_8013='90'">
													<xsl:attribute name="degree">
														<xsl:value-of select="'0'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013='45']">
													<xsl:attribute name="degree">
														<xsl:value-of select="'45'"/>
													</xsl:attribute>
												</xsl:if>
                        <xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013='45']">
                          <xsl:attribute name="degree">
                            <xsl:value-of select="'50'"/>
                          </xsl:attribute>
                        </xsl:if>
												<xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013='315']">
													<xsl:attribute name="degree">
														<xsl:value-of select="'135'"/>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013='135']">
													<xsl:attribute name="degree">
														<xsl:value-of select="'315'"/>
													</xsl:attribute>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:填充_E7A3/图:渐变_800D[not(@渐变方向_8013) or @渐变方向_8013='180' or @渐变方向_8013='0']">
												<xsl:attribute name="degree">
													<xsl:value-of select="'90'"/>
												</xsl:attribute>
											</xsl:if>
											<stop>
												<xsl:attribute name="position">
													<xsl:value-of select="'0'"/>
												</xsl:attribute>
												<color>
													<xsl:attribute name="rgb">
														<xsl:value-of select="$startColor"/>
													</xsl:attribute>
												</color>
											</stop>
											<stop>
												<xsl:attribute name="position">
													<xsl:value-of select="'1'"/>
												</xsl:attribute>
												<color>
													<xsl:attribute name="rgb">
														<xsl:value-of select="$endColor"/>
													</xsl:attribute>
												</color>
											</stop>
										</gradientFill>
									</xsl:if>
									<xsl:if test="表:填充_E7A3/图:渐变_800D[@渐变方向_8013='0' and  @种子X位置_8015='30']">
										<gradientFill>
											<xsl:attribute name="type">
												<xsl:value-of select="'path'"/>
											</xsl:attribute>
											<stop>
												<xsl:attribute name="position">
													<xsl:value-of select="'0'"/>
												</xsl:attribute>
												<color>
													<xsl:attribute name="rgb">
														<xsl:value-of select="$startColor"/>
													</xsl:attribute>
												</color>
											</stop>
											<stop>
												<xsl:attribute name="position">
													<xsl:value-of select="'1'"/>
												</xsl:attribute>
												<color>
													<xsl:attribute name="rgb">
														<xsl:value-of select="$endColor"/>
													</xsl:attribute>
												</color>
											</stop>
										</gradientFill>
									</xsl:if>
									<xsl:if test="表:填充_E7A3/图:渐变_800D[@种子类型_8010='square']">
                    <xsl:for-each select="表:填充_E7A3/图:渐变_800D">
                      <gradientFill>
                        <xsl:attribute name="type">
                          <xsl:value-of select="'path'"/>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="@种子X位置_8015='60' and @种子Y位置_8016='30'">
                            <xsl:attribute name="left">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="top">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="@种子X位置_8015='30' and @种子Y位置_8016='60'">
                            <xsl:attribute name="top">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="@种子X位置_8015='60' and @种子Y位置_8016='60'">
                            <xsl:attribute name="left">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="top">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="@种子X位置_8015='30' and @种子Y位置_8016='30'">
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:attribute name="left">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="top">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                          </xsl:otherwise>
                        </xsl:choose>
                        <stop>
                          <xsl:attribute name="position">
                            <xsl:value-of select="'0'"/>
                          </xsl:attribute>
                          <color>
                            <xsl:attribute name="rgb">
                              <xsl:value-of select="$startColor"/>
                            </xsl:attribute>
                          </color>
                        </stop>
                        <stop>
                          <xsl:attribute name="position">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                          <color>
                            <xsl:attribute name="rgb">
                              <xsl:value-of select="$endColor"/>
                            </xsl:attribute>
                          </color>
                        </stop>
                      </gradientFill>
                    </xsl:for-each>
									</xsl:if>
                  <!--添加渐变填充： 角部辐射+中心辐射。李杨2012-3-16-->
                  <xsl:if test ="表:填充_E7A3/图:渐变_800D[@种子类型_8010='rectangle']">
                    <xsl:for-each select ="表:填充_E7A3/图:渐变_800D">
                      <gradientFill>
                        <xsl:attribute name ="type">path</xsl:attribute>
                        <xsl:choose >
                          <xsl:when test ="@渐变方向_8013='0' and @种子X位置_8015='0' and @种子Y位置_8016='100'">
                            <xsl:attribute name="top">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test ="@渐变方向_8013='0' and @种子X位置_8015='100' and @种子Y位置_8016='0'">
                            <xsl:attribute name="left">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test ="@渐变方向_8013='0' and @种子X位置_8015='100' and @种子Y位置_8016='100'">
                            <xsl:attribute name="left">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="top">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test ="(@渐变方向_8013='315' or @渐变方向_8013='0') and @种子X位置_8015='50' and @种子Y位置_8016='50'">
                            <xsl:attribute name="left">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="right">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="top">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                            <xsl:attribute name="bottom">
                              <xsl:value-of select="'0.5'"/>
                            </xsl:attribute>
                          </xsl:when>
                        </xsl:choose>
                        <stop>
                          <xsl:attribute name="position">
                            <xsl:value-of select="'0'"/>
                          </xsl:attribute>
                          <color>
                            <xsl:attribute name="rgb">
                              <xsl:value-of select="$startColor"/>
                            </xsl:attribute>
                          </color>
                        </stop>
                        <stop>
                          <xsl:attribute name="position">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                          <color>
                            <xsl:attribute name="rgb">
                              <xsl:value-of select="$endColor"/>
                            </xsl:attribute>
                          </color>
                        </stop>
                      </gradientFill>
                    </xsl:for-each>
                  </xsl:if>
                  <xsl:if test ="表:填充_E7A3/图:渐变_800D[@种子类型_8010='axial']">
                    <xsl:for-each select ="表:填充_E7A3/图:渐变_800D[@种子类型_8010='axial']">
                      <gradientFill>
                        <xsl:attribute name ="degree">
                          <xsl:if test ="@渐变方向_8013='0'">
                            <xsl:value-of select ="'90'"/>
                          </xsl:if>
                          <xsl:if test ="@渐变方向_8013='315'">
                            <xsl:value-of select ="'135'"/>
                          </xsl:if>
                        </xsl:attribute>
                        <!--<xsl:if test ="@渐变方向_8013='0' and @种子X位置_8015='100' and @种子Y位置_8016='100'">-->
                          <stop>
                            <xsl:attribute name="position">
                              <xsl:value-of select="'0'"/>
                            </xsl:attribute>
                            <color>
                              <xsl:attribute name="rgb">
                                <xsl:value-of select="$startColor"/>
                              </xsl:attribute>
                            </color>
                          </stop>
                          <!--<stop>
                            <xsl:attribute name ="positon">
                              <xsl:value-of select ="'0.5'"/>
                            </xsl:attribute>
                            <color>
                              <xsl:attribute name ="rgb">
                                <xsl:value-of select ="$endColor"/>
                              </xsl:attribute>
                            </color>
                          </stop>-->
                          <stop>
                            <xsl:attribute name="position">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                            <color>
                              <xsl:attribute name="rgb">
                                <xsl:value-of select="$endColor"/>
                              </xsl:attribute>
                            </color>
                          </stop>
                        <!--</xsl:if>-->
                      </gradientFill>
                    </xsl:for-each>
                  </xsl:if>
								</xsl:if>
							</fill>
						</xsl:if>
						<xsl:if test="not(表:填充_E7A3)">
							<fill>
								<patternFill patternType="none"/>
							</fill>
						</xsl:if>
					</xsl:for-each>
				</fills>
        
        <!--Border-->
        <!--Modified by LDM in 2011/01/10-->
				<borders>
          <xsl:attribute name="count">
						<xsl:value-of select="$CellStyNum"/>
					</xsl:attribute>
					<xsl:for-each select="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916">
            <!--
						<xsl:if test="./表:边框_4133">
              <xsl:call-template name="Border_Style"/>
						</xsl:if>
            -->
            <xsl:if test="./表:边框_4133">
            <border>
              <xsl:if test="表:边框_4133[uof:对角线2_C618]">
                <xsl:attribute name="diagonalUp">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="表:边框_4133[uof:对角线1_C617]">
                <xsl:attribute name="diagonalDown">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              
              <xsl:if test="表:边框_4133/uof:左_C613">
                <left>
                  <xsl:attribute name="style">
                    <xsl:call-template name="BorderStyleTransfer">
                      <xsl:with-param name="lineType" select="表:边框_4133/uof:左_C613/@线型_C60D"/>
                      <xsl:with-param name="dashType" select="表:边框_4133/uof:左_C613/@虚实_C60E"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <xsl:if test="(表:边框_4133/uof:左_C613/@颜色_C611) and (表:边框_4133/uof:左_C613/@颜色_C611 != 'auto')">
                    <xsl:variable name="borderColor" select="表:边框_4133/uof:左_C613/@颜色_C611"/>
                    <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                    <color>
                      <xsl:attribute name="rgb">
                        <xsl:value-of select="$color"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test ="表:边框_4133/uof:左_C613/@颜色_C611='auto' or 表:边框_4133/uof:左_C613/@颜色_C611=''">
                    <color>
                      <xsl:attribute name ="rgb">
                        <xsl:value-of select ="'FF808080'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test="表:边框_4133/uof:左_C613[not(@颜色_C611)]">
                    <color>
                      <xsl:attribute name="auto">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                </left>
              </xsl:if>
              <xsl:if test="表:边框_4133/uof:右_C615">
                <right>
                  <xsl:attribute name="style">
                    <xsl:call-template name="BorderStyleTransfer">
                      <xsl:with-param name="lineType" select="表:边框_4133/uof:右_C615/@线型_C60D"/>
                      <xsl:with-param name="dashType" select="表:边框_4133/uof:右_C615/@虚实_C60E"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <xsl:if test="(表:边框_4133/uof:右_C615/@颜色_C611) and (表:边框_4133/uof:右_C615/@颜色_C611 != 'auto')">
                    <xsl:variable name="borderColor" select="表:边框_4133/uof:右_C615/@颜色_C611"/>
                    <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                    <color>
                      <xsl:attribute name="rgb">
                        <xsl:value-of select="$color"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test ="(表:边框_4133/uof:右_C615/@颜色_C611='auto') or (表:边框_4133/uof:右_C615/@颜色_C611='')">
                    <color>
                      <xsl:attribute name ="rgb">
                        <xsl:value-of select ="'FF808080'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test="表:边框_4133/uof:右_C615[not(@颜色_C611)]">
                    <color>
                      <xsl:attribute name="auto">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                </right>
              </xsl:if>
              <xsl:if test="表:边框_4133/uof:上_C614">
                <top>
                  <xsl:attribute name="style">
                    <xsl:call-template name="BorderStyleTransfer">
                      <xsl:with-param name="lineType" select="表:边框_4133/uof:上_C614/@线型_C60D"/>
                      <xsl:with-param name="dashType" select="表:边框_4133/uof:上_C614/@虚实_C60E"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <xsl:if test="表:边框_4133/uof:上_C614/@颜色_C611 and 表:边框_4133/uof:上_C614/@颜色_C611 != 'auto'">
                    <xsl:variable name="borderColor" select="表:边框_4133/uof:上_C614/@颜色_C611"/>
                    <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                    <color>
                      <xsl:attribute name="rgb">
                        <xsl:value-of select="$color"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test ="表:边框_4133/uof:上_C614/@颜色_C611='auto' or 表:边框_4133/uof:上_C614/@颜色_C611=''">
                    <color>
                      <xsl:attribute name ="rgb">
                        <xsl:value-of select ="'FF808080'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test="表:边框_4133/uof:上_C614[not(@颜色_C611)]">
                    <color>
                      <xsl:attribute name="auto">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                </top>
              </xsl:if>
              <xsl:if test="表:边框_4133/uof:下_C616">
                <bottom>
                  <xsl:attribute name="style">
                    <xsl:call-template name="BorderStyleTransfer">
                      <xsl:with-param name="lineType" select="表:边框_4133/uof:下_C616/@线型_C60D"/>
                      <xsl:with-param name="dashType" select="表:边框_4133/uof:下_C616/@虚实_C60E"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <xsl:if test="表:边框_4133/uof:下_C616/@颜色_C611 and 表:边框_4133/uof:下_C616/@颜色_C611 !='auto'">
                    <xsl:variable name="borderColor" select="表:边框_4133/uof:下_C616/@颜色_C611"/>
                    <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                    <color>
                      <xsl:attribute name="rgb">
                        <xsl:value-of select="$color"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test ="表:边框_4133/uof:下_C616/@颜色_C611='auto' or 表:边框_4133/uof:下_C616/@颜色_C611=''">
                    <color>
                      <xsl:attribute name ="rgb">
                        <xsl:value-of select ="'FF808080'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                  <xsl:if test="表:边框_4133/uof:下_C616[not(@颜色_C611)]">
                    <color>
                      <xsl:attribute name="auto">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </color>
                  </xsl:if>
                </bottom>
              </xsl:if>
              
		<!--20130116 gaoyuwei bug 2673 单元格边框丢失 start-->
		<!--对角线的样式 OOXML和UOF实现不同 UOF可以左对角和右对角不同 而OOXML中两边必须相同-->
			<!--有左对角无右对角-->	
				<xsl:if test ="表:边框_4133/uof:对角线1_C617 and not(表:边框_4133/uof:对角线2_C618)">
						<diagonal>
							<xsl:attribute name="style">
								<xsl:call-template name="BorderStyleTransfer">
									<xsl:with-param name="lineType" select="表:边框_4133/uof:对角线1_C617/@线型_C60D"/>
									<xsl:with-param name="dashType" select="表:边框_4133/uof:对角线1_C617/@虚实_C60E"/>
								</xsl:call-template>
							</xsl:attribute>
							<xsl:if test="表:边框_4133/uof:对角线1_C617/@颜色_C611 and 表:边框_4133/uof:对角线1_C617/@颜色_C611 !='auto'">
								<xsl:variable name="borderColor" select="表:边框_4133/uof:对角线1_C617/@颜色_C611"/>
								<xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
								<color>
									<xsl:attribute name="rgb">
										<xsl:value-of select="$color"/>
									</xsl:attribute>
								</color>
							</xsl:if>
							<xsl:if test ="表:边框_4133/uof:对角线1_C617/@颜色_C611='auto' or 表:边框_4133/uof:对角线1_C617/@颜色_C611=''">
								<color>
									<xsl:attribute name ="rgb">
										<xsl:value-of select ="'FF000000'"/>
									</xsl:attribute>
								</color>
							</xsl:if>
							<xsl:if test="表:边框_4133/uof:对角线1_C617[not(@颜色_C611)]">
								<color>
									<xsl:attribute name="auto">
										<xsl:value-of select="'1'"/>
									</xsl:attribute>
								</color>
							</xsl:if>
						</diagonal>
					</xsl:if>
				<!--有右对角无左对角-->
				    <xsl:if test ="not(表:边框_4133/uof:对角线1_C617) and 表:边框_4133/uof:对角线2_C618">
					<diagonal>
						<xsl:attribute name="style">
							<xsl:call-template name="BorderStyleTransfer">
								<xsl:with-param name="lineType" select="表:边框_4133/uof:对角线2_C618/@线型_C60D"/>
								<xsl:with-param name="dashType" select="表:边框_4133/uof:对角线2_C618/@虚实_C60E"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="表:边框_4133/uof:对角线2_C618/@颜色_C611 and 表:边框_4133/uof:对角线2_C618/@颜色_C611 !='auto'">
							<xsl:variable name="borderColor" select="表:边框_4133/uof:对角线2_C618/@颜色_C611"/>
							<xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
							<color>
								<xsl:attribute name="rgb">
									<xsl:value-of select="$color"/>
								</xsl:attribute>
							</color>
						</xsl:if>
						<xsl:if test ="表:边框_4133/uof:对角线2_C618/@颜色_C611='auto' or 表:边框_4133/uof:对角线2_C618/@颜色_C611=''">
							<color>
								<xsl:attribute name ="rgb">
									<xsl:value-of select ="'FF000000'"/>
								</xsl:attribute>
							</color>
						</xsl:if>
						<xsl:if test="表:边框_4133/uof:对角线2_C618[not(@颜色_C611)]">
							<color>
								<xsl:attribute name="auto">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</color>
						</xsl:if>
					</diagonal>
				</xsl:if>
				<!--既有右对角又有左对角-->
				<xsl:if test ="表:边框_4133/uof:对角线1_C617 and 表:边框_4133/uof:对角线2_C618">
					<diagonal>
						<xsl:attribute name="style">
							<xsl:call-template name="BorderStyleTransfer">
								<xsl:with-param name="lineType" select="表:边框_4133/uof:对角线2_C618/@线型_C60D"/>
								<xsl:with-param name="dashType" select="表:边框_4133/uof:对角线2_C618/@虚实_C60E"/>
							</xsl:call-template>
						</xsl:attribute>
						<xsl:if test="表:边框_4133/uof:对角线2_C618/@颜色_C611 and 表:边框_4133/uof:对角线2_C618/@颜色_C611 !='auto'">
							<xsl:variable name="borderColor" select="表:边框_4133/uof:对角线2_C618/@颜色_C611"/>
							<xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
							<color>
								<xsl:attribute name="rgb">
									<xsl:value-of select="$color"/>
								</xsl:attribute>
							</color>
						</xsl:if>
						<xsl:if test ="表:边框_4133/uof:对角线2_C618/@颜色_C611='auto' or 表:边框_4133/uof:对角线2_C618/@颜色_C611=''">
							<color>
								<xsl:attribute name ="rgb">
									<xsl:value-of select ="'FF000000'"/>
								</xsl:attribute>
							</color>
						</xsl:if>
						<xsl:if test="表:边框_4133/uof:对角线2_C618[not(@颜色_C611)]">
							<color>
								<xsl:attribute name="auto">
									<xsl:value-of select="'1'"/>
								</xsl:attribute>
							</color>
						</xsl:if>
					</diagonal>
					</xsl:if>
				
						<!--end-->
            </border>
            </xsl:if>
						<xsl:if test="not(./表:边框_4133)">
							<border >
								<left/>
								<right/>
								<top/>
								<bottom/>
								<diagonal/>
							</border>
						</xsl:if>
					</xsl:for-each>
				</borders>
				<!--cellStyleXfs-->
				<cellStyleXfs count="1">
					<xf numFmtId="0" fontId="0" fillId="0" borderId="0">
						<alignment vertical="center"/>
					</xf>
				</cellStyleXfs>
				<!--cellXfs-->
				<cellXfs>
					<xsl:call-template name="CellXfs">
						<xsl:with-param name="CS" select="$CellStyNum"/>
					</xsl:call-template>
				</cellXfs>
				<!--cellStyles-->
				<cellStyles count="1">
					<cellStyle name="常规" xfId="0" builtinId="0"/>
				</cellStyles>
				<dxfs>
				<!--20130408 update gaoyuwei bug_2665 格式化条件丢失 start-->
					<xsl:if test="表:电子表格文档_E826/表:工作表集/表:单工作表/表:单条件格式化 or .//表:筛选_E80F[表:条件_E811/表:颜色_E818]">
						<xsl:attribute name="count">
              
              <!--2014-3-27, update by Qihy, bug3151 单元格式样不正确  start-->
							<!--<xsl:value-of select="number(count(表:电子表格文档_E826/表:工作表集/表:单工作表/表:单条件格式化)) + number(count(.//表:筛选_E80F[表:条件_E811/表:颜色_E818]))"/>-->
              <xsl:value-of select="number(count(表:电子表格文档_E826/表:工作表集/表:单工作表/表:单条件格式化/规则:条件格式化_B629/规则:条件_B62B)) + number(count(.//表:筛选_E80F[表:条件_E811/表:颜色_E818]))"/>
						</xsl:attribute>
						<xsl:for-each select="表:电子表格文档_E826/表:工作表集/表:单工作表/表:单条件格式化/规则:条件格式化_B629/规则:条件_B62B">
            <!--<xsl:for-each select="式样:式样集_9915/式样:单元格式样_9916[contains(@标识符_E7AC, 'tjgsh_')]">-->

				<xsl:if test ="not(规则:格式_B62D/@式样引用_B62E)">
				<!--end-->

                <dxf>
                  <fill>
                    <patternFill>
                      <bgColor rgb="FF00B0F0"/>
                    </patternFill>
                  </fill>
                </dxf>
              </xsl:if>
              <xsl:if test="规则:格式_B62D/@式样引用_B62E">
                <xsl:variable name="styaa" select="规则:格式_B62D/@式样引用_B62E"/>
							<!--sty:CFSTYLE_9-->
							<dxf>
								<xsl:for-each select="ancestor::uof:UOF/式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[@名称_E7AD=$styaa]">
									<xsl:if test="表:字体格式_E7A7">
										<font>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[@中文字体引用_412A]">
												<xsl:variable name="cid" select="表:字体格式_E7A7/字:字体_4128/@中文字体引用_412A"/>
												<name>
													<xsl:variable name="font">
														<xsl:value-of select="ancestor::式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$cid]/@名称_9903"/>
													</xsl:variable>
													<xsl:attribute name="val">
														<xsl:choose>
															<xsl:when test="contains($font,'永中')">
																<xsl:value-of select="translate($font,'永中','')"/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select="$font"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
												</name>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@中文字体引用_412A)]">
												<name val="宋体"/>
											</xsl:if>
											<charset>
												<xsl:attribute name="val">
													<xsl:value-of select="'134'"/>
												</xsl:attribute>
											</charset>
											<xsl:if test="表:字体格式_E7A7[字:是否粗体_4130]">
												<xsl:if test="表:字体格式_E7A7/字:是否粗体_4130='true'">
													<b/>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7[字:是否斜体_4131]">
												<xsl:if test="表:字体格式_E7A7/字:是否斜体_4131='true'">
													<i/>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7[字:删除线_4135]">
												<xsl:if test="表:字体格式_E7A7/字:删除线_4135 != 'none'">
													<strike/>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7[字:是否空心_413E]">
												<xsl:if test="表:字体格式_E7A7/字:是否空心_413E='true'">
													<outline/>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7[字:是否阴影_4140]">
												<xsl:if test="表:字体格式_E7A7/字:是否阴影_4140='true'">
													<shadow/>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F]">
												<xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F='auto']">
													<color theme="1"/>
												</xsl:if>
												<xsl:if test="表:字体格式_E7A7/字:字体_4128[@颜色_412F!='auto']">
													<xsl:variable name="color" select="表:字体格式_E7A7/字:字体_4128/@颜色_412F"/>
													<color>
														<xsl:attribute name="rgb">
															<xsl:value-of select="concat('FF',translate($color,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
														</xsl:attribute>
													</color>
												</xsl:if>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@颜色_412F)]">
												<color theme="1"/>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[@字号_412D]">
												<sz>
													<xsl:attribute name="val">
														<xsl:variable name="mm">
															<xsl:value-of select="表:字体格式_E7A7/字:字体_4128/@字号_412D"/>
														</xsl:variable>
														<xsl:value-of select="substring-before($mm,'.')"/>
													</xsl:attribute>
												</sz>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:字体_4128[not(@字号_412D)]">
												<sz val="11"/>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7/字:下划线_4136[@线型_4137!='none']">
												<u>
													<xsl:variable name="sty" select="表:字体格式_E7A7/字:下划线_4136/@线型_4137"/>
													<xsl:attribute name="val">
														<xsl:choose>
															<xsl:when test="$sty='none'">
																<xsl:value-of select="'none'"/>
															</xsl:when>
															<xsl:when test="$sty='single'">
																<xsl:value-of select="'single'"/>
															</xsl:when>
															<xsl:when test="$sty='double'">
																<xsl:value-of select="'double'"/>
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select="'single'"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
												</u>
											</xsl:if>
											<xsl:if test="表:字体格式_E7A7[字:上下标类型_4143]">
												<xsl:if test="表:字体格式_E7A7/字:上下标类型_4143='sub'">
													<vertAlign val="subscript"/>
												</xsl:if>
												<xsl:if test="表:字体格式_E7A7/字:上下标类型_4143='sup'">
													<vertAlign val="superscript"/>
												</xsl:if>
											</xsl:if>
										</font>
									</xsl:if>
									
                  <!--2014-5-29,add by Qihy，单元格格式化不正确， start-->
                  <xsl:if test="表:数字格式_E7A9">
                    <numFmt>
                      <xsl:attribute name="numFmtId">
                        <xsl:value-of select="position()"/>
                      </xsl:attribute>
                      <xsl:attribute name="formatCode">
                        <xsl:value-of select="表:数字格式_E7A9/@格式码_E73F"/>
                      </xsl:attribute>
                    </numFmt>
                  </xsl:if>
                  <!--2014-5-29 end-->
                  
                  <xsl:if test="表:填充_E7A3">										<fill>
											<xsl:if test="表:填充_E7A3[图:颜色_8004] or 表:填充_E7A3[图:图案_800A]">
												<patternFill>
													<xsl:attribute name="patternType">
														<xsl:choose>
															<xsl:when test="表:填充_E7A3/图:颜色_8004">
																<xsl:value-of select="'solid'"/>
															</xsl:when>
															<xsl:when test="表:填充_E7A3/图:图案_800A">
																<!--xsl:choose><-异常子节点<xsl:call-template name="TU-sty"/>></xsl:choose-->
                                
                                <!--2014-3-26, add by Qihy, 增加图案填充的情况-->
                                <xsl:call-template name="PatternTypeTransfer_U2O"/>
                                <!--2014-3-26 end-->
                                
															</xsl:when>
															<xsl:otherwise>
																<xsl:value-of select="'solid'"/>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
													<xsl:if test="表:填充_E7A3/图:颜色_8004">
														<xsl:variable name="color" select="表:填充_E7A3/图:颜色_8004"/>
                            <xsl:if test="$color != ''">
                              <bgColor>
                                <xsl:attribute name="rgb">
                                  <xsl:value-of select="concat('FF',translate($color,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
                                </xsl:attribute>
                              </bgColor>
                            </xsl:if>
														<!--bgColor indexed="64"/-->
													</xsl:if>
													<xsl:if test="表:填充_E7A3/图:图案_800A">
														<xsl:variable name="color2" select="表:填充_E7A3/图:图案_800A/@前景色_800B"/>
														<xsl:variable name="color3" select="表:填充_E7A3/图:图案_800A/@背景色_800C"/>
														<fgColor>
															<xsl:attribute name="rgb">
																<xsl:value-of select="concat('FF',translate($color2,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
															</xsl:attribute>
														</fgColor>
														<bgColor>
															<xsl:attribute name="rgb">
																<xsl:value-of select="concat('FF',translate($color3,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
															</xsl:attribute>
														</bgColor>
													</xsl:if>
												</patternFill>
											</xsl:if>
										</fill>
									</xsl:if>
									<xsl:if test="表:边框_4133">
                    <xsl:call-template name="Border_Style"/>
									</xsl:if>
								</xsl:for-each>
							</dxf>
              </xsl:if>
						</xsl:for-each>

            <!--<xsl:if test="/uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825/表:筛选集_E83A/表:筛选_E80F[表:条件_E811/表:颜色]">-->
              <xsl:for-each select="/uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825/表:筛选集_E83A/表:筛选_E80F[表:条件_E811/表:颜色_E818]">
                <dxf>
                  <fill>
                    <patternFill patternType="solid">
                      <fgColor>
                        <xsl:attribute name="rgb">
                          <xsl:value-of select="concat('ff',translate(./表:条件_E811/表:颜色_E818,'#',''))"/>
                        </xsl:attribute>
                      </fgColor>
                      <bgColor>
                        <xsl:attribute name="rgb">
                          <xsl:value-of select="concat('ff',translate(./表:条件_E811/表:颜色_E818,'#',''))"/>
                        </xsl:attribute>
                      </bgColor>
                    </patternFill>
                  </fill>
                </dxf>
              </xsl:for-each>
            <!--</xsl:if>-->
          </xsl:if>
        </dxfs>
				<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
			</xsl:if>
		</styleSheet>
	</xsl:template>

  <xsl:template name="Fill_commom">
    <fill>
      <xsl:if test="表:填充_E7A3[图:颜色_8004] or 表:填充_E7A3[图:图案_800A]">
        <patternFill>
          <xsl:attribute name="patternType">
            <xsl:choose>
              <xsl:when test="表:填充_E7A3/图:颜色_8004">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <xsl:when test="表:填充_E7A3/图:图案_800A">
                
                <!--2014-4-28, add by Qihy, 图案填充不正确， start-->
                <xsl:call-template name="PatternTypeTransfer_U2O"/>
                <!--xsl:choose><-异常子节点<xsl:call-template name="TU-sty"/>></xsl:choose-->
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'solid'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:if test="表:填充_E7A3/图:颜色_8004">
            <xsl:variable name="color" select="表:填充_E7A3/图:颜色_8004"/>
            <bgColor>
              <xsl:attribute name="rgb">
                <xsl:value-of select="concat('FF',translate($color,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
              </xsl:attribute>
            </bgColor>
            <!--bgColor indexed="64"/-->
          </xsl:if>
          <xsl:if test="表:填充_E7A3/图:图案_800A">
            <xsl:if test="表:填充_E7A3/图:图案_800A[@前景色_800B]">
            <fgColor>
              <xsl:variable name="color2" select="表:填充_E7A3/图:图案_800A/@前景色_800B"/>
              <xsl:attribute name="rgb">
                <xsl:value-of select="concat('FF',translate($color2,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
              </xsl:attribute>
            </fgColor>
            </xsl:if>
            <xsl:if test="表:填充_E7A3/图:图案_800A[@背景色_800C]">
            <bgColor>
              <xsl:variable name="color3" select="表:填充_E7A3/图:图案_800A/@背景色_800C"/>
              <xsl:attribute name="rgb">
                <xsl:value-of select="concat('FF',translate($color3,'abcdefghijklmnopqrstuvwxyz#','ABCDEFGHIJKLMNOPQRSTUVWXYZ'))"/>
              </xsl:attribute>
            </bgColor>
            </xsl:if>
          </xsl:if>
          <!--2014-4-28 end-->
          
        </patternFill>
      </xsl:if>
    </fill>
  </xsl:template>
  <!--BorderStyle-->
  <!--Modified by LDM in 2011/01/10-->
  <xsl:template name="Border_Style">
  
  	  <!--gaoyuwei 2636条件格式化格式不正确 UOF-OOXML start-->
	  <border xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		  <!--end-->

      <!--2014-4-27,update by Qihy, 添加uof颜色取值为auto的情况， start-->
      <xsl:if test="表:边框_4133[uof:对角线2_C618]">
        <xsl:attribute name="diagonalUp">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="表:边框_4133[uof:对角线1_C617]">
        <xsl:attribute name="diagonalDown">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="表:边框_4133/uof:左_C613">
        <left>
          <xsl:attribute name="style">
            <xsl:call-template name="BorderStyleTransfer">
              <xsl:with-param name="lineType" select="表:边框_4133/uof:左_C613/@线型_C60D"/>
              <xsl:with-param name="dashType" select="表:边框_4133/uof:左_C613/@虚实_C60E"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test="表:边框_4133/uof:左_C613/@颜色_C611">
            <xsl:variable name="borderColor" select="表:边框_4133/uof:左_C613/@颜色_C611"/>
            <color>
              <xsl:choose>
                <xsl:when test="$borderColor = 'auto' or $borderColor = ''">
                  <xsl:attribute name="auto">
                    <xsl:value-of select="1"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                  <xsl:attribute name="rgb">
                    <xsl:value-of select="$color"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </color>
          </xsl:if>
          <xsl:if test="表:边框_4133/uof:左_C613[not(@颜色_C611)]">
            <color>
              <xsl:attribute name="auto">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </color>
          </xsl:if>
        </left>
      </xsl:if>
      <xsl:if test="表:边框_4133/uof:右_C615">
        <right>
          <xsl:attribute name="style">
            <xsl:call-template name="BorderStyleTransfer">
              <xsl:with-param name="lineType" select="表:边框_4133/uof:右_C615/@线型_C60D"/>
              <xsl:with-param name="dashType" select="表:边框_4133/uof:右_C615/@虚实_C60E"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test="表:边框_4133/uof:右_C615/@颜色_C611">
            <xsl:variable name="borderColor" select="表:边框_4133/uof:右_C615/@颜色_C611"/>
            <color>
              <xsl:choose>
                <xsl:when test="$borderColor = 'auto' or $borderColor = ''">
                  <xsl:attribute name="auto">
                    <xsl:value-of select="1"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                  <xsl:attribute name="rgb">
                    <xsl:value-of select="$color"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </color>
          </xsl:if>
          <xsl:if test="表:边框_4133/uof:右_C615[not(@颜色_C611)]">
            <color>
              <xsl:attribute name="auto">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </color>
          </xsl:if>
        </right>
      </xsl:if>
      <xsl:if test="表:边框_4133/uof:上_C614">
        <top>
          <xsl:attribute name="style">
            <xsl:call-template name="BorderStyleTransfer">
              <xsl:with-param name="lineType" select="表:边框_4133/uof:上_C614/@线型_C60D"/>
              <xsl:with-param name="dashType" select="表:边框_4133/uof:上_C614/@虚实_C60E"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test="表:边框_4133/uof:上_C614/@颜色_C611">
            <xsl:variable name="borderColor" select="表:边框_4133/uof:上_C614/@颜色_C611"/>
            <color>
              <xsl:choose>
                <xsl:when test="$borderColor = 'auto' or $borderColor = ''">
                  <xsl:attribute name="auto">
                    <xsl:value-of select="1"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                  <xsl:attribute name="rgb">
                    <xsl:value-of select="$color"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </color>
          </xsl:if>
          <xsl:if test="表:边框_4133/uof:上_C614[not(@颜色_C611)]">
            <color>
              <xsl:attribute name="auto">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </color>
          </xsl:if>
        </top>
      </xsl:if>
      <xsl:if test="表:边框_4133/uof:下_C616">
        <bottom>
          <xsl:attribute name="style">
            <xsl:call-template name="BorderStyleTransfer">
              <xsl:with-param name="lineType" select="表:边框_4133/uof:下_C616/@线型_C60D"/>
              <xsl:with-param name="dashType" select="表:边框_4133/uof:下_C616/@虚实_C60E"/>
            </xsl:call-template>
          </xsl:attribute>
          <xsl:if test="表:边框_4133/uof:下_C616/@颜色_C611">
            <xsl:variable name="borderColor" select="表:边框_4133/uof:下_C616/@颜色_C611"/>
            <color>
              <xsl:choose>
                <xsl:when test="$borderColor = 'auto' or $borderColor = ''">
                  <xsl:attribute name="auto">
                    <xsl:value-of select="1"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="color" select="concat('FF',substring-after($borderColor,'#'))"/>
                  <xsl:attribute name="rgb">
                    <xsl:value-of select="$color"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </color>
          </xsl:if>
          <xsl:if test="表:边框_4133/uof:下_C616[not(@颜色_C611)]">
            <color>
              <xsl:attribute name="auto">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </color>
          </xsl:if>
        </bottom>
      </xsl:if>
      <!--2014-4-27 end-->

    </border>
  </xsl:template>
  
  <!--边框线类型转换模板-->
  <!--Modified by LDM in 2011/01/10-->
  <xsl:template name="BorderStyleTransfer">
    <xsl:param name="lineType"/>
    <xsl:param name="dashType"/>
    <xsl:if test="$lineType = 'single'">
      <xsl:choose>
        <xsl:when test="$dashType = ''">
          <xsl:value-of select="'thin'"/>
        </xsl:when>
        <xsl:when test="$dashType = 'dash-dot'">
          <xsl:value-of select="'dashDot'"/>
        </xsl:when>
        <xsl:when test="$dashType = 'dash-dot-dot'">
          <xsl:value-of select="'dashDotDot'"/>
        </xsl:when>
        <xsl:when test="$dashType = 'dash'">
          <xsl:value-of select="'dashed'"/>
        </xsl:when>
        <xsl:when test="$dashType = 'dotted'">
          <xsl:value-of select="'dotted'"/>
        </xsl:when>
        <xsl:when test="$dashType = 'square-dot'">
        <!--20130325 gaoyuwei bug 2796 uof-oo 集成 B10单元格边框转换后与差异文档描述不符-->
			<!--<xsl:value-of select="'mediumDashed'"/>-->

			<xsl:value-of select="'hair'"/>
			<!--end-->
        </xsl:when> 
        <xsl:otherwise>
          <xsl:value-of select="'thin'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="$lineType = 'double'">
      <xsl:value-of select="'double'"/>
    </xsl:if>
    <xsl:if test="$lineType = 'thin-thick'">
      <xsl:value-of select="'double'"/>
    </xsl:if>
    <xsl:if test="$lineType = 'thick-thin'">
      <xsl:value-of select="'double'"/>
    </xsl:if>
    <xsl:if test="$lineType = 'thick-between-thin'">
      <xsl:value-of select="'double'"/>
    </xsl:if>
    <!--zl start-->
    <xsl:if test="$lineType = 'medium'">
      <xsl:value-of select="'medium'"/>
    </xsl:if>
    <!--zl end-->
     <!--20130116 gaoyuwei bug 2673 单元格边框丢失 start-->
	  <xsl:if test="$lineType = 'thick'">
		  <xsl:choose>
			  <xsl:when test="$dashType = ''">
				  <xsl:value-of select="'thin'"/>
			  </xsl:when>
			  <xsl:when test="$dashType = 'square-dot'">
				  <xsl:value-of select="'mediumDashed'"/>
			  </xsl:when>
			  <xsl:when test="$dashType = 'dotted'">
				  <xsl:value-of select="'dotted'"/>
			  </xsl:when>
		  </xsl:choose>
	 </xsl:if>
	  <!--end-->
	  
  </xsl:template>
  
	<xsl:template name="CellXfs">
		<xsl:param name="CS"/>
		<xsl:attribute name="count">
			<xsl:value-of select="$CS"/>
		</xsl:attribute>
    
    <!--2014-6-9, update by Qihy, 默认字体不正确，将fontId取值由0改为1 start-->
		<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
			<alignment vertical="center"/>
		</xf>
    <!--2014-6-9 end-->
    
		<xsl:for-each select="式样:式样集_990B/式样:单元格式样集_9915/式样:单元格式样_9916[not(@标识符_E7AC='DEFSTYLE')]">
			<xsl:variable name="pos" select="position()"/>
			<xf xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
				<xsl:attribute name="numFmtId">
					<xsl:value-of select="$pos"/>
				</xsl:attribute>
				<xsl:if test="表:字体格式_E7A7">
					<xsl:attribute name="fontId">
            <xsl:value-of select="$pos"/>
					</xsl:attribute>
					<xsl:attribute name="applyFont">
						<xsl:value-of select="'1'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="表:填充_E7A3">
					<xsl:attribute name="fillId">
            <xsl:value-of select="$pos"/>
					</xsl:attribute>
					<xsl:attribute name="applyFill">
						<xsl:value-of select="'1'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="not(表:填充_E7A3)">
					<xsl:attribute name="fillId">
						<xsl:value-of select="'0'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="表:边框_4133 and 表:边框_4133/*">
					<xsl:attribute name="borderId">
						<xsl:value-of select="$pos"/>
					</xsl:attribute>
					<xsl:attribute name="applyBorder">
						<xsl:value-of select="'1'"/>
					</xsl:attribute>
				</xsl:if>
        <xsl:if test ="表:边框_4133 and not(表:边框_4133/*)">
          <xsl:attribute name="borderId">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:if>
				<xsl:if test="not(表:边框_4133)">
					<xsl:attribute name="borderId">
						<xsl:value-of select="'0'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:attribute name="xfId">
					<xsl:value-of select="'0'"/>
				</xsl:attribute>
				<xsl:if test="表:对齐格式_E7A8">
					<xsl:attribute name="applyAlignment">
						<xsl:value-of select="'1'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="表:对齐格式_E7A8">
					<alignment>
						<xsl:if test="表:对齐格式_E7A8[表:水平对齐方式_E700]">
							<xsl:attribute name="horizontal">
								<xsl:variable name="sty" select="表:对齐格式_E7A8/表:水平对齐方式_E700"/>
								<xsl:choose>
									<xsl:when test="$sty ='center across selection'">
										<xsl:value-of select="'centerContinuous'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="$sty"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="表:对齐格式_E7A8[表:垂直对齐方式_E701]">
							<xsl:attribute name="vertical">
								<xsl:value-of select="表:对齐格式_E7A8/表:垂直对齐方式_E701"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="表:对齐格式_E7A8[表:缩进_E702]">
							<xsl:attribute name="indent">
								<xsl:value-of select="表:对齐格式_E7A8/表:缩进_E702"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="表:对齐格式_E7A8[表:文字旋转角度_E704]">
							<xsl:attribute name="textRotation">
								<xsl:variable name="dre" select="表:对齐格式_E7A8/表:文字旋转角度_E704"/>
								<xsl:choose>
									<xsl:when test="$dre &gt; 0">
										<xsl:value-of select="$dre"/>
									</xsl:when>
									<xsl:when test="$dre='0'">
										<xsl:value-of select="$dre"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="90 - $dre"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="表:对齐格式_E7A8[表:是否自动换行_E705]">
							<xsl:attribute name="wrapText">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
						</xsl:if>
						<xsl:if test="表:对齐格式_E7A8[表:是否缩小字体填充_E706]">
							<xsl:attribute name="shrinkToFit">
								<xsl:value-of select="'1'"/>
							</xsl:attribute>
						</xsl:if>
					</alignment>
				</xsl:if>
				<xsl:if test="not(表:对齐格式_E7A8)">
					<alignment vertical="center"/>
				</xsl:if>
			</xf>
		</xsl:for-each>
	</xsl:template>
  <!--PatternType-->
  <!--UOF To OOXML-->
  <!--Modified by LDM in 2011/01/11-->
  
  <!--2014-4-28, update by Qihy, 修改图案填充不正确问题，部分类型对应不正确 start-->
  <xsl:template name="PatternTypeTransfer_U2O">
    <xsl:choose>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn001']">
        <xsl:value-of select="'gray125'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn002']">
        <xsl:value-of select="'gray125'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn003']">
        <xsl:value-of select="'gray0625'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn004']">
        <xsl:value-of select="'lightGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn005']">
        <xsl:value-of select="'lightTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn006']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn007']">
        <xsl:value-of select="'mediumGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn008']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn009']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn010']">
        <xsl:value-of select="'darkTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn011']">
        <xsl:value-of select="'darkTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn012']">
        <xsl:value-of select="'darkTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn013']">
        <xsl:value-of select="'lightDown'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn014']">
        <xsl:value-of select="'lightUp'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn015']">
        <xsl:value-of select="'darkDown'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn016']">
        <xsl:value-of select="'lightUp'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn017']">
        <xsl:value-of select="'lightDown'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn018']">
        <xsl:value-of select="'lightUp'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn019']">
        <xsl:value-of select="'lightVertical'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn020']">
        <xsl:value-of select="'lightHorizontal'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn021']">
        <xsl:value-of select="'lightVertical'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn022']">
        <xsl:value-of select="'darkHorizontal'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn023']">
        <xsl:value-of select="'darkVertical'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn024']">
        <xsl:value-of select="'darkHorizontal'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn025']">
        <xsl:value-of select="'lightDown'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn026']">
        <xsl:value-of select="'lightUp'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn027']">
        <xsl:value-of select="'lightHorizontal'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn028']">
        <xsl:value-of select="'lightVertical'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn029']">
        <xsl:value-of select="'gray125'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn030']">
        <xsl:value-of select="'lightTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn031']">
        <xsl:value-of select="'lightGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn032']">
        <xsl:value-of select="'gray125'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn033']">
        <xsl:value-of select="'lightUp'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn034']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn035']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn036']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn037']">
        <xsl:value-of select="'lightGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn038']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn039']">
        <xsl:value-of select="'gray125'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn040']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn041']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn042']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn043']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn044']">
        <xsl:value-of select="'lightGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn045']">
        <xsl:value-of select="'darkGrid'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn046']">
        <xsl:value-of select="'darkTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn047']">
        <xsl:value-of select="'lightTrellis'"/>
      </xsl:when>
      <xsl:when test="表:填充_E7A3/图:图案_800A[@类型_8008='ptn048']">
        <xsl:value-of select="'darkGray'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--2014-4-28 end-->
	<xsl:template name="TU-sty">
		<!--异常子节点

		<xsl:when test="表:填充/图:图案[@图:类型='ptn017']">
			<xsl:value-of select="'darkDown'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn030']">
			<xsl:value-of select="'darkGray'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn008']">
			<xsl:value-of select="'mediumGray'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn004']">
			<xsl:value-of select="'lightGray'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn024']">
			<xsl:value-of select="'darkHorizontal'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn018']">
			<xsl:value-of select="'darkUp'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn027']">
			<xsl:value-of select="'darkTrellis'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn020']">
			<xsl:value-of select="'lightHorizontal'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn013']">
			<xsl:value-of select="'lightDown'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn014']">
			<xsl:value-of select="'lightUp'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn043']">
			<xsl:value-of select="'lightGrid'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn047']">
			<xsl:value-of select="'lightTrellis'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn002']">
			<xsl:value-of select="'gray125'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn001']">
			<xsl:value-of select="'gray0625'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn045']">
			<xsl:value-of select="'darkGrid'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn023']">
			<xsl:value-of select="'darkVertical'"/>
		</xsl:when>
		<xsl:when test="表:填充/图:图案[@图:类型='ptn019']">
			<xsl:value-of select="'lightVertical'"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="solid"/>
		</xsl:otherwise>
-->
	</xsl:template>

  <!--Not Checked-->
  <!--Marked by LDM in 2011/01/10-->
  <!--
								<xsl:if test="表:字体格式/字:字体[@字:中文字体引用 and @字:西文字体引用]">
									<xsl:variable name="esFontRef" select="表:字体格式/字:字体/@字:中文字体引用"/>
									<xsl:variable name="wsFontRef" select="表:字体格式/字:字体/@字:西文字体引用"/>
									<name>
										<xsl:variable name="esFontName">
											<xsl:value-of select="ancestor::uof:式样集/uof:字体集/uof:字体声明[@uof:标识符=$esFontRef]/@uof:名称"/>
										</xsl:variable>
										<xsl:variable name="wsFontName">
											<xsl:value-of select="ancestor::uof:式样集/uof:字体集/uof:字体声明[@uof:标识符=$wsFontRef]/@uof:名称"/>
										</xsl:variable>
										<xsl:if test="$font!='永中宋体' and $font1='Times New Roman'">
											<xsl:attribute name="val">
												<xsl:choose>
													<xsl:when test="contains($font,'永中')">
														<xsl:value-of select="translate($font,'永中','')"/>
													</xsl:when>
													<xsl:otherwise>
														<xsl:value-of select="$font"/>
													</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
										</xsl:if>
                    <xsl:if test="$font='永中宋体' and $font1='Times New Roman'">
                      <xsl:variable name="styleid">
                        <xsl:value-of select="@表:标识符"/>
                      </xsl:variable>
                        <xsl:attribute name="val">
                            <xsl:variable name="eaorlatin">
                            <xsl:value-of select="following::node()/表:单元格[@表:式样引用=$styleid]/表:数据/字:句/字:文本串"/>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="substring($eaorlatin,1,1)='1' or substring($eaorlatin,1,1)='2' or substring($eaorlatin,1,1)='3' or substring($eaorlatin,1,1)='4' or substring($eaorlatin,1,1)='5' or substring($eaorlatin,1,1)='6' or substring($eaorlatin,1,1)='7' or substring($eaorlatin,1,1)='8' or substring($eaorlatin,1,1)='9' or substring($eaorlatin,1,1)='0' or substring($eaorlatin,1,1)='a' or substring($eaorlatin,1,1)='b' or substring($eaorlatin,1,1)='c' or substring($eaorlatin,1,1)='d' or substring($eaorlatin,1,1)='e' or substring($eaorlatin,1,1)='f' or substring($eaorlatin,1,1)='g' or substring($eaorlatin,1,1)='h' or substring($eaorlatin,1,1)='i' or substring($eaorlatin,1,1)='j' or substring($eaorlatin,1,1)='k' or substring($eaorlatin,1,1)='l' or substring($eaorlatin,1,1)='m' or substring($eaorlatin,1,1)='n' or substring($eaorlatin,1,1)='o' or substring($eaorlatin,1,1)='p' or substring($eaorlatin,1,1)='q' or substring($eaorlatin,1,1)='r' or substring($eaorlatin,1,1)='s' or substring($eaorlatin,1,1)='t' or substring($eaorlatin,1,1)='u' or substring($eaorlatin,1,1)='v' or substring($eaorlatin,1,1)='w' or substring($eaorlatin,1,1)='x' or substring($eaorlatin,1,1)='y' or substring($eaorlatin,1,1)='z' or substring($eaorlatin,1,1)='A' or substring($eaorlatin,1,1)='B'
                                    or substring($eaorlatin,1,1)='C' or substring($eaorlatin,1,1)='D' or substring($eaorlatin,1,1)='E' or substring($eaorlatin,1,1)='F' or substring($eaorlatin,1,1)='G' or substring($eaorlatin,1,1)='H' or substring($eaorlatin,1,1)='I' or substring($eaorlatin,1,1)='J' or substring($eaorlatin,1,1)='K'
                                    or substring($eaorlatin,1,1)='L' or substring($eaorlatin,1,1)='M' or substring($eaorlatin,1,1)='N' or substring($eaorlatin,1,1)='O' or substring($eaorlatin,1,1)='P' or substring($eaorlatin,1,1)='Q' or substring($eaorlatin,1,1)='R' or substring($eaorlatin,1,1)='S' or substring($eaorlatin,1,1)='T' or substring($eaorlatin,1,1)='U'
                                    or substring($eaorlatin,1,1)='V' or substring($eaorlatin,1,1)='W' or substring($eaorlatin,1,1)='X' or substring($eaorlatin,1,1)='Y' or substring($eaorlatin,1,1)='Z'">
                              <xsl:value-of select="'Times New Roman'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'永中宋体'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
										</xsl:if>
										<xsl:if test="$font='永中宋体' and $font1!='Times New Roman'">
											<xsl:attribute name="val">
												<xsl:value-of select="$font1"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test="$font!='永中宋体' and $font1!='Times New Roman'">
											<xsl:attribute name="val">
												<xsl:value-of select="$font1"/>
											</xsl:attribute>
										</xsl:if>
										<xsl:if test="$font=''or $font1=''">
											<xsl:attribute name="val">
												<xsl:value-of select="'宋体'"/>
											</xsl:attribute>
										</xsl:if>
									</name>
								</xsl:if>
                -->
  <!--Obsoluted-->
  <!--Marked by LDM in 2011/01/10-->
  <!--
  <xsl:template name="bordersty">
    <xsl:param name="sty"/>
    <xsl:choose>
      <xsl:when test="$sty='dash-dot'">
        <xsl:value-of select="'dashDot'"/>
      </xsl:when>
      <xsl:when test="$sty='dash-dot-dot'">
        <xsl:value-of select="'dashDotDot'"/>
      </xsl:when>
      <xsl:when test="$sty='dash'">
        <xsl:value-of select="'dashed'"/>
      </xsl:when>
      <xsl:when test="$sty='dotted'">
        <xsl:value-of select="'dotted'"/>
      </xsl:when>
      <xsl:when test="$sty='dotted-heavy'">
        <xsl:value-of select="'mediumDashed'"/>
      </xsl:when>
      <xsl:when test="$sty='dash-dot-heavy'">
        <xsl:value-of select="'mediumDashDot'"/>
      </xsl:when>
      <xsl:when test="$sty='dash-dot-dot-heavy'">
        <xsl:value-of select="'mediumDashDotDot'"/>
      </xsl:when>
      <xsl:when test="$sty='dashed-heavy'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$sty='double'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$sty='none'">
        <xsl:value-of select="'none'"/>
      </xsl:when>
      <xsl:when test="$sty='thick'">
        <xsl:value-of select="'thick'"/>
      </xsl:when>
      <xsl:when test="$sty='single'">
        <xsl:value-of select="'thin'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'thin'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  -->
</xsl:stylesheet>
