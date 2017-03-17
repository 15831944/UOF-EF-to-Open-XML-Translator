<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph">
	<xsl:import href="PPr-commen.xsl"/>
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="a:p">
		<!--10.02.01 马有旭 -->
		<!--12.03.06 李娟-->
		<xsl:variable name="phtype" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
		<xsl:if test="($phtype!='ftr' and $phtype!='hdr') or not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph[@type]) or not(preceding-sibling::a:p)">
			<!--<字:段落 uof:locID="t0051" uof:attrList="标识符">-->
			<字:段落_416B>
				<xsl:attribute name="标识符_4169">
					<xsl:value-of select="translate(@id,':','')"/>
				</xsl:attribute>
				<!-- 10.01.28 马有旭 添加-->
				<xsl:if test="a:fld">
					<xsl:attribute name="序号">
						<xsl:for-each select="a:r|a:fld|a:br">
							<xsl:if test="name(.)='a:fld'">
								<xsl:value-of select="position()"/>
							</xsl:if>
						</xsl:for-each>
					</xsl:attribute>
				</xsl:if>
				<xsl:variable name="lvl" select="a:pPr/@lvl"/>
				<字:段落属性_419B>
				
					<xsl:if test="@styleID">
						<xsl:attribute name="式样引用_419C">
							<xsl:value-of select="@styleID"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="not(@styleID) and ancestor::p:sld">
						<xsl:attribute name="式样引用_419C">
							<xsl:variable name="pPrID">
								<xsl:if test="../a:lstStyle">
									<xsl:call-template name="lstStyle">
										<xsl:with-param name="lvl" select="a:pPr/@lvl"/>
									</xsl:call-template>
								</xsl:if>
							</xsl:variable>
							<xsl:if test="$pPrID">
								<xsl:value-of select="$pPrID"/>
							</xsl:if>
							<xsl:if test="$pPrID=''">
								<xsl:variable name="hID" select="../../@hID"/>
								<xsl:variable name="sID">
									<!--@lvl exist-->
									<xsl:if test="a:pPr/@lvl">
										<xsl:value-of select="//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl=$lvl]/@styleID"/>
									</xsl:if>
									<xsl:if test="not(a:pPr/@lvl)">
										<xsl:value-of select="//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl='0' or not(a:pPr/@lvl)]/@styleID"/>
									</xsl:if>
								</xsl:variable>
								<xsl:if test="$sID !=''">
									<xsl:value-of select="$sID"/>
								</xsl:if>
								<!--Master p with this lvl not exist;look for the style in the txStyles of Master-->
								<xsl:if test="$sID=''">
									<xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph">
										<xsl:variable name="pht" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
										<xsl:choose>
											<xsl:when test="$pht='title' or $pht='ctrTitle'">
												<xsl:call-template name="titleStyle">
                        </xsl:call-template>
											</xsl:when>
											<xsl:when test="$pht='body' or $pht='obj' or not($pht) or $pht='subTitle'">
												<xsl:call-template name="bodyStyle">
                        </xsl:call-template>
											</xsl:when>
											<xsl:otherwise>
												<xsl:call-template name="otherStyles"/>
											</xsl:otherwise>
										</xsl:choose>
									</xsl:if>
									<xsl:if test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
										<xsl:call-template name="defaultStyle"/>
									</xsl:if>
								</xsl:if>
							</xsl:if>
						</xsl:attribute>
					</xsl:if>

          <!-- 缩进、大纲级别在式样中都不起作用，所以将其内容之间添加到段落属性里 -->
          <!--2014-04-05, tangjiang, 缩进效果放在式样文件中不起作用，故为每个段落添加缩进;此处有待完善 -->
          <xsl:if test="not(./a:pPr/@indent | ./a:pPr/@marL | ./a:pPr/@marR) and ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx!='14'">
            <xsl:variable name="layoutID" select="ancestor::p:sld/@sldLayoutID"/>
            <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
            <xsl:choose>
              <xsl:when test="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/@marL |
                        //p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/@marR |
                        //p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/@indent">
                <xsl:variable name="lvlpPr" select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr"/>
                <xsl:variable name="marLValue" select="$lvlpPr/@marL"/>
                <xsl:variable name="marRValue" select="$lvlpPr/@marR"/>
                <xsl:variable name="indentValue" select="$lvlpPr/@indent"/>
                <字:缩进_411D>
                  <xsl:if test="$marLValue">
                    <字:左_410E>
                      <字:绝对_4107 >
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$marLValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:左_410E>
                  </xsl:if>
                  <xsl:if test="$marRValue">
                    <字:右_4110>
                      <字:绝对_4107>
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$marRValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:右_4110>
                  </xsl:if>
                  <xsl:if test="$indentValue">
                    <字:首行_4111>
                      <字:绝对_4107 >
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$indentValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:首行_4111>
                  </xsl:if>
                </字:缩进_411D>
              </xsl:when>
              <xsl:when test="$lvl">
                <xsl:choose>
                  <xsl:when test="$lvl = '0'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '1'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '2'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '3'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '4'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '5'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '6'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '7'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$lvl = '8'">
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:variable name="lvlpPrNode" select="/p:presentation/p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr"/>
                    <xsl:variable name="marLValue" select="$lvlpPrNode/@marL"/>
                    <xsl:variable name="marRValue" select="$lvlpPrNode/@marR"/>
                    <xsl:variable name="indentValue" select="$lvlpPrNode/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type ='subTitle'">
                <xsl:variable name="hID" select="ancestor::p:sp/@hID"/>
                <xsl:variable name="sldLayoutSp" select="//p:sp[@id=$hID]"/>
                <xsl:variable name="lvlpPr" select="$sldLayoutSp/p:txBody/a:lstStyle/a:lvl1pPr"/>
                <xsl:variable name="marLValue" select="$lvlpPr/@marL"/>
                <xsl:variable name="marRValue" select="$lvlpPr/@marR"/>
                <xsl:variable name="indentValue" select="$lvlpPr/@indent"/>
                <字:缩进_411D>
                  <xsl:if test="$marLValue">
                    <字:左_410E>
                      <字:绝对_4107 >
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$marLValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:左_410E>
                  </xsl:if>
                  <xsl:if test="$marRValue">
                    <字:右_4110>
                      <字:绝对_4107>
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$marRValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:右_4110>
                  </xsl:if>
                  <xsl:if test="$indentValue">
                    <字:首行_4111>
                      <字:绝对_4107 >
                        <xsl:attribute name="值_410F">
                          <xsl:value-of select="$indentValue div 12700"/>
                        </xsl:attribute>
                      </字:绝对_4107>
                    </字:首行_4111>
                  </xsl:if>
                </字:缩进_411D>
              </xsl:when>
              <xsl:when test="not(@styleID)">
                <xsl:variable name="hID" select="ancestor::p:sp/@hID"/>
                <xsl:variable name="styleID" select="//p:sp[@id=$hID]/p:txBody/a:p[a:pPr/@lvl='0']/@styleID"/>
                <xsl:variable name="bodypPr" select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr[@id= $styleID]"/>
                <xsl:variable name="titlepPr" select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr[@id= $styleID]"/>
                <xsl:choose>
                  <xsl:when test="$bodypPr">
                    <xsl:variable name="marLValue" select="$bodypPr/@marL"/>
                    <xsl:variable name="marRValue" select="$bodypPr/@marR"/>
                    <xsl:variable name="indentValue" select="$bodypPr/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                  <xsl:when test="$titlepPr">
                    <xsl:variable name="marLValue" select="$titlepPr/@marL"/>
                    <xsl:variable name="marRValue" select="$titlepPr/@marR"/>
                    <xsl:variable name="indentValue" select="$titlepPr/@indent"/>
                    <字:缩进_411D>
                      <xsl:if test="$marLValue">
                        <字:左_410E>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marLValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:左_410E>
                      </xsl:if>
                      <xsl:if test="$marRValue">
                        <字:右_4110>
                          <字:绝对_4107>
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$marRValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:右_4110>
                      </xsl:if>
                      <xsl:if test="$indentValue">
                        <字:首行_4111>
                          <字:绝对_4107 >
                            <xsl:attribute name="值_410F">
                              <xsl:value-of select="$indentValue div 12700"/>
                            </xsl:attribute>
                          </字:绝对_4107>
                        </字:首行_4111>
                      </xsl:if>
                    </字:缩进_411D>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
          <!--end 2014-04-05, tangjiang, 缩进效果放在式样文件中不起作用，故为每个段落添加缩进 -->

          <!--2014-04-07, tangjiang, 对齐方式在式样文件中不起作用，故为每个段落添加对齐方式;此处有待完善 -->
          <!--字:对齐_417D 水平对齐_421D="left"-->
          <xsl:if test="not(./a:pPr/@algn)">
            <xsl:choose>
              <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type ='subTitle' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type ='ctrTitle' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type ='title'">
                <xsl:variable name="hID" select="ancestor::p:sp/@hID"/>
                <xsl:variable name="sldLayoutSp" select="//p:sp[@id=$hID]"/>
                <xsl:variable name="lvlpPr" select="$sldLayoutSp/p:txBody/a:lstStyle/a:lvl1pPr"/>
                <xsl:variable name="algn" select="$lvlpPr/@algn"/>
                <xsl:choose>
                  <xsl:when test="$algn">
                    <字:对齐_417D>
                      <xsl:attribute name="水平对齐_421D">
                        <xsl:choose>
                          <xsl:when test="$algn='l'">left</xsl:when>
                          <xsl:when test="$algn='r'">right</xsl:when>
                          <xsl:when test="$algn='ctr'">center</xsl:when>
                          <xsl:when test="$algn='just' or @algn='justLow'">justified</xsl:when>
                          <xsl:when test="$algn='dist'or @algn='thaiDist'">distributed</xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                    </字:对齐_417D>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="ancestor::p:sp/p:txBody/a:p/a:pPr/@lvl='0'">
                <xsl:variable name="type" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
                <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@type='body' and ./p:nvSpPr/p:nvPr/p:ph/@idx='1']"/>
                <xsl:variable name="algn" select="$sp/p:txBody/a:lstStyle/a:lvl1pPr/@algn"/>
                <xsl:choose>
                  <xsl:when test="$algn">
                    <字:对齐_417D>
                      <xsl:attribute name="水平对齐_421D">
                        <xsl:choose>
                          <xsl:when test="$algn='l'">left</xsl:when>
                          <xsl:when test="$algn='r'">right</xsl:when>
                          <xsl:when test="$algn='ctr'">center</xsl:when>
                          <xsl:when test="$algn='just' or @algn='justLow'">justified</xsl:when>
                          <xsl:when test="$algn='dist'or @algn='thaiDist'">distributed</xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                    </字:对齐_417D>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx='1' and ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='body'">
                <xsl:variable name="hID" select="ancestor::p:sp/@hID"/>
                <xsl:variable name="sldLayoutSp" select="//p:sp[@id=$hID]"/>
                <xsl:variable name="lvlpPr" select="$sldLayoutSp/p:txBody/a:lstStyle/a:lvl1pPr"/>
                <xsl:variable name="layoutAlgn" select="$lvlpPr/@algn"/>

                <xsl:variable name="masterAlgn" select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/@algn"/>
                <xsl:choose>
                  <xsl:when test="$layoutAlgn">
                    <字:对齐_417D>
                      <xsl:attribute name="水平对齐_421D">
                        <xsl:choose>
                          <xsl:when test="$layoutAlgn='l'">left</xsl:when>
                          <xsl:when test="$layoutAlgn='r'">right</xsl:when>
                          <xsl:when test="$layoutAlgn='ctr'">center</xsl:when>
                          <xsl:when test="$layoutAlgn='just' or @layoutAlgn='justLow'">justified</xsl:when>
                          <xsl:when test="$layoutAlgn='dist'or @layoutAlgn='thaiDist'">distributed</xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                    </字:对齐_417D>
                  </xsl:when>
                  <xsl:when test="$masterAlgn">
                    <字:对齐_417D>
                      <xsl:attribute name="水平对齐_421D">
                        <xsl:choose>
                          <xsl:when test="$masterAlgn='l'">left</xsl:when>
                          <xsl:when test="$masterAlgn='r'">right</xsl:when>
                          <xsl:when test="$masterAlgn='ctr'">center</xsl:when>
                          <xsl:when test="$masterAlgn='just' or @masterAlgn='justLow'">justified</xsl:when>
                          <xsl:when test="$masterAlgn='dist'or @masterAlgn='thaiDist'">distributed</xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                    </字:对齐_417D>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:if>
          <!--end 2014-04-07, tangjiang, 对齐方式在式样文件中不起作用，故为每个段落添加对齐方式;此处有待完善 -->
          
          <!--2014-04-17, tangjiang, 段间距在式样文件中不起作用，故为每个段落添加段间距;此处有待完善 -->
          <xsl:if test="not(./a:pPr/a:spcBef | ./a:pPr/a:spcAft ) and (ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx)">
            <xsl:variable name="layoutID" select="ancestor::p:sld/@sldLayoutID"/>
            <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
            <xsl:choose>
              <xsl:when test="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:spcBef |
                        //p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:spcAft">
                <xsl:variable name="spcBefVal" select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:spcBef"/>
                <xsl:variable name="spcAftVal" select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:spcAft"/>

                <字:段间距_4180>
                  <xsl:if test="$spcBefVal">
                    <字:段前距_4181>
                      <!--字:自动_4182-->
                      <xsl:if test="$spcBefVal/a:spcPct">
                        <字:相对值_4184>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcBefVal/a:spcPct/@val div 100000"/>
                          <!--</xsl:attribute>-->
                        </字:相对值_4184>
                      </xsl:if>
                      <xsl:if test="$spcBefVal/a:spcPts">
                        <字:绝对值_4183>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcBefVal/a:spcPts/@val div 100"/>
                          <!--</xsl:attribute>-->
                        </字:绝对值_4183>
                      </xsl:if>
                      <!--自动 无对应-->
                    </字:段前距_4181>
                  </xsl:if>
                  <xsl:if test="$spcAftVal">
                    <字:段后距_4185>
                      <xsl:if test="$spcAftVal/a:spcPct">
                        <字:相对值_4184>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcAftVal/a:spcPct/@val div 100000"/>
                          <!--</xsl:attribute>-->
                        </字:相对值_4184>
                      </xsl:if>
                      <xsl:if test="$spcAftVal/a:spcPts">
                        <字:绝对值_4183>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcAftVal/a:spcPts/@val div 100"/>
                          <!--</xsl:attribute>-->
                        </字:绝对值_4183>
                      </xsl:if>
                      <!--自动 无对应-->
                    </字:段后距_4185>
                  </xsl:if>
                </字:段间距_4180>
              </xsl:when>
              <xsl:when test="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcBef | //p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcAft">
                <xsl:variable name="spcBefVal" select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcBef"/>
                <xsl:variable name="spcAftVal" select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:spcAft"/>

                <字:段间距_4180>
                  <xsl:if test="$spcBefVal">
                    <字:段前距_4181>
                      <!--字:自动_4182-->
                      <xsl:if test="$spcBefVal/a:spcPct">
                        <字:相对值_4184>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcBefVal/a:spcPct/@val div 100000"/>
                          <!--</xsl:attribute>-->
                        </字:相对值_4184>
                      </xsl:if>
                      <xsl:if test="$spcBefVal/a:spcPts">
                        <字:绝对值_4183>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcBefVal/a:spcPts/@val div 100"/>
                          <!--</xsl:attribute>-->
                        </字:绝对值_4183>
                      </xsl:if>
                      <!--自动 无对应-->
                    </字:段前距_4181>
                  </xsl:if>
                  <xsl:if test="$spcAftVal">
                    <字:段后距_4185>
                      <xsl:if test="$spcAftVal/a:spcPct">
                        <字:相对值_4184>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcAftVal/a:spcPct/@val div 100000"/>
                          <!--</xsl:attribute>-->
                        </字:相对值_4184>
                      </xsl:if>
                      <xsl:if test="$spcAftVal/a:spcPts">
                        <字:绝对值_4183>
                          <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="$spcAftVal/a:spcPts/@val div 100"/>
                          <!--</xsl:attribute>-->
                        </字:绝对值_4183>
                      </xsl:if>
                      <!--自动 无对应-->
                    </字:段后距_4185>
                  </xsl:if>
                </字:段间距_4180>
              </xsl:when>
              <xsl:otherwise>

              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
          <!--end 2014-04-17, tangjiang, 段间距在式样文件中不起作用，故为每个段落添加段间距;此处有待完善 -->
          
          <!--txStyles.xsl中定义的 PPr-commen-->
					<!--add by linyh 项目编号是图片的情况-->
					<xsl:for-each select="a:pPr">
						<xsl:call-template name="PPr-commen"/>
					</xsl:for-each>
        
					
					<xsl:if test="not(a:pPr)">
						
						<xsl:call-template name="PPr-commen"/>
					</xsl:if>
					
					<xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">
						<字:大纲级别_417C>
							<xsl:value-of select="a:pPr/@lvl"/>
							<!--修改母版排版案例 先注销这块 李娟 2012 04 06-->
							<!--<xsl:choose>
								--><!--
								
								<xsl:when test="a:pPr/@lvl='1'">1</xsl:when>
								<xsl:when test="name(.)='a:lvl2pPr'">2</xsl:when>
								<xsl:when test="name(.)='a:lvl3pPr'">3</xsl:when>
								<xsl:when test="name(.)='a:lvl4pPr'">4</xsl:when>
								<xsl:when test="name(.)='a:lvl5pPr'">5</xsl:when>
								<xsl:when test="name(.)='a:lvl6pPr'">6</xsl:when>
								<xsl:when test="name(.)='a:lvl7pPr'">7</xsl:when>
								<xsl:when test="name(.)='a:lvl8pPr'">8</xsl:when>
								<xsl:when test="name(.)='a:lvl9pPr'">9</xsl:when>
								<xsl:otherwise>0</xsl:otherwise>
							</xsl:choose>-->
						</字:大纲级别_417C>
					</xsl:if>
					<!--<xsl:call-template name="PPr-commen"/>-->
					<!--<xsl:apply-templates select="a:defRPr" mode="pPr-child"/>-->
				</字:段落属性_419B>
			   <xsl:for-each select="./a:r | ./a:br | ./a:fld">
					<xsl:apply-templates select=".">
            
            <!--2014-04-16, tangjiang, 修复超链接丢失 start -->
            <xsl:with-param name="position">
              <xsl:value-of select="position()"/>
            </xsl:with-param>
            <!--end 2014-04-16, tangjiang, 修复超链接丢失 -->
            
						<xsl:with-param name="phtype" select="$phtype"/>
					</xsl:apply-templates>
				</xsl:for-each>
			</字:段落_416B>
		</xsl:if>
	</xsl:template>
  <!-- 2011-4-22 罗文甜 修改日期时间和页号占位符的处理 -->
  <xsl:template match="a:fld">
    <xsl:choose>
      <xsl:when test="@type='slidenum'">
        <!--<字:句_419D>-->
		  <字:句_419D>
			<xsl:apply-templates select="a:rPr"/>
			  <字:文本串_415B>&lt;#&gt;</字:文本串_415B>
			  <!--<字:文本串_415B>&lt;#&gt;<字:文本串_415B>-->
        <!--</字:句_419D>-->
		  </字:句_419D>
      </xsl:when>
      <!--2011-4-22罗文甜修改日期时间bug-->
      <xsl:when test="contains(@type,'datetime')">
        <!--<字:句_419D>-->
        <字:句_419D>
          <xsl:apply-templates select="a:rPr"/>
          <!--<字:文本串_415B>-->
          <字:文本串_415B>
            <xsl:choose>
              <xsl:when test="./a:t">
                <xsl:value-of select="./a:t"/>
              </xsl:when>
              <xsl:otherwise>
                &lt;日期/时间&gt;
              </xsl:otherwise>
            </xsl:choose>
          </字:文本串_415B>
        </字:句_419D>
        <!--</字:句_419D>-->
      </xsl:when>
      <xsl:otherwise>
        <!--2011-3-28罗文甜，增加有关日期时间的域设置共12+13=25种-->
        <xsl:apply-templates select="." mode="field">
          <xsl:with-param name="position" select="position()"/>
        </xsl:apply-templates>
        <!--字:文本串 uof:locID="t0109" uof:attrList="标识符">&lt;日期/时间&gt;</字:文本串-->
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
	<!--2011-3-28罗文甜，增加域的设置共12+13=25种-->
  <xsl:template match="a:fld" mode="field">
    <xsl:param name="position"/>
    <xsl:variable name="pID" select="translate(../@id,':','')"/>
	  <字:域开始_419E>
    <!--<字:域开始 uof:locID="t0079" uof:attrList="类型 锁定" 字:锁定="false">-->
      <xsl:attribute name="类型_416E">
        <xsl:value-of select="@type"/>
      </xsl:attribute>
		  <xsl:attribute name="是否锁定_416F">false</xsl:attribute>
	  </字:域开始_419E>
    <!--</字:域开始>-->
	  <字:域代码_419F>
		  <字:段落_416B>
        <!--<xsl:attribute name="字:标识符">
          <xsl:value-of select="concat('fld_',$position,$pID)"/>
        </xsl:attribute>-->
        <xsl:for-each select="a:pPr">
            <xsl:call-template name="PPr-commen"/>
         
        </xsl:for-each>
		  
        <xsl:choose>
          <!--chinese-->
          <xsl:when test="@type='datetime1' and @id='{07875FEA-D690-4736-8A76-82A729C3C05C}'">
			  <字:句_419D>
				  <字:文本串_415B>DATE \@ "yyyy-M-d"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime2' and @id='{96BC9E2B-6B32-4D1B-BDCC-F50018547A65}'">
			  <字:句_419D>
				<字:文本串_415B>DATE \@ "yyyy年M月d日"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime3' and @id='{8BB06C08-29DE-407B-A438-1C7FF038DDE3}'">
			  <字:句_419D>
				  <字:文本串_415B>DATE \@ "yyyy年M月d日星期W"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime5' and @id='{9665A019-4947-41D6-9725-AD4AEDB7F53B}'">
			  <字:句_419D>
				<字:文本串_415B>DATE \@ "yyyy/M/d"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime6' and @id='{25F62B4B-0B15-4722-A740-B80BDE13B3A7}'">
			  <字:句_419D>
				  <字:文本串_415B>DATE \@ "yyyy年M月"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime7' and @id='{E0A96566-746A-4A30-9A4E-3786E5CCB122}'">
			  <字:句_419D>
				<字:文本串_415B>DATE \@ "yy.M.d"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime8' and @id='{7523B887-C219-4CF3-B9FF-32ECFA8B0133}'">
			  <字:句_419D>
				  <字:文本串_415B>DATE \@ "yyyy年M月d日"</字:文本串_415B>
				  </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime9' and @id='{1168CF14-B775-46C2-9930-D9F859AA02F4}'">
			  <字:句_419D>
				<字:文本串_415B>DATE \@ "yyyy年M月d日星期W"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime10' and @id='{508339BD-2B22-44FB-BA39-3BA87899AC6B}'">
			  <字:句_419D>
				  <字:文本串_415B>TIME \@ "HH时mm分"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime11' and @id='{D99B4EEC-0E7A-4824-A173-322870626312}'">
			  <字:句_419D>
				<字:文本串_415B>TIME \@ "HH时mm分ss秒"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime12' and @id='{8F11DB01-CAD1-45E6-955B-5DD3327F6386}'">
			  <字:句_419D>
				  <字:文本串_415B>TIME \@ "AMPMh时mm分"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime13' and @id='{CBFD99C6-F06A-4423-A82B-CEB801BEDD4A}'">
			  <字:句_419D>
				<字:文本串_415B>TIME \@ "AMPMh时m分"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <!--english-->
          <xsl:when test="@type='datetime1' and @id='{7B330999-4E9C-4010-8C46-70DA5A72AF27}'">
			  <字:句_419D>
				  <字:文本串_415B>DATE \@ "M/d/yyyy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime2' and @id='{DEC9F1D0-C326-433E-832C-464AF1F47893}'">
			  <字:句_419D>
				<字:文本串_415B>DATE \@ "dddd, MMMM dd, yyyy"</字:文本串_415B>
            </字:句_419D>
      </xsl:when>
          <xsl:when test="@type='datetime3' and @id='{1DDEB56B-E768-4798-BEC5-64372AB82BBA}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "d MMMM yyyy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime4' and @id='{DDDC536B-EF1F-47DC-AAF3-CCA2207CA5A6}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "MMMM d, yyyy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime5' and @id='{E3692678-25E9-46AF-9841-0CBA8D57A35C}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "dd-MMM-yy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime6' and @id='{C4A1ED76-F930-4115-97C8-3A834E14B9FE}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "MMMM yy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime7' and @id='{9AC57DB4-732E-4441-91A3-A18EDD453AB8}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "MMM-yy"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime8' and @id='{B85D4F64-3C8E-48B4-8E00-F11AFA59A441}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "M/d/yyyy h:mm AM/PM"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime9' and @id='{ED8D6149-125B-4C68-9197-4070E7C9D1DF}'">
            <字:句_419D>
              <字:文本串_415B>DATE \@ "M/d/yyyy h:mm:ss AM/PM"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime10' and @id='{0382C5FD-1CFE-493D-8F76-8086E8236E25}'">
            <字:句_419D>
              <字:文本串_415B>TIME \@ "HH:mm"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime11' and @id='{2BCDC95A-F615-4A6D-9A0D-C38301346329}'">
            <字:句_419D>
              <字:文本串_415B>TIME \@ "HH:mm:ss"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime12' and @id='{CE51008B-08D5-4258-85A5-F244CC987B57}'">
            <字:句_419D>
              <字:文本串_415B>TIME \@ "h:mm AM/PM"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
          <xsl:when test="@type='datetime13' and @id='{274F25F7-9462-443B-82A4-07E0FB7330FB}'">
            <字:句_419D>
              <字:文本串_415B>TIME \@ "h:mm:ss AM/PM"</字:文本串_415B>
            </字:句_419D>
          </xsl:when>
        </xsl:choose>
		  </字:段落_416B>
    </字:域代码_419F>
    <字:句_419D>
      <xsl:apply-templates select="a:rPr"/>
      <xsl:apply-templates select="a:t"/>
    </字:句_419D>
	  <字:域结束_41A0/>
  </xsl:template>

  
  <xsl:template match="a:r">
		<xsl:param name="position">
			<xsl:value-of select="position()"/>
		</xsl:param>
		<xsl:param name="phtype"/>
		
		<字:句_419D>
			<xsl:apply-templates select="a:rPr"/>
			<!--<xsl:choose>-->
				<xsl:if test="not(a:rPr/a:hlinkClick/@r:id='') or a:rPr/a:hlinkClick/@action">
					<xsl:if test="a:rPr/a:hlinkClick or a:rPr/a:hlinkMouseOver">
						
						<xsl:variable name="pID" select="translate(../@id,':','')"/>
						<字:区域开始_4165>
						<!--<字:区域开始 uof:locID="t0121" uof:attrList="标识符 名称 类型" 字:名称="hyperlink" 字:类型="hyperlink">-->
							<xsl:attribute name="标识符_4100">
								<xsl:value-of select="concat('hcS_',$position,$pID)"/>
							</xsl:attribute>
							<xsl:attribute name="名称_4166">hyperlink</xsl:attribute>
							<xsl:attribute name="类型_413B">hyperlink</xsl:attribute>
						</字:区域开始_4165>
						<!--</字:区域开始>-->
					</xsl:if>
					
				</xsl:if>
			<xsl:if test="a:t">
				<xsl:apply-templates select="a:t">
					<xsl:with-param name="phtype" select="$phtype"/>
				</xsl:apply-templates>
			</xsl:if>
			<xsl:choose>
				<xsl:when test="not(a:rPr/a:hlinkClick/@r:id='') or a:rPr/a:hlinkClick/@action">
					<xsl:if test="a:rPr/a:hlinkClick or a:rPr/a:hlinkMouseOver">
						<xsl:variable name="pID" select="translate(../@id,':','')"/>
						<字:区域结束_4167>
							<xsl:attribute name="标识符引用_4168">
								<xsl:value-of select="concat('hcS_',$position,$pID)"/>
							</xsl:attribute>
						</字:区域结束_4167>
					</xsl:if>
				</xsl:when>
				
			</xsl:choose>
		</字:句_419D>
   <!---15-03-06 liuyangyang 修改批注 注释代码 start-->
		<!--<xsl:if test="//p:cmLst/p:cm and not(ancestor::p:sldMaster)">
			
				--><!--<xsl:for-each select="//p:cmLst/p:cm">
					<字:句_419D>
						<字:区域开始_4165>
						<xsl:variable name="idx">
							<xsl:value-of select ="count(preceding-sibling::p:cm)"/>
						</xsl:variable>
						<xsl:attribute name="标识符_4100">
							<xsl:value-of select="concat('cmt_',$idx)"/>

						</xsl:attribute>
							<xsl:attribute name="类型_413B">annotation</xsl:attribute>
						</字:区域开始_4165>
						
							<字:区域结束_4167>
						<xsl:attribute name="标识符引用_4168">
							
								<xsl:variable name="idx">
									<xsl:value-of select ="count(preceding-sibling::p:cm)"/>
								</xsl:variable>
							<xsl:value-of select="concat('cmt_',$idx)"/>
						</xsl:attribute>
							</字:区域结束_4167>
					</字:句_419D>
						</xsl:for-each>--><!--
				
				</xsl:if>-->
    <!--end 15-03-06 liuyangyang 修改批注 注释代码 -->
	</xsl:template>
	<xsl:template match="a:rPr">
		<!--PPr-commen.xsl中定义-->
		<xsl:call-template name="rPr"/>
	</xsl:template>
	<xsl:template match="a:t">
		<!--字:文本串 uof:locID="t0109" uof:attrList="标识符">
      <xsl:value-of select="current()"/>
    </字:文本串-->
		<!--4月27日jjy改-->
		<!-- 10.02.01 马有旭 修改-->
		<xsl:param name="phtype"/>
		<xsl:variable name="str">
			<xsl:value-of select="current()"/>
		</xsl:variable>
		
		<xsl:call-template name="text">
			<xsl:with-param name="str" select="$str"/>
			<xsl:with-param name="phtype" select="$phtype"/>
		</xsl:call-template>
	</xsl:template>
	<xsl:template name="text">
		<xsl:param name="str"/>
		<xsl:param name="phtype"/>
		<xsl:choose>
			<xsl:when test="starts-with($str,'&#x20;')">
				<字:空格符_4161>
				<!--<字:空格符 uof:locID="t0126" uof:attrList="个数">-->
					<xsl:attribute name="个数_4162">
						<xsl:value-of select="number(1)"/>
					</xsl:attribute>
				<!--</字:空格符>-->
				</字:空格符_4161>
				<xsl:call-template name="text">
					<xsl:with-param name="str" select="substring-after($str,'&#x20;')"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test="contains($str,'&#x9;')">
				<xsl:choose>
					<xsl:when test="starts-with($str,'&#x9;')">
						<字:制表符_415E/>
						<xsl:call-template name="text">
							<xsl:with-param name="str" select="substring-after($str,'&#x9;')"/>
						</xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="text">
							<xsl:with-param name="str" select="substring-before($str,'&#x9;')"/>
						</xsl:call-template>
						<字:制表符_415E/>
						<xsl:call-template name="text">
							<xsl:with-param name="str" select="substring-after($str,'&#x9;')"/>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$str!=''">
					<字:文本串_415B>
            
            <!--2014-03-29, tangjiang, 修复页眉页脚内容转换错误 start -->
            <xsl:value-of select="$str"/>
            <!--<xsl:choose>
              <xsl:when test="$phtype='ftr'">
                <xsl:choose>
                  <xsl:when test="not(ancestor::p:handoutMaster) and not(ancestor::p:notesMaster)">
                    &lt;页脚&gt;
                  </xsl:when>
                  <xsl:when test="(ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@ftr='0' or not((ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf)">
                    <xsl:value-of select="$str"/>
                  </xsl:when>
                  <xsl:when test="(ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@ftr!='0' or not((ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@ftr)">
                    &lt;页脚&gt;
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="$phtype='hdr'">
                <xsl:choose>
                  <xsl:when test="not(ancestor::p:handoutMaster) and not(ancestor::p:notesMaster)"
                            >&lt;页眉&gt;</xsl:when>
                  <xsl:when test="(ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@hdr='0'
                            or not((ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf)">
                    <xsl:value-of select="$str"/>
                  </xsl:when>
                  <xsl:when test="(ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@hdr!='0' 
                        or not((ancestor::p:handoutMaster|ancestor::p:notesMaster)/p:hf/@hdr)"
                            >&lt;页眉&gt;</xsl:when>
                </xsl:choose>
              </xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$str"/>
							</xsl:otherwise>
						</xsl:choose>-->
            <!--end 2014-03-29, tangjiang, 修复页眉页脚内容转换错误 -->
					</字:文本串_415B>
          
				</xsl:if>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="a:br">
		<字:句_419D>
			<字:换行符_415F/>
		</字:句_419D>
	</xsl:template>
	<xsl:template name="lstStyle">
		<xsl:param name="lvl"/>
		<xsl:choose>
			<xsl:when test="$lvl='0' or $lvl='' or not($lvl)">
				<xsl:value-of select="../a:lstStyle/a:lvl1pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='1'">
				<xsl:value-of select="../a:lstStyle/a:lvl2pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='2'">
				<xsl:value-of select="../a:lstStyle/a:lvl3pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='3'">
				<xsl:value-of select="../a:lstStyle/a:lvl4pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='4'">
				<xsl:value-of select="../a:lstStyle/a:lvl5pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='5'">
				<xsl:value-of select="../a:lstStyle/a:lvl6pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='6'">
				<xsl:value-of select="../a:lstStyle/a:lvl7pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='7'">
				<xsl:value-of select="../a:lstStyle/a:lvl8pPr/@id"/>
			</xsl:when>
			<xsl:when test="$lvl='8'">
				<xsl:value-of select="../a:lstStyle/a:lvl9pPr/@id"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="otherStyles">
		<xsl:choose>
			<xsl:when test="a:pPr/@lvl='0'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl1pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='1'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl2pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='2'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl3pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='3'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl4pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='4'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl5pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='5'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl6pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='6'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl7pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='7'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl8pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='8'">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl9pPr/@id"/>
			</xsl:when>
			<xsl:when test="not(a:pPr/@lvl)">
				<xsl:value-of select="//p:txStyles/p:otherStyle/a:lvl1pPr/@id"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="titleStyle">
		<!--xsl:param name="titleStyle"/-->
		<xsl:choose>
			<xsl:when test="a:pPr/@lvl='0'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl1pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='1'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl2pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='2'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl3pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='3'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl4pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='4'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl5pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='5'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl6pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='6'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl7pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='7'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl8pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='8'">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl9pPr/@id"/>
			</xsl:when>
			<xsl:when test="not(a:pPr/@lvl)">
				<xsl:value-of select="//p:txStyles/p:titleStyle/a:lvl1pPr/@id"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="bodyStyle">
		<!--xsl:param name="titleStyle"/-->
		<xsl:choose>
			<xsl:when test="a:pPr/@lvl='0'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl1pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='1'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl2pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='2'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl3pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='3'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl4pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='4'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl5pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='5'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl6pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='6'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl7pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='7'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl8pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='8'">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl9pPr/@id"/>
			</xsl:when>
			<xsl:when test="not(a:pPr/@lvl)">
				<xsl:value-of select="//p:txStyles/p:bodyStyle/a:lvl1pPr/@id"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="defaultStyle">
		<!--xsl:param name="titleStyle"/-->
		<xsl:choose>
			<xsl:when test="a:pPr/@lvl='0'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl1pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='1'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl2pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='2'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl3pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='3'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl4pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='4'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl5pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='5'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl6pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='6'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl7pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='7'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl8pPr/@id"/>
			</xsl:when>
			<xsl:when test="a:pPr/@lvl='8'">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl9pPr/@id"/>
			</xsl:when>
			<xsl:when test="not(a:pPr/@lvl)">
				<xsl:value-of select="//p:defaultTextStyle/a:lvl1pPr/@id"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
</xsl:stylesheet>
