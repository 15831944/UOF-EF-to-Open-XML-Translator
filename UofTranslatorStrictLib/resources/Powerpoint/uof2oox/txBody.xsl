<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
				xmlns:pzip="urn:u2o:xmlns:post-processings:special"
				xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
				xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
	xmlns:规则="http://schemas.uof.org/cn/2009/rules"
	xmlns:式样="http://schemas.uof.org/cn/2009/styles"
				xmlns:图形="http://schemas.uof.org/cn/2009/graphics">
	<!--修改标签 对于 字。单元格 这块 待定 李娟 2012.01.07-->
	<xsl:import href="pPr.xsl"/>
	<!--	
	<xsl:import href ="region.xsl"/-->
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template name="txBody">
		<!--
    2010.3.16 增加文字表内容,命名空间不同  <a:txBody>与 <p:txBody>  
    
    -->
		<xsl:choose>
			<xsl:when test="name(.)='字:单元格_41BE'">
				<!--单元格是在文字表/字：行下定义的 李娟 2012.01.08-->
				<a:txBody>
					<xsl:call-template name="bodyPr"/>
					<xsl:for-each select="../演:lstStyle">
						<xsl:call-template name="lstStyle"/>
					</xsl:for-each>
					<xsl:for-each select="字:段落_416B">
						<xsl:call-template name="paragraph"/>
					</xsl:for-each>
				</a:txBody>
			</xsl:when>
			<xsl:otherwise>
				<p:txBody>

					<xsl:call-template name="bodyPr"/>

					<xsl:for-each select="../演:lstStyle">

						<xsl:call-template name="lstStyle"/>
					</xsl:for-each>
					<xsl:choose>
						<xsl:when test="图:内容_8043/字:段落_416B">
							<xsl:for-each select="图:内容_8043/字:段落_416B">
								<xsl:call-template name="paragraph"/>
								</xsl:for-each>
						</xsl:when>
						<xsl:when test="not(图:内容_8043/字:段落_416B)">
							<a:p/>
						</xsl:when>
					</xsl:choose>
				
				</p:txBody>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="lstStyle">
		<a:lstStyle>
			<xsl:for-each select="*">
				<xsl:if test="not(字:大纲级别_417C) and not(not(preceding-sibling::node()/字:大纲级别_417C))">
					<a:lvl1pPr>
						<xsl:call-template name="pPr"/>
					</a:lvl1pPr>
				</xsl:if>
				<xsl:if test="not(current()/字:大纲级别_417C = preceding-sibling::node()[字:大纲级别_417C]/字:大纲级别_417C)">
					<!--2011-2-21罗文甜，修改大纲级别bug,增加判断，垂直标题 -->
					<xsl:choose>
						<xsl:when test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='title' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='centertitle' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='vertical_title'">
							<xsl:choose>
								<xsl:when test="字:大纲级别_417C='0'">
									<a:lvl1pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl1pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='1'">
									<a:lvl2pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl2pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='2'">
									<a:lvl3pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl3pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='3'">
									<a:lvl4pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl4pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='4'">
									<a:lvl5pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl5pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='5'">
									<a:lvl6pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl6pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='6'">
									<a:lvl7pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl7pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='7'">
									<a:lvl8pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl8pPr>
								</xsl:when>
								<!--20110221罗文甜，uof中大纲级别是从0到8-->
								<xsl:when test="字:大纲级别_417C='8'">
									<a:lvl9pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl9pPr>
								</xsl:when>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<!--2011-2-21罗文甜，修改大纲级别bug -->
								<xsl:when test="字:大纲级别_417C='1'">
									<a:lvl1pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl1pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='2'">
									<a:lvl2pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl2pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='3'">
									<a:lvl3pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl3pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='4'">
									<a:lvl4pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl4pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='5'">
									<a:lvl5pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl5pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='6'">
									<a:lvl6pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl6pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='7'">
									<a:lvl7pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl7pPr>
								</xsl:when>
								<xsl:when test="字:大纲级别_417C='8'">
									<a:lvl8pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl8pPr>
								</xsl:when>
								<!--20110221罗文甜，uof中大纲级别是从0到8-->
								<xsl:when test="字:大纲级别_417C='9'">
									<a:lvl9pPr>
										<xsl:call-template name="pPr"/>
									</a:lvl9pPr>
								</xsl:when>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:for-each>
		</a:lstStyle>
	</xsl:template>
	<xsl:template name="paragraph">
		<a:p>
			<!--2011-3-28罗文甜，修改域代码-->
			<xsl:choose>
				<xsl:when test="字:域开始_419E and 字:域代码_419F">
					<xsl:for-each select="字:段落属性_419B">
						<a:pPr>
							<xsl:call-template name="pPr"/>
						</a:pPr>
					</xsl:for-each>
					<xsl:for-each select="字:域代码_419F/字:段落_416B">
						<xsl:call-template name="field">

              <!--start liuyin 20121231 修改对象数据集时间和符号丢失 -->
							<xsl:with-param name="string" select="concat(字:句_419D[2]/字:文本串_415B[2],字:句_419D[2]/字:文本串_415B[3])"/>
              <!--end liuyin 20121231 修改对象数据集时间和符号丢失 -->
              
						</xsl:call-template>
					</xsl:for-each>
				</xsl:when>
				<xsl:otherwise>
					<!--<xsl:if test="/uof:UOF/uof:演示文稿//uof:锚点_C644[uof:占位符_C626/@类型_C627='centertitle' or uof:占位符_C626/@类型_C627='subtitle']">-->
					<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627!='footer' and ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627!='header' and ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627!='date' and ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627!='number' or ancestor::uof:锚点_C644[not(uof:占位符_C626/@类型_C627)] ">
						<xsl:for-each select="字:段落属性_419B">
							<a:pPr>
								<xsl:call-template name="pPr"/>
							</a:pPr>
						</xsl:for-each>
						<xsl:for-each select="字:句_419D">
              
              <!--start liuyin 20130310 修改集成测试中，文本中转换后莫名出现一个空行-->
              <xsl:apply-templates select="./字:换行符_415F | ./字:文本串_415B | ./字:空格符_4161 | ./字:制表符_415E"/>
              
              <!--<xsl:if test="字:换行符_415F">
                <a:br>
                  <a:rPr lang="en-US" altLang="zh-CN" dirty="0" smtClean="0"/>
                </a:br>
              </xsl:if>
							<xsl:if test="字:文本串_415B or 字:空格符_4161 or 字:制表符_415E">
								<a:r>
									<xsl:apply-templates select="字:句属性_4158"/>

									<a:t>
										<xsl:for-each select="*">
											<xsl:choose>
												<xsl:when test="name(.)='字:空格符_4161'">
													<xsl:call-template name="blankSpace">
														<xsl:with-param name="num" select="@个数_4162"/>
													</xsl:call-template>
												</xsl:when>


												<xsl:when test="name(.)='字:制表符_415E'">
													<xsl:value-of select="'&#x9;'"/>
												</xsl:when>
												<xsl:when test="name(.)='字:文本串_415B'">
													<xsl:value-of select="."/>
												</xsl:when>
											</xsl:choose>
										</xsl:for-each>
									</a:t>
								</a:r>
							</xsl:if>-->
              <!--end liuyin 20130310 修改集成测试中，文本中转换后莫名出现一个空行-->
              
						</xsl:for-each>
           
					</xsl:if>
					<!--<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='header'">-->
					<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='header'">
						<!--xsl:if test="not(ancestor::uof:锚点/@uof:页眉) or ancestor::uof:锚点/@uof:页眉='true'"-->
						<xsl:for-each select="字:段落属性_419B">
							<a:pPr>
								<xsl:call-template name="pPr"/>
							</a:pPr>
						</xsl:for-each>
						<xsl:for-each select="字:句_419D">
							<xsl:if test="字:文本串_415B or 字:空格符_4161 or 字:制表符_415E">
								<a:r>
									<xsl:apply-templates select="字:句属性_4158"/>
									<a:t>
										<xsl:for-each select="*">
											<xsl:choose>
												<xsl:when test="name(.)='字:空格符_4161'">
													<xsl:call-template name="blankSpace">
														<xsl:with-param name="num" select="@个数_4162"/>
													</xsl:call-template>
												</xsl:when>

												<xsl:when test="name(.)='字:制表符_415E'">
													<xsl:value-of select="'&#x9;'"/>
												</xsl:when>
												<!--2011-1-10罗文甜：修改页眉页脚对应-->
												<xsl:when test="name(.)='字:文本串_415B'">
													<xsl:choose>
														<xsl:when test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:页眉_B64D">
															<xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:页眉_B64D"/>
														</xsl:when>

														<xsl:otherwise>
															<!--xsl:value-of select="'&lt;页眉&gt;'"/-->
															<xsl:value-of select="."/>
															<!--xsl:choose>
                              <xsl:when test="./text()='&lt;页脚&gt;'">
                                <xsl:value-of select="ancestor::uof:锚点/@uof:页脚内容"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="."/>
                              </xsl:otherwise>
                            </xsl:choose-->
														</xsl:otherwise>
													</xsl:choose>
												</xsl:when>
											</xsl:choose>
										</xsl:for-each>
									</a:t>
								</a:r>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
					<!--/xsl:if-->
					<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='footer'">

						<xsl:for-each select="字:段落属性_419B">
							<a:pPr>
								<xsl:call-template name="pPr"/>
							</a:pPr>
						</xsl:for-each>
						<xsl:for-each select="字:句_419D">
							<xsl:if test="字:文本串_415B or 字:空格符_4161 or 字:制表符_415E">
								<a:r>
									<xsl:apply-templates select="字:句属性_4158"/>
									<a:t>
										<xsl:for-each select="*">
											<xsl:choose>
												<xsl:when test="name(.)='字:空格符_4161'">
													<xsl:call-template name="blankSpace">
														<xsl:with-param name="num" select="@个数_4162"/>
													</xsl:call-template>
												</xsl:when>
												<xsl:when test="name(.)='字:制表符_415E'">
													<xsl:value-of select="'&#x9;'"/>
												</xsl:when>
												<!--2011-1-10罗文甜：修改页眉页脚对应-->
												<xsl:when test="name(.)='字:文本串_415B'">
													<xsl:choose>
														<xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='slide' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:页脚_B644">
															<xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:页脚_B644"/>
														</xsl:when>
														<xsl:when test="(ancestor::演:母版_6C0D/@类型_6BEA='handout' or ancestor::演:母版_6C0D/@类型_6BEA='notes') and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:页脚_B644">
															<xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:页脚_B644"/>
														</xsl:when>
														<xsl:otherwise>
															<!--xsl:value-of select="'&lt;页脚&gt;'"/-->
															<xsl:value-of select="."/>
															<!--xsl:choose>
                              <xsl:when test="./text()='&lt;页脚&gt;'">
                                <xsl:value-of select="ancestor::uof:锚点/@uof:页脚内容"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="."/>
                              </xsl:otherwise>
                            </xsl:choose-->
														</xsl:otherwise>
													</xsl:choose>
												</xsl:when>
											</xsl:choose>
										</xsl:for-each>
									</a:t>
								</a:r>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>
					<!--/xsl:if-->
					<!--luowentian 2011-4-11修改编号样式 -->
					<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627 ='number'">
						<xsl:for-each select="字:段落属性_419B">
							<a:pPr>
								<xsl:call-template name="pPr"/>
							</a:pPr>
						</xsl:for-each>
						<xsl:for-each select="字:句_419D">
							<xsl:if test="字:文本串_415B/text()='&lt;#&gt;' or 字:文本串_415B/text()='‹#›'">
								<!--2011-2-22罗文甜：修改页眉页脚对应-->
								<xsl:choose>
									<xsl:when test="((ancestor::演:母版_6C0D/@类型_6BEA='slide' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/@是否显示幻灯片编号_B64A='true') or 
                 ((ancestor::演:母版_6C0D/@类型_6BEA='handout' or ancestor::演:母版_6C0D/@类型_6BEA='notes') and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/@是否显示页码_B650='true'))">
										<a:fld>
											<!--2011-2-22罗文甜：修改页眉页脚对应-->
											<xsl:choose>
												<xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='slide'">
													<xsl:attribute name="id">{0C913308-F349-4B6D-A68A-DD1791B4A57B}</xsl:attribute>
												</xsl:when>
												<xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='handout' or ancestor::演:母版_6C0D/@类型_6BEA='notes'">
													<xsl:attribute name="id">{FC29FB5B-8407-4AF8-93CA-468E4D28B29A}</xsl:attribute>
												</xsl:when>
											</xsl:choose>
											<xsl:attribute name="type">slidenum</xsl:attribute>
											<xsl:apply-templates select="字:句属性_4158"/>
											<a:t>‹#›</a:t>
										</a:fld>
									</xsl:when>
									<xsl:otherwise>
										<xsl:if test="字:文本串_415B or 字:空格符_4161 or 字:制表符_415E">
											<a:r>
												<xsl:apply-templates select="字:句属性_4158"/>
												<a:t>
													<xsl:for-each select="*">
														<xsl:choose>
															<xsl:when test="name(.)='字:空格符_4161'">
																<xsl:call-template name="blankSpace">
																	<xsl:with-param name="num" select="@个数_4162"/>
																</xsl:call-template>
															</xsl:when>
															<xsl:when test="name(.)='字:制表符_415E'">
																<xsl:value-of select="'&#x9;'"/>
															</xsl:when>
															<xsl:when test="name(.)='字:文本串_415B'">
																<xsl:value-of select="."/>
															</xsl:when>
														</xsl:choose>
													</xsl:for-each>
												</a:t>
											</a:r>
										</xsl:if>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:if>
						</xsl:for-each>
					</xsl:if>

					<!--2011-1-11罗文甜：修改页眉页脚对应-->
					<xsl:if test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='date'">
            <xsl:for-each select="字:段落属性_419B">
							<a:pPr>
								<xsl:call-template name="pPr"/>
							</a:pPr>
						</xsl:for-each>
						<xsl:for-each select="字:句_419D">
							<xsl:choose>
								<xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='slide' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/@是否显示日期和时间_B647='true' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/@是否自动更新日期和时间_B649='true'">
									<xsl:if test="字:文本串_415B">
                    <a:fld>
                      <xsl:variable name="dateFormatStr" select="//规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:日期和时间字符串_B643"/>
                      <xsl:choose>
                        <!-- yyyy年M月d日星期Wh时mm分s秒  -->
                        <xsl:when test="contains($dateFormatStr,'星期') and contains($dateFormatStr,'时')">
                          <xsl:attribute name="id">{091D00B7-26A8-4F7D-A019-54B603A8017E}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime9'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- yyyy年M月d日星期W -->
                        <xsl:when test="contains($dateFormatStr,'星期')">
                          <xsl:attribute name="id">{C40BE7F6-1A14-440E-9BB5-AB73A721E0A0}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime3'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- yyyy年M月d日H时mm分 -->
                        <xsl:when test="contains($dateFormatStr,'时') and not(contains($dateFormatStr,'上午') or contains($dateFormatStr,'下午'))">
                          <xsl:attribute name="id">{589B3BD2-DB7A-4E6D-B895-04DD8F61B8FD}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime8'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- 上午/下午h时mm分ss秒 -->
                        <xsl:when test="contains($dateFormatStr,'时') and contains($dateFormatStr,'秒') and (contains($dateFormatStr,'上午') or contains($dateFormatStr,'下午'))">
                          <xsl:attribute name="id">{FA3DF4B1-ED33-49F3-9920-DE2E598A7D22}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime13'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- 上午/下午h时m分 -->
                        <xsl:when test="contains($dateFormatStr,'时') and not(contains($dateFormatStr,'秒')) and (contains($dateFormatStr,'上午') or contains($dateFormatStr,'下午'))">
                          <xsl:attribute name="id">{471AD6EA-DD4D-443F-B2B9-96FA4F31CF31}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime12'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- yyyy年M月d日 -->
                        <xsl:when test="contains($dateFormatStr,'日')">
                          <xsl:attribute name="id">{BB86A603-1F96-426C-8663-3F92B673318E}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime2'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- yyyy年M月 -->
                        <xsl:when test="contains($dateFormatStr,'月')">
                          <xsl:attribute name="id">{C2905042-C7BD-4005-B246-6BA1B30FC1F3}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime6'"/>
                          </xsl:attribute>
                        </xsl:when>
                        <!-- yyyy/M/d -->
                        <xsl:otherwise>
                          <xsl:attribute name="id">{6A7691D5-F326-4BCD-8387-1A0AE58275E5}</xsl:attribute>
                          <xsl:attribute name="type">
                            <xsl:value-of select="'datetime1'"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                      <xsl:apply-templates select="字:句属性_4158"/>
                      <a:t>
                        <xsl:value-of select="//规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:日期和时间字符串_B643"/>
                      </a:t>
                    </a:fld>
									</xsl:if>
								</xsl:when>
								<xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='notes' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/@是否显示日期和时间_B647='true' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/@是否自动更新日期和时间_B649='true'">
									<xsl:if test="字:文本串_415B">
										<a:fld>
                      <!-- 2011=3=20罗文甜，讲义和备注的日期时间转换有待完善 -->
											<xsl:attribute name="id">{34D8FDF5-13D8-4142-B988-3F7CAC4E5DCF}</xsl:attribute>
											<xsl:attribute name="type">datetimeFigureOut</xsl:attribute>
											<xsl:apply-templates select="字:句属性_4158"/>
											<a:t>
												<!--2011=3=20罗文甜，uof标准不完善，备注中的页眉页脚内容在软件中没有记录>
                        <xsl:value-of select="'&lt;日期/时间&gt-->
												<xsl:value-of select="."/>
											</a:t>
										</a:fld>
									</xsl:if>
								</xsl:when>
								<xsl:otherwise>
									<xsl:if test="字:文本串_415B or 字:空格符_4161 or 字:制表符_415E">
										<a:r>
											<xsl:apply-templates select="字:句属性_4158"/>
											<a:t>
												<xsl:for-each select="*">
													<xsl:choose>
														<xsl:when test="name(.)='字:空格符_4161'">
															<xsl:call-template name="blankSpace">
																<xsl:with-param name="num" select="@个数_4162"/>
															</xsl:call-template>
														</xsl:when>
														<xsl:when test="name(.)='字:制表符_415E'">
															<xsl:value-of select="'&#x9;'"/>
														</xsl:when>
														<xsl:when test="name(.)='字:文本串_415B'">
                              
															<xsl:choose>

                                <!--start liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->
                                <!--<xsl:when test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:幻灯片_B641/规则:日期和时间字符串_B643">-->
                                <!--<xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:幻灯片_B641/规则:日期和时间字符串_B643"/>-->
                                <xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='notes'">
                                  <xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:日期和时间字符串_B643"/>
                                </xsl:when>
                                <xsl:when test="ancestor::演:母版_6C0D/@类型_6BEA='slide'">
                                  <xsl:value-of select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:日期和时间字符串_B643"/>
                                </xsl:when>
                                <!--end liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->
                                
																<xsl:otherwise>
																	<!--xsl:value-of select="'&lt;日期/时间&gt;'"/-->
																	<xsl:value-of select="."/>
																</xsl:otherwise>
															</xsl:choose>
														</xsl:when>
													</xsl:choose>
												</xsl:for-each>
											</a:t>
										</a:r>
									</xsl:if>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
						<!--xsl:for-each select="字:句">
          <xsl:choose>
            <xsl:when test="not(字:文本串) or 字:文本串/text()!='&lt;日期/时间&gt;'">
              <xsl:apply-templates select="."/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:apply-templates select="." mode="date"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:for-each-->
					</xsl:if>
					<!--</xsl:for-each>-->
				</xsl:otherwise>
			</xsl:choose>
		</a:p>
	</xsl:template>
  
  <!--start liuyin 20130310 修改集成测试中，文本中转换后莫名出现一个空行-->
  <xsl:template match="字:换行符_415F">
    <a:br>
      <a:rPr lang="en-US" altLang="zh-CN" dirty="0" smtClean="0"/>
    </a:br>
  </xsl:template>
  <xsl:template match="字:文本串_415B">
    <a:r>
      <xsl:apply-templates select="../字:句属性_4158"/>
      <a:t>
          <xsl:value-of select="."/>
      </a:t>
    </a:r>
  </xsl:template>
  <xsl:template match="字:空格符_4161">
    <a:r>
      <xsl:apply-templates select="../字:句属性_4158"/>
      <a:t>
           <xsl:call-template name="blankSpace">
             <xsl:with-param name="num" select="@个数_4162"/>
           </xsl:call-template>
      </a:t>
    </a:r>
  </xsl:template>
  <xsl:template match="字:制表符_415E">
    <a:r>
      <xsl:apply-templates select="../字:句属性_4158"/>
      <a:t>
           <xsl:value-of select="'&#x9;'"/>
      </a:t>
    </a:r>
  </xsl:template>
  <!--end liuyin 20130310 修改集成测试中，文本中转换后莫名出现一个空行-->
  
	<!--2011-3-28罗文甜增加域代码25种-->
	<xsl:template name="field">
		<xsl:param name="string"/>
		<xsl:variable name="type" select="substring-before($string,'\@')"/>
    
    <!--start liuyin 20121231 修改对象数据集时间和符号丢失 -->
    <!--<xsl:variable name="datetime" select="substring-after($string,'\@ ')"/>-->
		<xsl:variable name="datetime" select="substring-after($string,'\@')"/>
    <!--end liuyin 20121231 修改对象数据集时间和符号丢失 -->
    
		<a:fld>

			<xsl:choose>
				<!--chinese-->
        
        <!--start liuyin 20121231 修改对象数据集时间和符号丢失 -->
        <xsl:when test="$datetime='&quot;yyyy年M月d日星期Wh时mm分s秒&quot;'">
          <xsl:attribute name="id">
            <xsl:value-of select="'{07875FEA-D690-4736-8A76-82A729C3C05C}'"/>
          </xsl:attribute>
          <xsl:attribute name="type">
            <xsl:value-of select="'datetime9'"/>
          </xsl:attribute>
        </xsl:when>
        <!--end liuyin 20121231 修改对象数据集时间和符号丢失 -->
        
				<xsl:when test="$datetime='&quot;yyyy-M-d&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{07875FEA-D690-4736-8A76-82A729C3C05C}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime1'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;yyyy年M月d日&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{96BC9E2B-6B32-4D1B-BDCC-F50018547A65}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime2'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;yyyy年M月d日星期W&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{8BB06C08-29DE-407B-A438-1C7FF038DDE3}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime3'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;yyyy/M/d&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{9665A019-4947-41D6-9725-AD4AEDB7F53B}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime5'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;yyyy年M月&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{25F62B4B-0B15-4722-A740-B80BDE13B3A7}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime6'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;yy.M.d&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{E0A96566-746A-4A30-9A4E-3786E5CCB122}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime7'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;HH时mm分&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{508339BD-2B22-44FB-BA39-3BA87899AC6B}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime10'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;HH时mm分ss秒&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{D99B4EEC-0E7A-4824-A173-322870626312}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime11'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;AMPMh时mm分&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{8F11DB01-CAD1-45E6-955B-5DD3327F6386}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime12'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;AMPMh时m分&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{CBFD99C6-F06A-4423-A82B-CEB801BEDD4A}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime13'"/>
					</xsl:attribute>
				</xsl:when>
				<!--english-->
				<xsl:when test="$datetime='&quot;M/d/yyyy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{7B330999-4E9C-4010-8C46-70DA5A72AF27}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime1'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;dddd, MMMM dd, yyyy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{DEC9F1D0-C326-433E-832C-464AF1F47893}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime2'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;d MMMM yyyy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{1DDEB56B-E768-4798-BEC5-64372AB82BBA}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime3'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;MMMM d, yyyy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{DDDC536B-EF1F-47DC-AAF3-CCA2207CA5A6}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime4'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;dd-MMM-yy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{E3692678-25E9-46AF-9841-0CBA8D57A35C}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime5'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;MMMM yy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{C4A1ED76-F930-4115-97C8-3A834E14B9FE}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime6'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;MMM-yy&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{9AC57DB4-732E-4441-91A3-A18EDD453AB8}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime7'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;M/d/yyyy h:mm AM/PM&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{B85D4F64-3C8E-48B4-8E00-F11AFA59A441}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime8'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;M/d/yyyy h:mm:ss AM/PM&quot;' and $type='DATE'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{ED8D6149-125B-4C68-9197-4070E7C9D1DF}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime9'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;HH:mm&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{0382C5FD-1CFE-493D-8F76-8086E8236E25}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime10'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;HH:mm:ss&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{2BCDC95A-F615-4A6D-9A0D-C38301346329}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime11'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;h:mm AM/PM&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{CE51008B-08D5-4258-85A5-F244CC987B57}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime12'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="$datetime='&quot;h:mm:ss AM/PM&quot;' and $type='TIME'">
					<xsl:attribute name="id">
						<xsl:value-of select="'{274F25F7-9462-443B-82A4-07E0FB7330FB}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime13'"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="id">
						<xsl:value-of select="'{07875FEA-D690-4736-8A76-82A729C3C05C}'"/>
					</xsl:attribute>
					<xsl:attribute name="type">
						<xsl:value-of select="'datetime1'"/>
					</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
      
      <!--start liuyin 20130317 修改bug2768文档转换后需要修复-->
      <!--<xsl:apply-templates select="../../字:句_419D[字:文本串_415B]/字:句属性_4158"/>-->
			<xsl:apply-templates select="../../字:句_419D[字:文本串_415B]/字:句属性_4158[字:语言标识_414C]"/>
      <!--end liuyin 20130317 修改bug2768文档转换后需要修复-->

      <a:t>
				<xsl:for-each select="../../字:句_419D/*">
					<xsl:choose>
						<xsl:when test="name(.)='字:空格符_4161'">
							<xsl:call-template name="blankSpace">
								<xsl:with-param name="num" select="@个数_4162"/>
							</xsl:call-template>
						</xsl:when>
						<xsl:when test="name(.)='字:制表符_415E'">
							<xsl:value-of select="'&#x9;'"/>
						</xsl:when>
						<xsl:when test="name(.)='字:文本串_415B'">
							<xsl:value-of select="."/>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
			</a:t>
		</a:fld>
	</xsl:template>
	<!--4月23日jjyan改-->

	<xsl:template name="blankSpace">
		<xsl:param name="num"/>
    
    <!--start liuyin 20130102 修改句中的空格符，如果字之间只有一个空格则转换后丢失-->
    <xsl:choose>
      <xsl:when test="$num &gt; 0">
        <xsl:value-of select="' '"/>
        <xsl:call-template name="blankSpace">
          <xsl:with-param name="num" select="$num - 1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="' '"/>
      </xsl:otherwise>
    </xsl:choose>
    <!--end liuyin 20130102 修改句中的空格符，如果字之间只有一个空格则转换后丢失-->
    
	</xsl:template>

	<xsl:template name="bodyPr">
		<a:bodyPr>
			<!-- 10.27 黎美秀修改 全部增加round-->
			<xsl:if test="图:边距_803D/@左_C608">
				<xsl:attribute name="lIns">
					<xsl:value-of select="round(图:边距_803D/@左_C608 * 12700)"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="图:边距_803D/@右_C60A">
				<xsl:attribute name="rIns">
					<xsl:value-of select="round(图:边距_803D/@右_C60A * 12700)"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="图:边距_803D/@上_C609">
				<xsl:attribute name="tIns">
					<xsl:value-of select="round(图:边距_803D/@上_C609 * 12700)"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="图:边距_803D/@下_C60B">
				<xsl:attribute name="bIns">
					<xsl:value-of select="round(图:边距_803D/@下_C60B * 12700)"/>
				</xsl:attribute>
			</xsl:if>
      
      <!--2014-03-26, tangjiang, 修复文本框水平对齐方式 start -->
      <!--@anchorCtr不存在时默认值为0；uof中的center对应过来，其他指定值及未指定值转为非center-->
      <xsl:choose>
        <xsl:when test="图:对齐_803E/@水平对齐_421D='center'">
          <xsl:attribute name="anchorCtr">1</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="anchorCtr">0</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
			<!--这个有问题 以后有待解决 李娟 2012.02.26-->
      <!--2014-03-26, tangjiang, 修复文本框水平对齐方式 -->
      
			<xsl:for-each select="图:对齐_803E/@文字对齐_421E">
				<xsl:attribute name="anchor">
					<xsl:choose>
						<xsl:when test=".='bottom'">b</xsl:when>
            <!--start liuyin 20130327 修改2783 文本框中的对齐不正确-->
						<xsl:when test=".='center'">ctr</xsl:when>
            <!--end liuyin 20130327 修改2783 文本框中的对齐不正确-->
            <xsl:when test=".='middle'">ctr</xsl:when>
						<xsl:when test=".='top'">t</xsl:when>
						<xsl:when test=".='justify'">just</xsl:when>
						<xsl:otherwise>t</xsl:otherwise>
					</xsl:choose>

				</xsl:attribute>
			</xsl:for-each>

			<!--2010.10.25 罗文甜 修改文字排列方向-->
			<xsl:for-each select="图:文字排列方向_8042">
				<xsl:attribute name="vert">
					<xsl:choose>
						<xsl:when test=".='t2b-l2r-0e-0w'">horz</xsl:when>
						<xsl:when test=".='r2l-t2b-90e-90w'">vert</xsl:when>
						<xsl:when test=".='r2l-t2b-0e-90w'">eaVert</xsl:when>
						<xsl:when test=".='l2r-b2t-270e-270w'">vert270</xsl:when>
						<xsl:when test=".='t2b-r2l-0e-0w'">horz</xsl:when>
						<xsl:otherwise>horz</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:for-each>
			<xsl:for-each select="@是否自动换行_8047">
				<xsl:attribute name="wrap">
					<xsl:choose>
						<xsl:when test=".='false' or .='0' or .='off'">none</xsl:when>
						<xsl:otherwise>square</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:for-each>
			<xsl:for-each select="@是否大小适应文字_8048">
				<xsl:choose>
					<xsl:when test=".='true' or .='1'or .='on'">
						<a:spAutoFit/>
					</xsl:when>
				</xsl:choose>
			</xsl:for-each>
			<!--a:normAutofit/-->
			<!--前一链接后一链接还没转-->
		</a:bodyPr>
	</xsl:template>
	<xsl:template match="式样:文本式样_9914">
		<xsl:call-template name="pPr"/>
	</xsl:template>
	<xsl:template match="式样:段落式样_9912">
		<xsl:call-template name="pPr"/>
	</xsl:template>
	<xsl:template match="字:段落属性_419B">
		<xsl:call-template name="pPr"/>
	</xsl:template>
	<xsl:template match="字:句属性_4158">
		<a:rPr>
			<xsl:call-template name="rPr"/>
		</a:rPr>
	</xsl:template>

</xsl:stylesheet>
