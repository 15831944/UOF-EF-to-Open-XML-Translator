<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                 xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
           
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:公式="http://schemas.uof.org/cn/2009/equations"
	xmlns:对象="http://schemas.uof.org/cn/2009/objects"
				xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
				xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <!--李娟 2012.01.05-->
  <xsl:import href="txBody.xsl"/>
  <xsl:import href="table.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="cSld">
    <p:cSld>

      <xsl:if test=".//演:背景_6B2C/*">
        <xsl:for-each select="演:背景_6B2C">
          <p:bg>
            <p:bgPr>
              <xsl:call-template name="fill"/>
              <a:effectLst/>
            </p:bgPr>
          </p:bg>
        </xsl:for-each>
      </xsl:if>
      <p:spTree>
        <p:nvGrpSpPr>
          <p:cNvPr id="1" name=""/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="0" cy="0"/>
            <a:chOff x="0" y="0"/>
            <a:chExt cx="0" cy="0"/>
          </a:xfrm>
        </p:grpSpPr>
        <!--2011-1-10罗文甜，页眉页脚：待添加，备注页眉页脚-->

        <!--<xsl:choose>
          <xsl:when test="name(.)='演:幻灯片_6C0F'">-->
        <!--2011-4-22 罗文甜，增加首页页眉页脚不显示-->

        <xsl:choose>

          <xsl:when test="name(.)='演:幻灯片_6C0F'">
            <xsl:variable name="headerFooterRef" select="./@页眉页脚引用_6C15"/>
            <xsl:choose>
              <xsl:when test="position()='1' and /uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/@标题幻灯片中是否显示_B64B='false'">
                <xsl:for-each select="uof:锚点_C644">
                  <xsl:call-template name="shape"/>
                </xsl:for-each>
              </xsl:when>
              <xsl:otherwise>
                <xsl:variable name="masterid" select="@母版引用_6B26"/>
                <xsl:for-each select="uof:锚点_C644">
                  <xsl:call-template name="shape"/>
                </xsl:for-each>
                <!--xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页眉页脚集/演:幻灯片页眉页脚/@演:标题幻灯片中不显示='false' or
                    not(/uof:UOF/uof:演示文稿/演:公用处理规则/演:页眉页脚集/演:幻灯片页眉页脚/@演:标题幻灯片中不显示)"-->
                <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[./@标识符_B646=$headerFooterRef]/@是否显示页脚_B648='true'">
                  <xsl:for-each select="//演:母版_6C0D[@标识符_6BE8=$masterid]//uof:锚点_C644[uof:占位符_C626/@类型_C627='footer']">
                    <xsl:call-template name="shape"/>
                  </xsl:for-each>
                </xsl:if>
                <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[./@标识符_B646=$headerFooterRef]/@是否显示幻灯片编号_B64A='true'">
                  <xsl:for-each select="//演:母版_6C0D[@标识符_6BE8=$masterid]//uof:锚点_C644[uof:占位符_C626/@类型_C627='number']">
                    <xsl:call-template name="shape"/>
                  </xsl:for-each>
                </xsl:if>
                <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[./@标识符_B646=$headerFooterRef]/@是否显示日期和时间_B647='true'">
                  <xsl:for-each select="//演:母版_6C0D[@标识符_6BE8=$masterid]//uof:锚点_C644[uof:占位符_C626/@类型_C627='date']">
                    <xsl:call-template name="shape"/>
                  </xsl:for-each>
                </xsl:if>

                <!--start liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->
                <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[./@标识符_B646=$headerFooterRef]/规则:日期和时间字符串_B643">
                  <xsl:for-each select="//演:母版_6C0D[@标识符_6BE8=$masterid]//uof:锚点_C644[uof:占位符_C626/@类型_C627='date']">
                    <xsl:call-template name="shape"/>
                  </xsl:for-each>
                </xsl:if>
                <!--end liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->

              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

          <!--start liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->
          <!--<xsl:when test="name(.)=演:幻灯片备注_6B1D">-->
          <xsl:when test="name(.)='演:幻灯片备注_6B1D'">
            <!--end liuyin 20130522 修改2916 备注页的页眉页脚丢失，普通视图下幻灯片编号转为文字-->
            <!--xsl:variable name="noteMasterid" select="@演:母版引用"/-->
            <xsl:for-each select="uof:锚点_C644">
              <xsl:call-template name="shape"/>
            </xsl:for-each>
            <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/规则:页脚_B644">
              <xsl:for-each select="//演:母版_6C0D[@类型_6BEA='notes']//uof:锚点_C644[uof:占位符_C626/@类型_C627='footer']">
                <xsl:call-template name="shape"/>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/规则:页眉_B64D">
              <xsl:for-each select="//演:母版_6C0D[@类型_6BEA='notes']//uof:锚点_C644[uof:占位符_C626/@类型_C627='header']">
                <xsl:call-template name="shape"/>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/@是否显示页码_B650='true'">
              <xsl:for-each select="//演:母版_6C0D[@类型_6BEA='notes']//uof:锚点_C644[uof:占位符_C626/@类型_C627='number']">
                <xsl:call-template name="shape"/>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C/@是否显示日期和时间_B647='true'">
              <xsl:for-each select="//演:母版_6C0D[@类型_6BEA='notes']//uof:锚点_C644[uof:占位符_C626/@类型_C627='date']">
                <xsl:call-template name="shape"/>
              </xsl:for-each>
            </xsl:if>
          </xsl:when>

          <xsl:when test="name(.)='演:母版_6C0D'">
            <xsl:for-each select="uof:锚点_C644">
              <xsl:call-template name="mastershape"/>
            </xsl:for-each>
          </xsl:when>

          <!--2014-03-29, tangjiang, 修复母板占位符重复出现两次 start -->
          <xsl:when test="name(.)='规则:页面版式_B652'">
            <xsl:for-each select="./uof:锚点_C644[not(uof:占位符_C626)]">
              <xsl:call-template name="shape"/>
            </xsl:for-each>
            
            <xsl:for-each select="./uof:锚点_C644[uof:占位符_C626/@类型_C627!='date' and uof:占位符_C626/@类型_C627!='footer' and uof:占位符_C626/@类型_C627!='number']">
              <xsl:call-template name="shape"/>
            </xsl:for-each>

            <xsl:variable name="dateHolderCount" select="count(.//uof:锚点_C644/uof:占位符_C626[@类型_C627='date'])"/>
            <xsl:for-each select="./uof:锚点_C644[uof:占位符_C626/@类型_C627='date']">
              <xsl:if test="position()=$dateHolderCount">
                <xsl:call-template name="shape"/>
              </xsl:if>
            </xsl:for-each>
            <xsl:variable name="footerHolderCount" select="count(.//uof:锚点_C644/uof:占位符_C626[@类型_C627='footer'])"/>
            <xsl:for-each select="./uof:锚点_C644[uof:占位符_C626/@类型_C627='footer']">
              <xsl:if test="position()=$footerHolderCount">
                <xsl:call-template name="shape"/>
              </xsl:if>
            </xsl:for-each>
            <xsl:variable name="numberHolderCount" select="count(.//uof:锚点_C644/uof:占位符_C626[@类型_C627='number'])"/>
            <xsl:for-each select="./uof:锚点_C644[uof:占位符_C626/@类型_C627='number']">
              <xsl:if test="position()=$numberHolderCount">
                <xsl:call-template name="shape"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:when>
          <!--end 2014-03-29, tangjiang, 修复母板占位符重复出现两次 -->
          
          <xsl:otherwise>
            <!--start liuyin 20130307 修改幻灯片母版引用相关，转换后需要修复才能打开-->
            <xsl:for-each select="./uof:锚点_C644">
              <xsl:call-template name="shape"/>
            </xsl:for-each>
            <!--end liuyin 20130307 修改幻灯片母版引用相关，转换后需要修复才能打开-->
          </xsl:otherwise>
        </xsl:choose>
      </p:spTree>
      <p:extLst>
        <p:ext>
          <xsl:attribute name="uri">
            <xsl:value-of select="concat('{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}','')"/>
          </xsl:attribute>
          <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="747778485"/>
        </p:ext>
      </p:extLst>
    </p:cSld>
  </xsl:template>
  <xsl:template name="shape">
    <xsl:variable name="obj">
      <xsl:value-of select="@图形引用_C62E"/>
    </xsl:variable>

    <xsl:variable name="name">
      <xsl:choose>
        <xsl:when test="uof:占位符_C626/@类型_C627">
          <xsl:value-of select="uof:占位符_C626/@类型_C627"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>

    <!--<xsl:if test="图:图形/@图:标识符=$obj and not(图:图形/@图:组合列表)">-->
    <xsl:if test="图:图形_8062/@标识符_804B=$obj and not(图:图形_8062/@组合列表_8064)">

      <xsl:choose>
        <!--liuyangyang 2015-02-02 修改ole对象转换丢失效果-->
        <xsl:when test="图:图形_8062/图:图片数据引用_8037 and not(图:图形_8062/ole) and not(图:图形_8062/图:其他对象引用_8038)">
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
        <!--<xsl:when test="图:图形_8062/图:其他对象引用_8038 | 图:图形_8062/图:图片数据引用_8037 ">-->
        <!--<xsl:when test="图:图形_8062/图:图片数据引用_8037 and not(图:图形_8062/图:其他对象引用_8038)">-->
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <!--<xsl:when test="图:图形_8062/图:图片数据引用_8037">-->
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <!--end liuyangyang -->
          <xsl:variable name="picId">
            <xsl:value-of select="图:图形_8062/图:图片数据引用_8037"/>
          </xsl:variable>
          <xsl:variable name="tag">
            <xsl:value-of select="图:图形_8062/@标识符_804B"/>
          </xsl:variable>
          <!--2011-4-5罗文甜-->
          <xsl:variable name="picture">
            <xsl:choose>
              <xsl:when test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                <xsl:value-of select=".//图:Web文字_804F"/>
              </xsl:when>
              <xsl:otherwise>null</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          
          <!--<xsl:for-each select="ancestor::uof:UOF/uof:对象集/uof:其他对象[@uof:标识符=$picId]">-->

          <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
          <!--<xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picId]">-->
          <!--start liqiuling 2013-03-29 修改bug：oochart转uof picture 互操作需要修复 -->
          <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
            <xsl:choose>
              <xsl:when test="@公共类型_D706='png' or 'jpg' or 'jpeg' or 'gif' or 'tiff' or 'bmp' or 'wdp'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$picId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('pic',$picId)"/>
                      </xsl:attribute>
                      <!--end liuyin 20121231 修改图片裁剪效果丢失 -->
                      <!--end liqiuling 2013-03-29 修改bug：oochart转uof picture 互操作需要修复 -->

                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>
                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr/>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
                    <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022 and ./图:图形_8062/图:图片数据引用_8037=$picId ">
                      <a:srcRect>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="t">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
                          <xsl:attribute name="b">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="l">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
                          <xsl:attribute name="r">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026*1000)"/>
                          </xsl:attribute>
                        </xsl:if>

                        <!--<xsl:if test="//图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
											<xsl:attribute name="b">
												<xsl:value-of  select="round((//图:大小_8060/@长_C604)-(//图:图片属性_801E/图:图片裁剪_8022/图:上_8023+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024))"/>
                      </xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
											<xsl:attribute name="b">
												<xsl:value-of  select="round(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024)"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
											<xsl:attribute name="r">
												<xsl:value-of  select="round((//图:大小_8060/@宽_C605)-(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026))"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
											<xsl:attribute name="r">
												<xsl:value-of  select="round(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026)"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--end liuyin 20130110 修改图片裁剪效果丢失 -->
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <!--   
                    12.29 黎美秀修改 xpath路径不完整 可能是layout中
                     <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿/演:主体//uof:锚点_C644[@uof:图形引用=$tag]">
                     
                     2908620
                    -->
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                      <!-- 09.10.30 deleted by myx
                      <xsl:if test=".//图:旋转角度!='0.0'">
                        <xsl:call-template name="rot"/>
                      </xsl:if>-->
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <!--liuyangyang 2015-02-02 修改ole对象转换丢失效果-->
        <xsl:when test="图:图形_8062/图:图片数据引用_8037 and 图:图形_8062/ole">
          <xsl:copy-of select="图:图形_8062/ole/p:graphicFrame"/>
        </xsl:when>
        <!--end liuyangyang-->
        <!--start liuyin 20121211 修改音视频丢失-->
        <xsl:when test="图:图形_8062/图:其他对象引用_8038">
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <xsl:variable name="picId">
            <xsl:value-of select="图:图形_8062/图:图片数据引用_8037"/>
          </xsl:variable>
          <xsl:variable name="aviId">
            <xsl:value-of select="图:图形_8062/图:其他对象引用_8038"/>
          </xsl:variable>
          <xsl:variable name="tag">
            <xsl:value-of select="图:图形_8062/@标识符_804B"/>
          </xsl:variable>
          <xsl:variable name="picture">
            <xsl:choose>
              <xsl:when test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                <xsl:value-of select=".//图:Web文字_804F"/>
              </xsl:when>
              <xsl:otherwise>null</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!--start liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->
          <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picId]">
            <xsl:choose>
              <xsl:when test="@私有类型_D707='emf'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$picId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('pic',$picId)"/>
                      </xsl:attribute>
                      <!--end liuyin 20121231 修改图片裁剪效果丢失 -->

                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>
                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr/>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
                    <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022 and ./图:图形_8062/图:图片数据引用_8037=$picId ">
                      <a:srcRect>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="t">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
                          <xsl:attribute name="b">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="l">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
                          <xsl:attribute name="r">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <!--end liuyin 20130110 修改图片裁剪效果丢失 -->
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
          <!--end liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->

          <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$aviId]">
            <xsl:choose>

              <!--start liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->
              <xsl:when test="@私有类型_D707='eof'">
              </xsl:when>

              <!--start liuyin 20130316 修改OLE对象经过转换后，ooxml文档无法打开-->
              <xsl:when test="@私有类型_D707='bin'">
              </xsl:when>
              <!--start liuyin 20130316 修改OLE对象经过转换后，ooxml文档无法打开-->

              <!--end liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->

              <xsl:when test="@公共类型_D706='avi' or 'mp3' or 'wav' or 'mid' or 'wmv' or 'mpeg' or 'rmvb' or 'rm'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:variable name="mediaId" select="./@标识符_D704"/>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$mediaId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('media',$mediaId)"/>
                      </xsl:attribute>
                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>

                      <a:hlinkClick r:id="" action="ppaction://media"/>

                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr>

                      <!--start liuyin 20121211 修改音视频丢失-->
                      <!--<xsl:choose>
                        <xsl:when test="@公共类型_D706='mp3' or 'wav' or 'mid'">
                          <a:audioFile>
                            <xsl:attribute name="r:link">
                              <xsl:value-of select="concat('rId',@标识符_D704)"/>
                            </xsl:attribute>
                          </a:audioFile>
                        </xsl:when>
                        <xsl:when test="@公共类型_D706='avi' or'wmv' or 'mpeg' or 'rmvb' or 'rm'">-->
                      <a:videoFile>
                        <xsl:attribute name="r:link">
                          <xsl:value-of select="concat('rId',@标识符_D704)"/>
                        </xsl:attribute>
                      </a:videoFile>
                      <!--</xsl:when>
                      </xsl:choose>-->
                      <p:extLst>
                        <p:ext>
                          <xsl:attribute name="uri">
                            <xsl:value-of select="concat('{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}','')"/>
                          </xsl:attribute>
                          <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                            <xsl:attribute name="r:embed">
                              <xsl:value-of select="concat('rId',@标识符_D704,'1')"/>
                            </xsl:attribute>
                          </p14:media>
                        </p:ext>
                      </p:extLst>
                      <!--end liuyin 20121211 修改音视频丢失-->

                    </p:nvPr>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <xsl:if test="//图:图片属性_801E/图:图片裁剪_8022">
                      <a:srcRect>
                        <xsl:if test="//图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="b">
                            <!--代表高度-->
                            <xsl:value-of  select="round((//图:大小_8060/@长_C604)-(//图:图片属性_801E/图:图片裁剪_8022/图:上_8023+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024))"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="r">
                            <xsl:value-of  select="round((//图:大小_8060/@宽_C605)-(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026))"/>
                          </xsl:attribute>
                        </xsl:if>
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <!--   
                    12.29 黎美秀修改 xpath路径不完整 可能是layout中
                     <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿/演:主体//uof:锚点_C644[@uof:图形引用=$tag]">
                     
                     2908620
                    -->
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                      <!-- 09.10.30 deleted by myx
                      <xsl:if test=".//图:旋转角度!='0.0'">
                        <xsl:call-template name="rot"/>
                      </xsl:if>-->
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>

            </xsl:choose>
          </xsl:for-each>
          <xsl:if test="//公式:公式集_C200/公式:数学公式_C201/@标识符_C202=$aviId">
            <xsl:value-of disable-output-escaping="yes" select="'&lt;mc:AlternateContent xmlns:mc=&quot;http://schemas.openxmlformats.org/markup-compatibility/2006&quot; xmlns:a14=&quot;http://schemas.microsoft.com/office/drawing/2010/main&quot; &gt;'"/>
            <mc:Choice Requires="a14">
              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr>
                    <xsl:attribute name="id">
                      <xsl:call-template name="mediaIdConvetor">
                        <xsl:with-param name="mediaId" select="$tag"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="'TextBox 1'"/>
                    </xsl:attribute>
                  </p:cNvPr>
                  <p:cNvSpPr txBox="1"/>
                  <p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                  <a:xfrm>
                    <!-- 09.10.30 added by myx -->
                    <xsl:if test=".//图:旋转角度_804D!='0.0'">
                      <xsl:attribute name="rot">
                        <!--2011-5-26 luowentian-->
                        <xsl:choose>
                          <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                            <xsl:value-of select="round((360-.//图:旋转角度_804D)*60000)"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="round(.//图:旋转角度_804D*60000)"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                    </xsl:if>
                    <!--2011-3-20罗文甜，修改Bug-->
                    <xsl:if test=".//图:翻转_803A">
                      <xsl:call-template name="flip"/>
                    </xsl:if>
                    <a:off>
                      <xsl:attribute name="x">
                        <xsl:choose>
                          <xsl:when test="contains(./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108,'-')">
                            <xsl:value-of select="'0'"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="round(./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108*12700)"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:attribute name="y">
                        <xsl:value-of select="round(./uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108 * 12700)"/>
                      </xsl:attribute>
                    </a:off>
                    <a:ext>
                      <xsl:attribute name="cx">
                        <xsl:value-of select="round(./uof:大小_C621/@宽_C605 * 12700 * 1.1)"/>
                      </xsl:attribute>
                      <xsl:attribute name="cy">
                        <xsl:value-of select="round(./uof:大小_C621/@长_C604 * 12700)"/>
                      </xsl:attribute>
                    </a:ext>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                  <a:noFill/>
                </p:spPr>
                <p:txBody>
                  <a:bodyPr wrap="square" rtlCol="0">
                    <a:spAutoFit/>
                  </a:bodyPr>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr/>
                    <a14:m>
                      <m:oMathPara xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
                        <m:oMathParaPr>
                          <m:jc m:val="centerGroup"/>
                        </m:oMathParaPr>
                        <xsl:copy-of select="//公式:公式集_C200/公式:数学公式_C201[@标识符_C202=$aviId]/m:oMath"/>
                      </m:oMathPara>
                    </a14:m>
                    <a:endParaRPr lang="zh-CN" altLang="en-US" dirty="0"/>
                  </a:p>
                </p:txBody>
              </p:sp>
            </mc:Choice>
            <xsl:value-of select="'&lt;/mc:AlternateContent&gt;'" disable-output-escaping="yes"/>
          </xsl:if>
        </xsl:when>
        <!--end liuyin 20121211 修改音视频丢失-->

        <!--       2010.3.9 黎美秀 增加文字表的处理       -->
        <xsl:when test=".//图:内容_8043/字:文字表_416C">
          <xsl:call-template name="table">
            <xsl:with-param name="tblobj" select="$obj"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>

          <xsl:choose>
            <!--xsl:when test=".//图:类别_8019='61' or .//图:类别_8019='71' or .//图:类别_8019='74' or .//图:类别_8019='77'"-->
            <xsl:when test=".//图:类别_8019='61' or .//图:类别_8019='71' or .//图:类别_8019='74' or (.//图:类别_8019='64' and .//图:名称_801A='Curved Connector')">
              <p:cxnSp>
                <p:nvCxnSpPr>
                  <p:cNvPr>
                    <xsl:attribute name="id">
                      <xsl:value-of select="substring($obj,4,5)"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="$name"/>
                    </xsl:attribute>
                    <!--2011-4-5罗文甜，增加web文字-->
                    <xsl:if test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                      <xsl:attribute name="descr">
                        <xsl:value-of select=".//图:Web文字_804F"/>
                      </xsl:attribute>
                    </xsl:if>
                  </p:cNvPr>
                  <p:cNvCxnSpPr>
                    <!--2010-12-20罗文甜：增加连接线规则-->
                    <xsl:if test="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027">
                      <xsl:apply-templates select="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027"/>
                    </xsl:if>
                  </p:cNvCxnSpPr>
                  <p:nvPr/>
                </p:nvCxnSpPr>

                <p:spPr>
                  <xsl:if test="@图形引用_C62E=$obj">

                    <xsl:call-template name="xfrm"/>
                  </xsl:if>
                  <xsl:call-template name="prstGeom"/>

                  <xsl:if test=".//图:填充_804C">

                    <xsl:call-template name="fill"/>
                  </xsl:if>
                  <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                    <xsl:call-template name="ln"/>
                  </xsl:if>
                  <!--2011-01-11罗文甜，增加阴影判断-->
                  <xsl:if test=".//图:阴影_8051">
                    <xsl:call-template name="shadow"/>
                  </xsl:if>
                  <!--09.10.30 deleted by myx
                    <xsl:if test=".//图:旋转角度!='0.0' or .//图:翻转">
                      <xsl:call-template name="rot"/>
                    </xsl:if>-->
                </p:spPr>

              </p:cxnSp>
            </xsl:when>

            <xsl:otherwise>

              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr>
                    <xsl:attribute name="id">

                      <xsl:value-of select="substring($obj,4,5)"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:choose>
                        <xsl:when test="not($name='')">
                          <xsl:value-of select="$name"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="'Title 2'"/>
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:attribute>
                    <!--2011-4-5罗文甜，增加web文字-->
                    <xsl:if test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                      <xsl:attribute name="descr">
                        <xsl:value-of select=".//图:Web文字_804F"/>
                      </xsl:attribute>
                    </xsl:if>
                  </p:cNvPr>
                  <!--2011-3-9罗文甜，增加判断-->
                  <p:cNvSpPr>
                    <!--2012 05 08 增加对文本框的判断 lj-->
                    <xsl:if test="$name='' and ancestor::演:幻灯片_6C0F">
                      <xsl:attribute  name="txBox">1</xsl:attribute>
                    </xsl:if>
                    <xsl:if test="string-length($name) &gt; 0 or .//图:缩放是否锁定纵横比_8055='1'">
                      <a:spLocks>
                        <xsl:if test="string-length($name) &gt; 0">
                          <xsl:attribute name="noGrp">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test=".//图:缩放是否锁定纵横比_8055='1'">
                          <xsl:attribute name="noChangeAspect">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>
                      </a:spLocks>
                    </xsl:if>
                  </p:cNvSpPr>
                  <p:nvPr>
                    <xsl:if test="string-length($name) &gt; 0">
                      <p:ph>
                        <xsl:variable name="type">
                          <xsl:choose>
                            <xsl:when test="$name='object'">obj</xsl:when>
                            <xsl:when test="$name='vertical-title' or $name='title'">title</xsl:when>
                            <xsl:when test="$name='vertical-text' or $name='text'">body</xsl:when>

                            <!--start liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->
                            <xsl:when test="$name='centertitle'">ctrTitle</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">subTitle</xsl:when>
                            <!--<xsl:when test="$name='centertitle'">title</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">title</xsl:when>-->
                            <!--end liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->

                            <xsl:when test="$name='date'">dt</xsl:when>
                            <xsl:when test="$name='number'">sldNum</xsl:when>
                            <xsl:when test="$name='footer'">ftr</xsl:when>
                            <xsl:when test="$name='header'">hdr</xsl:when>
                            <xsl:when test="$name='chart'">chart</xsl:when>
                            <xsl:when test="$name='table'">tbl</xsl:when>
                            <xsl:when test="$name='clipart'">clipArt</xsl:when>
                            <xsl:when test="$name='media-clip'">media</xsl:when>
                            <xsl:when test="$name='notes' and @是否显示缩略图_C630='true'">sldImg</xsl:when>
                            <xsl:when test="$name='notes' and not(@是否显示缩略图_C630)">body</xsl:when>
                            <xsl:otherwise>body</xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <!--2010.04.27 myx -->
                        <xsl:variable name="idx">
                          <xsl:choose>
                            <xsl:when test="$name='vertical-title' or $name='title'">-1</xsl:when>
                            <xsl:when test="$name='centertitle'">-1</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">1</xsl:when>
                            <xsl:when test="$name='date'">2</xsl:when>
                            <xsl:when test="$name='number'">3</xsl:when>
                            <xsl:when test="$name='footer'">4</xsl:when>
                            <xsl:when test="$name='header'">5</xsl:when>
                            <xsl:when test="$name='chart'">6</xsl:when>
                            <xsl:when test="$name='table'">7</xsl:when>
                            <xsl:when test="$name='clipart'">8</xsl:when>
                            <xsl:when test="$name='media-clip'">9</xsl:when>
                            <xsl:when test="$name='notes' and @是否显示缩略图_C630='true'">10</xsl:when>
                            <xsl:when test="$name='notes' and not(@是否显示缩略图_C630)">11</xsl:when>
                            <xsl:when test="$name='object'">12</xsl:when>
                            <xsl:when test="$name='vertical-text' or $name='text'">
                              <xsl:value-of select="13+count(preceding-sibling::*[uof:占位符_C626[@类型_C627]='vertical_text' or @uof:占位符_C626[@类型_C627]='text'])"/>
                            </xsl:when>
                            <xsl:otherwise>15</xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <xsl:attribute name="type">
                          <xsl:value-of select="$type"/>
                        </xsl:attribute>

                        <!--2014-03-20, tangjiang, 修复互操作多出项目符号 start -->
                        <!--2010.04.27 myx -->
                        <!--为修改amation的fade，paragraph的互操作 李娟 2012 0507-->
                        <xsl:choose>
                          <xsl:when test="$type='body'and $idx ='13'">
                            <xsl:attribute name="idx">
                              <xsl:value-of select="'1'"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="$idx!='-1' and $idx!='13' and $idx!='1'">
                            <xsl:attribute name="idx">
                              <xsl:value-of select="$idx"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:otherwise>

                          </xsl:otherwise>
                        </xsl:choose>
                        <!-- end 2014-03-20, tangjiang, 修复互操作多出项目符号 -->
                        
                        <xsl:if test="$name='vertical-title' or $name='vertical-text' or $name='vertical-subtitle'">
                          <xsl:attribute name="orient">vert</xsl:attribute>
                        </xsl:if>

                      </p:ph>
                    </xsl:if>
                  </p:nvPr>
                </p:nvSpPr>
                <p:spPr>

                  <xsl:if test="@图形引用_C62E=$obj">
                    <xsl:call-template name="xfrm"/>
                  </xsl:if>
                  <!--2010-12-26罗文甜，增加自定义图形(自定义曲线，自由曲线)-->
                  <xsl:if test=".//图:类别_8019">
                    <xsl:choose>
                      <xsl:when test=".//图:路径_801C">
                        <!--uof2.0用路径来描述关键点坐标 李娟 2012.01.05-->
                        <xsl:call-template name="custGeom"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:call-template name="prstGeom"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <!--增加对艺术字的处理 李娟 2012.02.28-->
                  <xsl:if test=".//图:填充_804C  and not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">

                    <xsl:call-template name="fill"/>
                  </xsl:if>
                  <xsl:if test=".//图:填充_804C and .//图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
                    <a:noFill/>
                  </xsl:if>

                  <!--<xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) and not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">-->
                  <xsl:if test="(.//图:线_8057 | .//图:箭头_805D) and not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">

                    <xsl:call-template name="ln"/>
                  </xsl:if>
                  <xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) 
            and (./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </xsl:if>
                  <!--2011-01-11罗文甜，增加阴影判断-->
                  <xsl:if test=".//图:阴影_8051 and  not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
                    <xsl:call-template name="shadow"/>
                  </xsl:if>
                  <!--增加三维效果 李娟 2012。02.22-->
                  <xsl:if test=".//图:三维效果_8061">

                    <xsl:call-template name="图:三维效果_8061"/>
                  </xsl:if>
                </p:spPr>

                <xsl:if test="图:图形_8062/图:文本_803C">
                  <xsl:for-each select="图:图形_8062/图:文本_803C">
                    <xsl:call-template name="txBody"/>
                  </xsl:for-each>
                </xsl:if>
                <!--新增艺术字转换 李娟 2012.02.28-->
                <xsl:if test="图:图形_8062//图:艺术字_802D">
                  <xsl:for-each select="图:图形_8062//图:艺术字_802D">
                    <xsl:call-template name="art"/>
                    <!--<xsl:apply-templates select="图:艺术字_802D" mode="txbody"/>-->
                  </xsl:for-each>
                </xsl:if>
                <!--2010.04.09 myx add 复制母版 body类型的锚点的文本内容到版式 -->
                <xsl:if test="not(/uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062/图:文本_803C/图:内容_8043) and ancestor::规则:页面版式_B652 and (uof:占位符_C626[@类型_C627]='vertical-text' or uof:占位符_C626[@类型_C627]='text' or uof:占位符_C626[@类型_C627]='title') ">

                  <xsl:variable name="layout_id" select="ancestor::规则:页面版式_B652/@标识符_6B0D"/>
                  <xsl:variable name="master" select="//演:母版_6C0D[@页面版式引用_6BEC/text()=$layout_id]"/>
                  <!--<xsl:variable name="master" select="//演:母版[演:页面版式引用/text()=$layout_id]"/>-->
                  <!--注销这段 不是很清楚text（） 什么意思i 李娟 2012.01.04·········-->
                  <xsl:variable name="target_type">
                    <xsl:choose>
                      <xsl:when test="uof:占位符_C626/@类型_C627='vertical-text' or uof:占位符_C626/@类型_C627='text'">text</xsl:when>
                      <xsl:otherwise>title</xsl:otherwise>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:apply-templates select="$master/uof:锚点_C644[uof:占位符_C626/@类型_C627=$target_type]/图:图形_8062/图:文本_803C/图:内容_8043" mode="layout"/>
                </xsl:if>
                <!--xsl:if test="图:图形/@图:其他对象">
            <xsl:call-template name="txBody"/>
          </xsl:if-->
              </p:sp>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="图:图形_8062/@标识符_804B=$obj and 图:图形_8062/@组合列表_8064">

      <xsl:call-template name="grpshape"/>
    </xsl:if>
  </xsl:template>

  <!--start liuyin 20130307 修改幻灯片母版引用相关，转换后需要修复才能打开-->
  <xsl:template name="mastershape">
    <xsl:variable name="obj">
      <xsl:value-of select="@图形引用_C62E"/>
    </xsl:variable>

    <xsl:variable name="name">
      <xsl:choose>
        <xsl:when test="uof:占位符_C626/@类型_C627">
          <xsl:value-of select="uof:占位符_C626/@类型_C627"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <!--<xsl:if test="图:图形/@图:标识符=$obj and not(图:图形/@图:组合列表)">-->
    <xsl:if test="图:图形_8062/@标识符_804B=$obj and not(图:图形_8062/@组合列表_8064)">

      <xsl:choose>
        <!--<xsl:when test="图:图形_8062/图:其他对象引用_8038 | 图:图形_8062/图:图片数据引用_8037 ">-->
        <xsl:when test="图:图形_8062/图:图片数据引用_8037 and not(图:图形_8062/图:其他对象引用_8038)">
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <!--<xsl:when test="图:图形_8062/图:图片数据引用_8037">-->
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <xsl:variable name="picId">
            <xsl:value-of select="图:图形_8062/图:图片数据引用_8037"/>
          </xsl:variable>
          <xsl:variable name="tag">
            <xsl:value-of select="图:图形_8062/@标识符_804B"/>
          </xsl:variable>
          <!--2011-4-5罗文甜-->
          <xsl:variable name="picture">
            <xsl:choose>
              <xsl:when test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                <xsl:value-of select=".//图:Web文字_804F"/>
              </xsl:when>
              <xsl:otherwise>null</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!--<xsl:for-each select="ancestor::uof:UOF/uof:对象集/uof:其他对象[@uof:标识符=$picId]">-->

          <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
          <!--<xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picId]">-->
          <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
            <xsl:choose>
              <xsl:when test="@公共类型_D706='png' or 'jpg' or 'jpeg' or 'gif' or 'tiff' or 'bmp' or 'wdp'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$picId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('pic',$picId)"/>
                      </xsl:attribute>
                      <!--end liuyin 20121231 修改图片裁剪效果丢失 -->

                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>
                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr/>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
                    <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022 and ./图:图形_8062/图:图片数据引用_8037=$picId ">
                      <a:srcRect>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="t">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
                          <xsl:attribute name="b">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="l">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
                          <xsl:attribute name="r">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026*1000)"/>
                          </xsl:attribute>
                        </xsl:if>

                        <!--<xsl:if test="//图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
											<xsl:attribute name="b">
												<xsl:value-of  select="round((//图:大小_8060/@长_C604)-(//图:图片属性_801E/图:图片裁剪_8022/图:上_8023+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024))"/>
                      </xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
											<xsl:attribute name="b">
												<xsl:value-of  select="round(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024)"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
											<xsl:attribute name="r">
												<xsl:value-of  select="round((//图:大小_8060/@宽_C605)-(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026))"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--<xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
											<xsl:attribute name="r">
												<xsl:value-of  select="round(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026)"/>
											</xsl:attribute>
										</xsl:if>-->
                        <!--end liuyin 20130110 修改图片裁剪效果丢失 -->
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <!--   
                    12.29 黎美秀修改 xpath路径不完整 可能是layout中
                     <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿/演:主体//uof:锚点_C644[@uof:图形引用=$tag]">
                     
                     2908620
                    -->
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                      <!-- 09.10.30 deleted by myx
                      <xsl:if test=".//图:旋转角度!='0.0'">
                        <xsl:call-template name="rot"/>
                      </xsl:if>-->
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>

        <!--start liuyin 20121211 修改音视频丢失-->
        <xsl:when test="图:图形_8062/图:其他对象引用_8038">
          <!--暂时去掉图形其他对象引用 李娟 2012.02.17-->
          <xsl:variable name="picId">
            <xsl:value-of select="图:图形_8062/图:图片数据引用_8037"/>
          </xsl:variable>
          <xsl:variable name="aviId">
            <xsl:value-of select="图:图形_8062/图:其他对象引用_8038"/>
          </xsl:variable>
          <xsl:variable name="tag">
            <xsl:value-of select="图:图形_8062/@标识符_804B"/>
          </xsl:variable>
          <xsl:variable name="picture">
            <xsl:choose>
              <xsl:when test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                <xsl:value-of select=".//图:Web文字_804F"/>
              </xsl:when>
              <xsl:otherwise>null</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <!--start liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->
          <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picId]">
            <xsl:choose>
              <xsl:when test="@私有类型_D707='emf'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$picId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('pic',$picId)"/>
                      </xsl:attribute>
                      <!--end liuyin 20121231 修改图片裁剪效果丢失 -->

                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>
                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr/>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <!--start liuyin 20121231 修改图片裁剪效果丢失 -->
                    <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022 and ./图:图形_8062/图:图片数据引用_8037=$picId ">
                      <a:srcRect>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="t">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:上_8023*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024">
                          <xsl:attribute name="b">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="l">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026">
                          <xsl:attribute name="r">
                            <xsl:value-of select="round(./图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026*1000)"/>
                          </xsl:attribute>
                        </xsl:if>
                        <!--end liuyin 20130110 修改图片裁剪效果丢失 -->
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
          <!--end liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->

          <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$aviId]">
            <xsl:choose>

              <!--start liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->
              <xsl:when test="@私有类型_D707='eof'">
              </xsl:when>
              <!--end liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->

              <xsl:when test="@公共类型_D706='avi' or 'mp3' or 'wav' or 'mid' or 'wmv' or 'mpeg' or 'rmvb' or 'rm'">
                <p:pic>
                  <p:nvPicPr>
                    <p:cNvPr>
                      <xsl:variable name="mediaId" select="./@标识符_D704"/>
                      <xsl:attribute name="id">
                        <xsl:call-template name="mediaIdConvetor">
                          <xsl:with-param name="mediaId" select="$mediaId"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <xsl:attribute name="name">
                        <xsl:value-of select="concat('media',$mediaId)"/>
                      </xsl:attribute>
                      <!--2011-4-5罗文甜-->
                      <xsl:if test="$picture!='null'">
                        <xsl:attribute name="descr">
                          <xsl:value-of select="$picture"/>
                        </xsl:attribute>
                      </xsl:if>

                      <a:hlinkClick r:id="" action="ppaction://media"/>

                    </p:cNvPr>
                    <p:cNvPicPr>
                      <a:picLocks noChangeAspect="1"/>
                    </p:cNvPicPr>
                    <p:nvPr>

                      <!--start liuyin 20121211 修改音视频丢失-->
                      <!--<xsl:choose>
                        <xsl:when test="@公共类型_D706='mp3' or 'wav' or 'mid'">
                          <a:audioFile>
                            <xsl:attribute name="r:link">
                              <xsl:value-of select="concat('rId',@标识符_D704)"/>
                            </xsl:attribute>
                          </a:audioFile>
                        </xsl:when>
                        <xsl:when test="@公共类型_D706='avi' or'wmv' or 'mpeg' or 'rmvb' or 'rm'">-->
                      <a:videoFile>
                        <xsl:attribute name="r:link">
                          <xsl:value-of select="concat('rId',@标识符_D704)"/>
                        </xsl:attribute>
                      </a:videoFile>
                      <!--</xsl:when>
                      </xsl:choose>-->
                      <p:extLst>
                        <p:ext>
                          <xsl:attribute name="uri">
                            <xsl:value-of select="concat('{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}','')"/>
                          </xsl:attribute>
                          <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                            <xsl:attribute name="r:embed">
                              <xsl:value-of select="concat('rId',@标识符_D704,'1')"/>
                            </xsl:attribute>
                          </p14:media>
                        </p:ext>
                      </p:extLst>
                      <!--end liuyin 20121211 修改音视频丢失-->

                    </p:nvPr>
                  </p:nvPicPr>
                  <p:blipFill>
                    <a:blip>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="concat('rId',$picId)"/>
                      </xsl:attribute>
                      <!--2011-2-9罗文甜，修改图片属性-->
                      <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(grayscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
                      <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='greyscale'">
                          <a:grayscl/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='erosion'">
                          <a:lum bright="70000" contrast="-70000"/>
                        </xsl:if>
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F='monochrome'">
                          <a:biLevel thresh="50000"/>
                        </xsl:if>

                        <!-- 添加亮度 属性 李娟 2012 04.16-->
                        <xsl:if test="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 | 图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020">
                          <a:extLst>
                            <!--<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">-->
                            <a:ext>
                              <xsl:attribute name="uri">
                                <xsl:value-of select="'{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}'"/>
                              </xsl:attribute>
                              <a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
                                <a14:imgLayer>
                                  <!--<xsl:attribute name="r:embed">
																<xsl:value-of select="concat('rId',./@图形引用_C62E)"/>
															</xsl:attribute>-->
                                  <a14:imgEffect>
                                    <a14:brightnessContrast>
                                      <xsl:attribute name="bright">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020 * 1000"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="contrast">
                                        <xsl:value-of select="图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021 * 1000"/>
                                      </xsl:attribute>
                                    </a14:brightnessContrast>
                                  </a14:imgEffect>
                                </a14:imgLayer>
                              </a14:imgProps>
                            </a:ext>
                          </a:extLst>
                        </xsl:if>

                      </xsl:for-each>
                    </a:blip>
                    <!--添加图片剪裁 李娟 2012 04 16-->

                    <xsl:if test="//图:图片属性_801E/图:图片裁剪_8022">
                      <a:srcRect>
                        <xsl:if test="//图:图片属性_801E/图:图片裁剪_8022/图:上_8023">
                          <xsl:attribute name="b">
                            <!--代表高度-->
                            <xsl:value-of  select="round((//图:大小_8060/@长_C604)-(//图:图片属性_801E/图:图片裁剪_8022/图:上_8023+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:下_8024))"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025">
                          <xsl:attribute name="r">
                            <xsl:value-of  select="round((//图:大小_8060/@宽_C605)-(//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:左_8025+//图:图形_8062/图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:图片裁剪_8022/图:右_8026))"/>
                          </xsl:attribute>
                        </xsl:if>
                      </a:srcRect>
                    </xsl:if>
                    <!--<a:srcRect/>-->
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </p:blipFill>
                  <p:spPr>
                    <!--   
                    12.29 黎美秀修改 xpath路径不完整 可能是layout中
                     <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿/演:主体//uof:锚点_C644[@uof:图形引用=$tag]">
                     
                     2908620
                    -->
                    <xsl:for-each select="ancestor::uof:UOF//uof:锚点_C644[@图形引用_C62E=$tag]">
                      <xsl:call-template name="xfrm"/>
                      <xsl:call-template name="prstGeom"/>
                      <xsl:if test=".//图:填充_804C">
                        <xsl:call-template name="fill"/>
                      </xsl:if>
                      <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                        <xsl:call-template name="ln"/>
                      </xsl:if>
                      <!-- 09.10.30 deleted by myx
                      <xsl:if test=".//图:旋转角度!='0.0'">
                        <xsl:call-template name="rot"/>
                      </xsl:if>-->
                    </xsl:for-each>
                  </p:spPr>
                </p:pic>
              </xsl:when>

            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <!--end liuyin 20121211 修改音视频丢失-->

        <!--       2010.3.9 黎美秀 增加文字表的处理       -->
        <xsl:when test=".//图:内容_8043/字:文字表_416C">
          <xsl:call-template name="table">
            <xsl:with-param name="tblobj" select="$obj"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>

          <xsl:choose>
            <!--xsl:when test=".//图:类别_8019='61' or .//图:类别_8019='71' or .//图:类别_8019='74' or .//图:类别_8019='77'"-->
            <xsl:when test=".//图:类别_8019='61' or .//图:类别_8019='71' or .//图:类别_8019='74' or (.//图:类别_8019='64' and .//图:名称_801A='Curved Connector')">
              <p:cxnSp>
                <p:nvCxnSpPr>
                  <p:cNvPr>
                    <xsl:attribute name="id">
                      <xsl:value-of select="substring($obj,4,5)"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="$name"/>
                    </xsl:attribute>
                    <!--2011-4-5罗文甜，增加web文字-->
                    <xsl:if test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                      <xsl:attribute name="descr">
                        <xsl:value-of select=".//图:Web文字_804F"/>
                      </xsl:attribute>
                    </xsl:if>
                  </p:cNvPr>
                  <p:cNvCxnSpPr>
                    <!--2010-12-20罗文甜：增加连接线规则-->
                    <xsl:if test="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027">
                      <xsl:apply-templates select="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027"/>
                    </xsl:if>
                  </p:cNvCxnSpPr>
                  <p:nvPr/>
                </p:nvCxnSpPr>

                <p:spPr>
                  <xsl:if test="@图形引用_C62E=$obj">

                    <xsl:call-template name="xfrm"/>
                  </xsl:if>
                  <xsl:call-template name="prstGeom"/>

                  <xsl:if test=".//图:填充_804C">

                    <xsl:call-template name="fill"/>
                  </xsl:if>
                  <xsl:if test=".//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                    <xsl:call-template name="ln"/>
                  </xsl:if>
                  <!--2011-01-11罗文甜，增加阴影判断-->
                  <xsl:if test=".//图:阴影_8051">
                    <xsl:call-template name="shadow"/>
                  </xsl:if>
                  <!--09.10.30 deleted by myx
                    <xsl:if test=".//图:旋转角度!='0.0' or .//图:翻转">
                      <xsl:call-template name="rot"/>
                    </xsl:if>-->
                </p:spPr>

              </p:cxnSp>
            </xsl:when>

            <xsl:otherwise>

              <p:sp>
                <p:nvSpPr>
                  <p:cNvPr>
                    <xsl:attribute name="id">

                      <xsl:value-of select="substring($obj,4,5)"/>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:choose>
                        <xsl:when test="not($name='')">
                          <xsl:value-of select="$name"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="'Title 2'"/>
                        </xsl:otherwise>
                      </xsl:choose>

                    </xsl:attribute>
                    <!--2011-4-5罗文甜，增加web文字-->
                    <xsl:if test=".//图:Web文字_804F and .//图:Web文字_804F!=''">
                      <xsl:attribute name="descr">
                        <xsl:value-of select=".//图:Web文字_804F"/>
                      </xsl:attribute>
                    </xsl:if>
                  </p:cNvPr>
                  <!--2011-3-9罗文甜，增加判断-->
                  <p:cNvSpPr>
                    <!--2012 05 08 增加对文本框的判断 lj-->
                    <xsl:if test="$name='' and ancestor::演:幻灯片_6C0F">
                      <xsl:attribute  name="txBox">1</xsl:attribute>
                    </xsl:if>
                    <xsl:if test="string-length($name) &gt; 0 or .//图:缩放是否锁定纵横比_8055='1'">
                      <a:spLocks>
                        <xsl:if test="string-length($name) &gt; 0">
                          <xsl:attribute name="noGrp">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test=".//图:缩放是否锁定纵横比_8055='1'">
                          <xsl:attribute name="noChangeAspect">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>
                      </a:spLocks>
                    </xsl:if>
                  </p:cNvSpPr>
                  <p:nvPr>
                    <xsl:if test="string-length($name) &gt; 0">
                      <p:ph>

                        <xsl:variable name="type">
                          <xsl:choose>
                            <!--start liuyin 20130316 修改bug2716需要修复才能打开-->
                            <!--<xsl:when test="$name='object'">obj</xsl:when>-->
                            <xsl:when test="$name='object'">body</xsl:when>
                            <!--start liuyin 20130316 修改bug2716需要修复才能打开-->

                            <xsl:when test="$name='vertical-title' or $name='title'">title</xsl:when>
                            <xsl:when test="$name='vertical-text' or $name='text'">body</xsl:when>

                            <!--start liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->
                            <!--<xsl:when test="$name='centertitle'">liuyin1</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">liuyin2</xsl:when>-->
                            <xsl:when test="$name='centertitle'">title</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">title</xsl:when>
                            <!--end liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->

                            <xsl:when test="$name='date'">dt</xsl:when>
                            <xsl:when test="$name='number'">sldNum</xsl:when>
                            <xsl:when test="$name='footer'">ftr</xsl:when>
                            <xsl:when test="$name='header'">hdr</xsl:when>
                            <xsl:when test="$name='chart'">chart</xsl:when>
                            <xsl:when test="$name='table'">tbl</xsl:when>
                            <xsl:when test="$name='clipart'">clipArt</xsl:when>
                            <xsl:when test="$name='media-clip'">media</xsl:when>
                            <xsl:when test="$name='notes' and @是否显示缩略图_C630='true'">sldImg</xsl:when>
                            <xsl:when test="$name='notes' and not(@是否显示缩略图_C630)">body</xsl:when>
                            <xsl:otherwise>body</xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <!--2010.04.27 myx -->
                        <xsl:variable name="idx">
                          <xsl:choose>
                            <xsl:when test="$name='vertical-title' or $name='title'">-1</xsl:when>
                            <xsl:when test="$name='centertitle'">-1</xsl:when>
                            <xsl:when test="$name='vertical-subtitle' or $name='subtitle'">1</xsl:when>
                            <xsl:when test="$name='date'">2</xsl:when>
                            <xsl:when test="$name='number'">3</xsl:when>
                            <xsl:when test="$name='footer'">4</xsl:when>
                            <xsl:when test="$name='header'">5</xsl:when>
                            <xsl:when test="$name='chart'">6</xsl:when>
                            <xsl:when test="$name='table'">7</xsl:when>
                            <xsl:when test="$name='clipart'">8</xsl:when>
                            <xsl:when test="$name='media-clip'">9</xsl:when>
                            <xsl:when test="$name='notes' and @是否显示缩略图_C630='true'">10</xsl:when>
                            <xsl:when test="$name='notes' and not(@是否显示缩略图_C630)">11</xsl:when>
                            <xsl:when test="$name='object'">12</xsl:when>
                            <xsl:when test="$name='vertical-text' or $name='text'">
                              <xsl:value-of select="13+count(preceding-sibling::*[uof:占位符_C626[@类型_C627]='vertical_text' or @uof:占位符_C626[@类型_C627]='text'])"/>
                            </xsl:when>
                            <xsl:otherwise>15</xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <xsl:attribute name="type">
                          <xsl:value-of select="$type"/>
                        </xsl:attribute>
                        <!--2010.04.27 myx -->
                        <!--为修改amation的fade，paragraph的互操作 李娟 2012 0507-->
                        <xsl:if test="$idx!='-1' and $idx!='13' and $idx!='1'">
                          <xsl:attribute name="idx">
                            <xsl:value-of select="$idx"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="$name='vertical-title' or $name='vertical-text' or $name='vertical-subtitle'">
                          <xsl:attribute name="orient">vert</xsl:attribute>
                        </xsl:if>

                      </p:ph>
                    </xsl:if>
                  </p:nvPr>
                </p:nvSpPr>
                <p:spPr>

                  <xsl:if test="@图形引用_C62E=$obj">
                    <xsl:call-template name="xfrm"/>
                  </xsl:if>
                  <!--2010-12-26罗文甜，增加自定义图形(自定义曲线，自由曲线)-->
                  <xsl:if test=".//图:类别_8019">
                    <xsl:choose>
                      <xsl:when test=".//图:路径_801C">
                        <!--uof2.0用路径来描述关键点坐标 李娟 2012.01.05-->
                        <xsl:call-template name="custGeom"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:call-template name="prstGeom"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <!--增加对艺术字的处理 李娟 2012.02.28-->
                  <xsl:if test=".//图:填充_804C  and not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">

                    <xsl:call-template name="fill"/>
                  </xsl:if>
                  <xsl:if test=".//图:填充_804C and .//图:预定义图形_8018/图:属性_801D/图:艺术字_802D">
                    <a:noFill/>
                  </xsl:if>

                  <!--<xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) and not(./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">-->
                  <xsl:if test="(.//图:线_8057 | .//图:箭头_805D) and not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">

                    <xsl:call-template name="ln"/>
                  </xsl:if>
                  <xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) 
            and (./图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </xsl:if>
                  <!--2011-01-11罗文甜，增加阴影判断-->
                  <xsl:if test=".//图:阴影_8051 and  not(.//图:预定义图形_8018/图:属性_801D/图:艺术字_802D)">
                    <xsl:call-template name="shadow"/>
                  </xsl:if>
                  <!--增加三维效果 李娟 2012。02.22-->
                  <xsl:if test=".//图:三维效果_8061">

                    <xsl:call-template name="图:三维效果_8061"/>
                  </xsl:if>
                </p:spPr>

                <xsl:if test="图:图形_8062/图:文本_803C">
                  <xsl:for-each select="图:图形_8062/图:文本_803C">
                    <xsl:call-template name="txBody"/>
                  </xsl:for-each>
                </xsl:if>
                <!--新增艺术字转换 李娟 2012.02.28-->
                <xsl:if test="图:图形_8062//图:艺术字_802D">
                  <xsl:for-each select="图:图形_8062//图:艺术字_802D">
                    <xsl:call-template name="art"/>
                    <!--<xsl:apply-templates select="图:艺术字_802D" mode="txbody"/>-->
                  </xsl:for-each>
                </xsl:if>
                <!--2010.04.09 myx add 复制母版 body类型的锚点的文本内容到版式 -->
                <xsl:if test="not(/uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062/图:文本_803C/图:内容_8043) and ancestor::规则:页面版式_B652 and (uof:占位符_C626[@类型_C627]='vertical-text' or uof:占位符_C626[@类型_C627]='text' or uof:占位符_C626[@类型_C627]='title') ">

                  <xsl:variable name="layout_id" select="ancestor::规则:页面版式_B652/@标识符_6B0D"/>
                  <xsl:variable name="master" select="//演:母版_6C0D[@页面版式引用_6BEC/text()=$layout_id]"/>
                  <!--<xsl:variable name="master" select="//演:母版[演:页面版式引用/text()=$layout_id]"/>-->
                  <!--注销这段 不是很清楚text（） 什么意思i 李娟 2012.01.04·········-->
                  <xsl:variable name="target_type">
                    <xsl:choose>
                      <xsl:when test="uof:占位符_C626/@类型_C627='vertical-text' or uof:占位符_C626/@类型_C627='text'">text</xsl:when>
                      <xsl:otherwise>title</xsl:otherwise>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:apply-templates select="$master/uof:锚点_C644[uof:占位符_C626/@类型_C627=$target_type]/图:图形_8062/图:文本_803C/图:内容_8043" mode="layout"/>
                </xsl:if>
                <!--xsl:if test="图:图形/@图:其他对象">
            <xsl:call-template name="txBody"/>
          </xsl:if-->
              </p:sp>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="图:图形_8062/@标识符_804B=$obj and 图:图形_8062/@组合列表_8064">

      <xsl:call-template name="grpshape"/>
    </xsl:if>
  </xsl:template>
  <!--end liuyin 20130307 修改幻灯片母版引用相关，转换后需要修复才能打开-->
  
  <xsl:template name="图:三维效果_8061">
    <a:scene3d>
      <a:camera>

        <xsl:attribute name="prst">
          <xsl:if test=".//uof:方向_C63C">
            <xsl:call-template name="scene3dprst"/>
          </xsl:if>
          <xsl:if test="not(.//图:三维效果_8061/uof:方向_C63C)">
            <xsl:value-of select="'orthographicFront'"/>
          </xsl:if>
        </xsl:attribute>
        <!--<xsl:if test="not(uof:角度_C635/uof:x方向_C636=0 and uof:角度_C635/uof:y方向_C637=0)">-->
        <xsl:choose>
          <xsl:when test=".//uof:角度_C635/uof:x方向_C636!=0.0 and .//uof:角度_C635/uof:y方向_C637!=0.0">
            <a:rot>
              <xsl:attribute name="lon">
                <!--<xsl:value-of select=""/>-->
                <xsl:value-of select="round(((.//uof:角度_C635/uof:x方向_C636+360) mod 360 ) * 60000)"/>

              </xsl:attribute>
              <xsl:attribute name="lat">
                <xsl:value-of select="round(((.//uof:角度_C635/uof:y方向_C637+360) mod 360 ) * 60000)"/>
              </xsl:attribute>
              <xsl:attribute name="rev">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
            </a:rot>
          </xsl:when>
          <xsl:otherwise>
            <a:rot>
              <xsl:attribute name="lon">0</xsl:attribute>
              <xsl:attribute name="lat">0</xsl:attribute>
              <xsl:attribute name="rev">0</xsl:attribute>
            </a:rot>
          </xsl:otherwise>
        </xsl:choose>
        <!--lijuan 2012 05 15-->
        <!--<xsl:if test="not(.//uof:角度_C635/uof:x方向_C636=0.0 and .//uof:角度_C635/uof:y方向_C637=0.0)">
					<a:rot>
						<xsl:attribute name="lon">
							-->
        <!--<xsl:value-of select=""/>-->
        <!--
							<xsl:value-of select="((.//uof:角度_C635/uof:x方向_C636+360) mod 360 ) * 60000"/>
							
						</xsl:attribute>
						<xsl:attribute name="lat">
							<xsl:value-of select="((.//uof:角度_C635/uof:y方向_C637+360) mod 360 ) * 60000"/>
						</xsl:attribute>
						<xsl:attribute name="rev">
							<xsl:value-of select="'0'"/>
						</xsl:attribute>
					</a:rot>
				</xsl:if>-->
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">
          <xsl:choose>
            <xsl:when test=".//uof:照明_C638/uof:强度_C63A='bright'">
              <xsl:value-of select="'balanced'"/>
            </xsl:when>
            <xsl:when test=".//uof:照明_C638/uof:强度_C63A='dim'">
              <xsl:value-of select="'balanced'"/>
            </xsl:when>
            <xsl:when test=".//uof:照明_C638/uof:强度_C63A='normal'">
              <xsl:value-of select="'balanced'"/>
            </xsl:when>
            <xsl:otherwise>balanced</xsl:otherwise>
          </xsl:choose>
          <!--lijuan 2012 05 15-->
          <!--<xsl:if test=".//uof:照明_C638/uof:强度_C63A='bright'">
						<xsl:value-of select="'balanced'"/>
					</xsl:if>-->
          <!--<xsl:if test=".//uof:照明_C638/uof:强度_C63A='dim'">
						<xsl:value-of select="'harsh'"/>
					</xsl:if>-->
          <!--<xsl:if test=".//uof:照明_C638/uof:强度_C63A='normal'">
						<xsl:value-of select="'threePt'"/>
					</xsl:if>-->
        </xsl:attribute>
        <xsl:attribute name="dir">
          <xsl:value-of select="'t'"/>
        </xsl:attribute>
        <!--<xsl:if test="not(.//图:三维效果_8061/uof:照明_C638/uof:角度_C639='0.0')">-->
        <xsl:choose>
          <xsl:when test=".//图:三维效果_8061/uof:照明_C638/uof:角度_C639!='0.0'">
            <a:rot>
              <xsl:attribute name="lat">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name="lon">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name ="rev">
                <xsl:value-of select=".//uof:照明_C638/uof:角度_C639"/>
              </xsl:attribute>
            </a:rot>
          </xsl:when>
          <xsl:otherwise>
            <a:rot>
              <xsl:attribute name="lat">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name="lon">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
              <xsl:attribute name ="rev">
                <xsl:value-of select="'0'"/>
              </xsl:attribute>
            </a:rot>
          </xsl:otherwise>
        </xsl:choose>

        <!--</xsl:if>-->

      </a:lightRig>
    </a:scene3d>
    <a:sp3d>
      <!--start 2015.04.20 guoyongbin 修复三维效果 透视转换失败-->
      <xsl:if test=".//图:三维效果_8061/uof:深度_C63B">
        <xsl:attribute name="extrusionH">
          <xsl:value-of select=".//图:三维效果_8061/uof:深度_C63B * 500"/>
        </xsl:attribute>
      </xsl:if>
      <!--end 2015.04.20 guoyongbin 修复三维效果 透视转换失败-->
      <xsl:if test=".//图:三维效果_8061/uof:表面效果_C63E">
        <xsl:attribute name="prstMaterial">
          <xsl:if test=".//图:三维效果_8061/uof:表面效果_C63E='metal'">
            <xsl:value-of select="'metal'"/>
          </xsl:if>
          <xsl:if test=".//图:三维效果_8061/uof:表面效果_C63E='plastic'">
            <xsl:value-of select="'plastic'"/>
          </xsl:if>
          <xsl:if test=".//图:三维效果_8061/uof:表面效果_C63E='matte'">
            <xsl:value-of select="'matte'"/>
          </xsl:if>
          <xsl:if test=".//图:三维效果_8061/uof:表面效果_C63E='wire-frame'">
            <xsl:value-of select="'legacyWireframe'"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:三维效果_8061/uof:颜色_C63F">
        <a:extrusionClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(.//图:三维效果_8061/uof:颜色_C63F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:extrusionClr>
      </xsl:if>
    </a:sp3d>
  </xsl:template>
  <xsl:template name="scene3dprst">
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='none' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'orthographicFront'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'obliqueTopRight'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'obliqueTopLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'perspectiveBelow'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'perspectiveRight'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'perspectiveLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'obliqueBottomLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'isometricLeftDown'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='parallel'">
      <xsl:value-of select="'perspectiveAbove'"/>
    </xsl:if>

    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='none' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'perspectiveFront'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'obliqueTopLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'perspectiveBelow'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-top-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'obliqueTopRight'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'perspectiveRight'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'perspectiveLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom-left' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'obliqueBottomLeft'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'perspectiveAbove'"/>
    </xsl:if>
    <xsl:if test=".//图:三维效果_8061/uof:方向_C63C/uof:角度_C639='to-bottom-right' and .//图:三维效果_8061/uof:方向_C63C/uof:方式_C63D='perspective'">
      <xsl:value-of select="'obliqueBottomRight'"/>
    </xsl:if>
  </xsl:template>
  <!--2011-01-11罗文甜：增加阴影-->
  <xsl:template name="shadow">
    <!--start liuyin 20130302 修改预定义图形阴影没有按照差异文档转-->
    <xsl:if test=".//图:阴影_8051/@是否显示阴影_C61C='true'">
      <a:effectLst>
        <xsl:variable name="x_C606">
          <xsl:value-of select="round(.//图:阴影_8051/uof:偏移量_C61B/@x_C606)"/>
        </xsl:variable>
        <xsl:variable name="y_C607">
          <xsl:value-of select="round(.//图:阴影_8051/uof:偏移量_C61B/@y_C607)"/>
        </xsl:variable>
        <xsl:variable name="长_C604">
          <xsl:value-of select="round(.//图:大小_8060/@长_C604)"/>
        </xsl:variable>
        <a:outerShdw>
          <xsl:if test=".//图:阴影_8051/uof:偏移量_C61B/dist">
            <xsl:choose>
              <xsl:when test=".//图:阴影_8051/@类型_C61D='shaperelative'">
                <xsl:choose>
                  <xsl:when test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="dist">
                      <xsl:value-of select="round((.//图:阴影_8051/uof:偏移量_C61B/dist) div 10)"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="dist">
                      <xsl:value-of select="round((.//图:阴影_8051/uof:偏移量_C61B/dist) div 1.5)"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
                <xsl:if test="substring($x_C606 ,1,1) = '-'">
                  <xsl:if test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'-2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
                <xsl:if test="substring($x_C606 ,1,1) != '-'">
                  <xsl:if test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'-2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
              </xsl:when>
              <xsl:when test=".//图:阴影_8051/@类型_C61D='perspective'">
                <xsl:choose>
                  <xsl:when test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="dist">
                      <xsl:value-of select="round((.//图:阴影_8051/uof:偏移量_C61B/dist) div 10)"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="dist">
                      <xsl:value-of select="round((.//图:阴影_8051/uof:偏移量_C61B/dist) div 1.5)"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
                <xsl:if test="substring($x_C606 ,1,1) = '-'">
                  <xsl:if test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'-2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
                <xsl:if test="substring($x_C606 ,1,1) != '-'">
                  <xsl:if test="$y_C607 &lt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'-2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="$y_C607 &gt; $长_C604">
                    <xsl:attribute name="kx">
                      <xsl:value-of select="'2400000'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="dist">
                  <xsl:value-of select="round(.//图:阴影_8051/uof:偏移量_C61B/dist)"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
          <xsl:if test=".//图:阴影_8051/uof:偏移量_C61B/dir">
            <xsl:attribute name="dir">
              <xsl:value-of select="round(.//图:阴影_8051/uof:偏移量_C61B/dir)"/>
            </xsl:attribute>
          </xsl:if>
          <!--<xsl:choose>-->
          <!--<xsl:when test=".//图:阴影_8051/@类型_C61D='3'or .//图:阴影/@类型_C61D='11'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-1200000'"/>
              </xsl:attribute>
            </xsl:when>-->
          <!--<xsl:when test=".//图:阴影_8051/@类型_C61D='shaperelative'">
              
              <xsl:attribute name="kx">
                <xsl:value-of select="'2400000'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:阴影_8051/@类型_C61D='perspective'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'2400000'"/>
              </xsl:attribute>
            </xsl:when>-->
          <!--<xsl:when test=".//图:阴影_8051/@类型_C61D=''">
              <xsl:attribute name="kx">
                <xsl:value-of select="'800400'"/>
              </xsl:attribute>
            </xsl:when>-->
          <!--<xsl:when test=".//图:阴影_8051/@类型_C61D='single'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-800400'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>-->
          <!--</xsl:choose>-->
          <!--end liuyin 20130302 修改预定义图形阴影没有按照差异文档转-->

          <a:srgbClr>
            <xsl:choose>
              <xsl:when test="contains(.//图:阴影_8051/@颜色_C61E,'#')">
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(.//图:阴影_8051/@颜色_C61E,'#')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test=".//图:阴影_8051/@颜色_C61E='auto' or not(.//图:阴影_8051/@颜色_C61E)">
                <xsl:attribute name="val">
                  <xsl:value-of select="'808080'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:if test=".//图:阴影_8051/@透明度_C61F">
              <a:alpha>
                <xsl:attribute name="val">

                  <!--start liuyin 20130301 修改预定义图形阴影没有按照差异文档转-->
                  <!--<xsl:value-of select="round((.//图:阴影_8051/@透明度_C61F div 256) * 100 * 1000)"/>-->
                  <xsl:value-of select="round((.//图:阴影_8051/@透明度_C61F) * 1000)"/>
                  <!--end liuyin 20130301 修改预定义图形阴影没有按照差异文档转-->

                </xsl:attribute>
              </a:alpha>
            </xsl:if>
          </a:srgbClr>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>
  </xsl:template>
  <!--2010-12-20 罗文甜：增加连接线规则-->
  <xsl:template match="图:连接线规则_8027">
    <xsl:if test="./@始端对象引用_8029 or (./@始端对象连接点索引_802B and ./@始端对象连接点索引_802B!='-1')">
      <a:stCxn>
        <!--liuyangyang 2015-03-19 修改应用对象值为空，oo文件需复的要修问题 start-->
        <xsl:if test="./@始端对象引用_8029 !=''">
          <xsl:attribute name="id">
            <xsl:value-of select="substring-after(./@始端对象引用_8029,'Obj')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@始端对象引用_8029 =''">
          <xsl:attribute name="id">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:if>
        <!--end liuyangyang 2015-03-19 修改应用对象值为空，oo文件需复的要修问题-->
        <xsl:if test="./@始端对象连接点索引_802B and ./@始端对象连接点索引_802B!='-1'">
          <xsl:attribute name="idx">
            <xsl:value-of select="./@始端对象连接点索引_802B"/>
          </xsl:attribute>
        </xsl:if>
      </a:stCxn>
    </xsl:if>
    <xsl:if test="./@终端对象引用_802A or (./@终端对象连接点索引_802C and ./@终端对象连接点索引_802C!='-1')">
      <a:endCxn>
        <!--liuyangyang 2015-03-19 修改应用对象值为空，oo文件需复的要修问题 start-->
        <xsl:if test="./@终端对象引用_802A !=''">
          <xsl:attribute name="id">
            <xsl:value-of select="substring-after(./@终端对象引用_802A,'Obj')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="./@终端对象引用_802A =''">
          <xsl:attribute name="id">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:if>
        <!-- end liuyangyang 2015-03-19 修改应用对象值为空，oo文件需复的要修问题-->
        <xsl:if test="./@终端对象连接点索引_802C and ./@终端对象连接点索引_802C!='-1'">
          <xsl:attribute name="idx">
            <xsl:value-of select="./@终端对象连接点索引_802C"/>
          </xsl:attribute>
        </xsl:if>
      </a:endCxn>
    </xsl:if>
  </xsl:template>

  <xsl:template name="grpshape">
    <xsl:variable name="list">
      <xsl:value-of select="图:图形_8062/@组合列表_8064"/>
    </xsl:variable>
    <xsl:variable name="id">
      <xsl:value-of select="图:图形_8062/@标识符_804B"/>
    </xsl:variable>
    <p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr>
          <xsl:attribute name="id">
            <xsl:value-of select="substring(@图形引用_C62E,4,5)"/>
          </xsl:attribute>
          <xsl:attribute name="name">组合</xsl:attribute>
          <!--2011-4-5罗文甜，增加web文字-->
          <xsl:if test="图:图形_8062//图:Web文字_804F and 图:图形_8062//图:Web文字_804F!=''">
            <xsl:attribute name="descr">
              <xsl:value-of select="图:图形_8062//图:Web文字_804F"/>
            </xsl:attribute>
          </xsl:if>
        </p:cNvPr>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <xsl:if test="@图形引用_C62E=$id">
          <xsl:for-each select="./图:图形_8062[@标识符_804B=$id]">

            <xsl:call-template name="grpXfrm"/>
            <!-- 09.10.30 deleted by myx
            <xsl:if test=".//图:旋转角度!='0.0'">
              <xsl:call-template name="rot"/>
            </xsl:if
            >-->
          </xsl:for-each>
        </xsl:if>
      </p:grpSpPr>

      <xsl:for-each select="./图:图形_8062">
        <xsl:choose>
          <xsl:when test="contains($list,@标识符_804B) and not(@组合列表_8064)">

            <xsl:call-template name="prtshape1"/>
          </xsl:when>
          <xsl:when test="contains($list,@标识符_804B) and @组合列表_8064">

            <xsl:call-template name="grpshape1"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </p:grpSp>
  </xsl:template>
  <xsl:template name="prtshape1">
    <xsl:choose>
      <!--<xsl:when test="./@图:其他对象">-->
      <xsl:when test="./图:其他对象引用_8038 or ./图:图片数据引用_8037">
        <!--xsl:call-template name="pic"/-->
        <xsl:variable name="obj" select="@标识符_804B"/>
        <xsl:variable name="picId">
          <xsl:value-of select="./图:其他对象引用_8038 | ./图:图片数据引用_8037"/>
        </xsl:variable>

        <!--start liuyin 20130112 修改组合图形对象中的图片颜色模式，转换后效果不一致-->
        <xsl:variable name="picPr">
          <xsl:value-of select="./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:颜色模式_801F"/>
        </xsl:variable>
        <xsl:variable name="picPrliangdu">
          <xsl:value-of select="./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:亮度_8020"/>
        </xsl:variable>
        <xsl:variable name="picPrduibidu">
          <xsl:value-of select="./图:预定义图形_8018/图:属性_801D/图:图片属性_801E/图:对比度_8021"/>
        </xsl:variable>
        <!--end liuyin 20130112 修改组合图形对象中的图片颜色模式，转换后效果不一致-->
        <xsl:variable name="tag">
          <xsl:value-of select="./@标识符_804B"/>
        </xsl:variable>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picId]">
          <xsl:choose>
            <xsl:when test="@公共类型_D706='png' or 'jpg' or 'jpeg' or 'gif' or 'tiff' or 'bmp'">
              <p:pic>
                <p:nvPicPr>
                  <p:cNvPr>
                    <xsl:variable name="mediaId" select="./@标识符_D704"/>
                    <xsl:attribute name="id">
                      <xsl:call-template name="mediaIdConvetor">
                        <xsl:with-param name="mediaId" select="$mediaId"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <xsl:attribute name="name">
                      <xsl:value-of select="concat('pic',$mediaId)"/>
                    </xsl:attribute>
                    <xsl:attribute name="descr">
                      <xsl:value-of select="picture"/>
                    </xsl:attribute>
                  </p:cNvPr>
                  <p:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                  </p:cNvPicPr>
                  <p:nvPr/>
                </p:nvPicPr>
                <p:blipFill>
                  <a:blip>
                    <xsl:attribute name="r:embed">
                      <xsl:value-of select="concat('rId',@标识符_D704)"/>
                    </xsl:attribute>

                    <!--start liuyin 20130112 修改组合图形对象中的图片颜色模式，转换后效果不一致-->
                    <xsl:choose>
                      <xsl:when test="$picPr='monochrome'">
                        <a:biLevel thresh="50000"/>
                      </xsl:when>
                      <xsl:when test="$picPr='erosion'">
                        <a:lum bright="70000" contrast="-70000"/>
                      </xsl:when>
                      <xsl:when test="$picPr='greyscale'">
                        <a:grayscl/>
                      </xsl:when>
                      <!--end liuyin 20130112 修改组合图形对象中的图片颜色模式，转换后效果不一致-->

                    </xsl:choose>
                  </a:blip>
                  <a:srcRect/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </p:blipFill>
                <p:spPr>
                  <!--
                  2010.1.10 黎美秀修改 版式中的组合图形位置
                   <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿/演:主体//图:图形[@图:标识符=$tag]">
                  
                  -->
                  <xsl:for-each select="ancestor::uof:UOF/uof:演示文稿//图:图形_8062[@标识符_804B=$tag]">
                    <!--因组合图形位置更改，图形不再位于演示文稿之下 李娟 2012.01.05-->
                    <xsl:choose>
                      <xsl:when test="图:组合位置_803B/@x_C606 or 图:组合位置_803B/@y_C607">
                        <a:xfrm>
                          <xsl:if test=".//图:旋转角度_804D!='0.0'">
                            <xsl:attribute name="rot">
                              <!--2011-5-26 luowentian-->
                              <xsl:choose>
                                <!--<xsl:when test="图:翻转/@图:方向='y' or 图:翻转/@图:方向='x'">-->
                                <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                                  <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="round(.//图:旋转角度_804D*60000)"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="图:翻转_803A">
                            <xsl:call-template name="flip"/>
                          </xsl:if>
                          <!---->
                          <xsl:if test="图:组合位置_803B/@x_C606 or 图:组合位置_803B/@y_C607">
                            <a:off>
                              <xsl:attribute name="x">
                                <xsl:value-of select="round(图:组合位置_803B/@x_C606*12700)"/>
                              </xsl:attribute>
                              <xsl:attribute name="y">
                                <xsl:value-of select="round(图:组合位置_803B/@y_C607*12700)"/>
                              </xsl:attribute>
                            </a:off>
                          </xsl:if>
                          <xsl:if test="图:预定义图形_8018//图:大小_8060/@长_C604 or 图:预定义图形_8018//图:大小_8060/@宽_C605">
                            <a:ext>
                              <xsl:attribute name="cx">
                                <xsl:value-of select="round(图:预定义图形_8018//图:大小_8060/@宽_C605*12700)"/>
                              </xsl:attribute>
                              <xsl:attribute name="cy">
                                <xsl:value-of select="round(图:预定义图形_8018//图:大小_8060/@长_C604*12700)"/>
                              </xsl:attribute>
                            </a:ext>
                          </xsl:if>
                        </a:xfrm>
                      </xsl:when>
                      <xsl:when test="ancestor::uof:锚点_C644/图:图形_8062[@组合列表_8064]">
                        <a:xfrm>
                          <xsl:if test=".//图:旋转角度_804D!='0.0'">
                            <xsl:attribute name="rot">
                              <!--2011-5-26 luowentian-->
                              <xsl:choose>
                                <!--<xsl:when test="图:翻转/@图:方向='y' or 图:翻转/@图:方向='x'">-->
                                <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                                  <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="round(.//图:旋转角度_804D*60000)"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                          </xsl:if>
                          <xsl:if test="图:翻转_803A">
                            <xsl:call-template name="flip"/>
                          </xsl:if>
                          <!---->
                          <xsl:if test="ancestor::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108
                                  or ancestor::uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108">
                            <a:off>
                              <xsl:attribute name="x">
                                <xsl:choose>
                                  <xsl:when test="contains(ancestor::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108,'-')">
                                    <xsl:value-of select="'0'"/>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    <xsl:value-of select="round(ancestor::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108*12700)"/>
                                  </xsl:otherwise>
                                </xsl:choose>
                              </xsl:attribute>
                              <xsl:attribute name="y">
                                <xsl:value-of select="round(ancestor::uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108*12700)"/>
                              </xsl:attribute>
                            </a:off>
                          </xsl:if>
                          <xsl:if test="图:预定义图形_8018//图:大小_8060/@长_C604 or 图:预定义图形_8018//图:大小_8060/@宽_C605">
                            <a:ext>
                              <xsl:attribute name="cx">
                                <xsl:value-of select="round(图:预定义图形_8018//图:大小_8060/@宽_C605*11500)"/>
                              </xsl:attribute>
                              <xsl:attribute name="cy">
                                <xsl:value-of select="round(图:预定义图形_8018//图:大小_8060/@长_C604*11500)"/>
                              </xsl:attribute>
                            </a:ext>
                          </xsl:if>
                        </a:xfrm>
                      </xsl:when>
                    </xsl:choose>
                    
                    <xsl:call-template name="prstGeom"/>
                    <xsl:if test=".//图:填充_804C">
                      <xsl:call-template name="fill"/>
                    </xsl:if>
                    <xsl:if test=".//图:线颜色_8058|图:线类型_8059|图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
                      <xsl:call-template name="ln"/>
                    </xsl:if>
                  </xsl:for-each>
                </p:spPr>
              </p:pic>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test=".//图:类别_8019='61' or .//图:类别_8019='71' or .//图:类别_8019='74' or .//图:类别_8019='77'">
        <p:cxnSp>
          <p:nvCxnSpPr>
            <p:cNvPr>
              <xsl:attribute name="id">
                <xsl:value-of select="substring(@标识符_804B,4,5)"/>
              </xsl:attribute>
              <xsl:attribute name="name">
                <xsl:call-template name="prstName"/>
              </xsl:attribute>
            </p:cNvPr>
            <p:cNvCxnSpPr>
              <!--2010-12-20罗文甜：增加连接线规则-->
              <xsl:if test="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027">
                <xsl:apply-templates select="图:图形_8062/图:预定义图形_8018/图:连接线规则_8027"/>
              </xsl:if>
            </p:cNvCxnSpPr>

            <p:nvPr/>
          </p:nvCxnSpPr>
          <p:spPr>
            <xsl:call-template name="subXfrm"/>
            <a:prstGeom>
              <xsl:if test="图:预定义图形_8018/图:类别_8019|图:名称_801A">
                <xsl:attribute name="prst">
                  <xsl:call-template name="prstName"/>
                </xsl:attribute>
                <a:avLst/>
              </xsl:if>
            </a:prstGeom>
            <xsl:if test=".//图:填充_804C">
              <xsl:call-template name="fill"/>
            </xsl:if>
            <xsl:if test=".//图:线颜色_8058|图:线类型_8059|图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F">
              <xsl:call-template name="ln"/>
            </xsl:if>
            <!-- 09.10.30 deleted by myx
            <xsl:if test=".//图:旋转角度!='0.0'">
              <xsl:call-template name="rot"/>
            </xsl:if>
            -->
          </p:spPr>

        </p:cxnSp>
      </xsl:when>
      <xsl:otherwise>

        <p:sp>
          <p:nvSpPr>
            <p:cNvPr>
              <xsl:attribute name="id">
                <xsl:value-of select="substring(@标识符_804B,4,5)"/>
              </xsl:attribute>
              <xsl:attribute name="name">
                <xsl:call-template name="prstName"/>
              </xsl:attribute>
            </p:cNvPr>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <xsl:call-template name="subXfrm"/>
            <!--2010-12-26罗文甜，增加自定义图形(自定义曲线，自由曲线)-->
            <xsl:if test=".//图:类别_8019 and not(.//图:艺术字_802D)">
              <xsl:choose>
                <xsl:when test=".//图:路径_801C">
                  <xsl:call-template name="custGeom"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="prstGeom"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>

            <!--xsl:call-template name="prstGeom"/-->
            <xsl:if test=".//图:填充_804C and not(.//图:艺术字_802D)">
              <xsl:call-template name="fill"/>
            </xsl:if>
            <xsl:if test="not(.//图:艺术字_802D)">
              <xsl:if test=".//图:线颜色_8058|图:线类型_8059|图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F ">
                <xsl:call-template name="ln"/>
              </xsl:if>
            </xsl:if>

            <!--start liuyin 20130304 修改组合图形对象中的阴影，为组合图形对象列举了uof中所有的预设阴影样式，在转换后的ooxml文档中阴影样式丢失-->
            <xsl:if test=".//图:阴影_8051">
              <xsl:call-template name="shadow"/>
            </xsl:if>
            <!--end liuyin 20130304 修改组合图形对象中的阴影，为组合图形对象列举了uof中所有的预设阴影样式，在转换后的ooxml文档中阴影样式丢失-->

            <!--start liuyin 20130304 修改组合图形对象中的三维效果，uof文档中列举一些三维效果，在转换后的ooxml文档中丢失。-->
            <xsl:if test=".//图:三维效果_8061">
              <xsl:call-template name="图:三维效果_8061"/>
            </xsl:if>
            <!--end liuyin 20130304 修改组合图形对象中的三维效果，uof文档中列举一些三维效果，在转换后的ooxml文档中丢失。-->

            <!-- 09.10.30 deleted by myx
            <xsl:if test=".//图:旋转角度!='0.0'">
              <xsl:call-template name="rot"/>
            </xsl:if>
            -->
          </p:spPr>
          <xsl:if test=".//图:艺术字_802D">
            <xsl:for-each select=".//图:艺术字_802D">
              <xsl:call-template name="art"/>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test="图:文本_803C">
            <xsl:for-each select="图:文本_803C">
              <xsl:call-template name="txBody"/>
            </xsl:for-each>
          </xsl:if>
        </p:sp>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--2010.04.09 myx add -->
  <xsl:template match="图:内容_8043" mode="layout">
    <xsl:call-template name="txBody"/>
  </xsl:template>

  <xsl:template name="grpshape1">
    <xsl:variable name="list1">
      <xsl:value-of select="@组合列表_8064"/>
    </xsl:variable>
    <xsl:variable name="id1">
      <xsl:value-of select="@标识符_804B"/>
    </xsl:variable>
    <p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr>
          <xsl:attribute name="id">
            <xsl:value-of select="substring(@标识符_804B,4,5)"/>
          </xsl:attribute>
          <xsl:attribute name="name">组合</xsl:attribute>
          <!--2011-4-5罗文甜，增加web文字-->
          <xsl:if test="图:图形_8062//图:Web文字_804F and 图:图形_8062//图:Web文字_804F!=''">
            <xsl:attribute name="descr">
              <xsl:value-of select="图:图形_8062//图:Web文字_804F"/>
            </xsl:attribute>
          </xsl:if>
        </p:cNvPr>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <!-- 09.10.30 added by myx -->
          <xsl:if test=".//图:旋转角度_804D!='0.0'">
            <xsl:attribute name="rot">
              <!--2011-5-26 luowentian-->
              <xsl:choose>
                <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                  <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round(.//图:旋转角度_804D*60000)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test=".//图:翻转_803A">
            <xsl:call-template name="flip"/>
          </xsl:if>
          <a:off>
            <xsl:attribute name="x">

              <!--start liuyin 20130315 修改bug_2763，需要修复 -->
              <xsl:choose>
                <xsl:when test="contains(./图:组合位置_803B/@x_C606,'-') ">
                  <xsl:value-of select="'0'"/>
                </xsl:when>
                <xsl:when test="./图:组合位置_803B/@x_C606">
                  <xsl:value-of select="round(./图:组合位置_803B/@x_C606*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end liuyin 20130315 修改bug_2763，需要修复 -->

            </xsl:attribute>
            <xsl:attribute name="y">
              <xsl:choose>
                <xsl:when test="./图:组合位置_803B/@y_C607">
                  <xsl:value-of select="round(./图:组合位置_803B/@y_C607*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:off>
          <a:ext>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:ext>
          <a:chOff>
            <xsl:attribute name="x">

              <!--start liuyin 20130315 修改bug_2763，需要修复 -->
              <xsl:choose>
                <xsl:when test="contains(./图:组合位置_803B/@x_C606,'-') ">
                  <xsl:value-of select="'0'"/>
                </xsl:when>
                <xsl:when test="./图:组合位置_803B/@x_C606">
                  <xsl:value-of select="round(./图:组合位置_803B/@x_C606*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end liuyin 20130315 修改bug_2763，需要修复 -->

            </xsl:attribute>
            <xsl:attribute name="y">
              <xsl:choose>
                <xsl:when test="./图:组合位置_803B/@y_C607">
                  <xsl:value-of select="round(./图:组合位置_803B/@y_C607 * 12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:chOff>
          <a:chExt>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605 * 12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604 * 12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:chExt>
        </a:xfrm>
        <!-- 09.10.30 deleted by myx
        <xsl:if test=".//图:旋转角度!='0.0'">
          <xsl:call-template name="rot"/>
        </xsl:if>
        -->
      </p:grpSpPr>
      <xsl:for-each select="following-sibling::图:图形_8062">
        <xsl:choose>
          <xsl:when test="contains($list1,@标识符_804B) and not(@组合列表_8064)">
            <xsl:call-template name="prtshape1"/>
          </xsl:when>
          <xsl:when test="contains($list1,@标识符_804B) and @组合列表_8064">
            <xsl:call-template name="grpshape1"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </p:grpSp>
  </xsl:template>
  <!--
2010.3.9 黎美秀修改 添加文字表处理
<xsl:template name="xfrm">
-->
  <xsl:template name="subXfrm">
    <a:xfrm>

      <xsl:if test=".//图:旋转角度_804D!='0.0'">

        <xsl:attribute name="rot">
          <!--2011-5-26 luowentian-->
          <xsl:choose>
            <xsl:when test=".//图:翻转_803A='x' or .//图:翻转_803A='y'">

              <!--<xsl:value-of select="round((.//图:属性_801D/图:旋转角度_804D)*800000)"/>-->
              <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
            </xsl:when>
            <!--<xsl:when test=".//图:翻转_803A='y'">
				  <xsl:value-of select="round((.//图:属性_801D/图:旋转角度_804D)*230000)"/>
			  </xsl:when>-->
            <xsl:otherwise>
              <xsl:value-of select="round(.//图:旋转角度_804D)*60000"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:翻转_803A">
        <xsl:call-template name="flip"/>
      </xsl:if>
      <a:off>

        <!--2014-03-20, tangjiang, 修改组合图形位置，修正SmartArt重叠错误 start -->
        <xsl:variable name="xPostion">
          <xsl:choose>
            <xsl:when test="ancestor::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108">
              <xsl:value-of select="ancestor::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="x">

          <!--start liuyin 20130315 修改bug_2763，需要修复；调整组合图形的X位置 -->
          <xsl:choose>
            <xsl:when test="contains(./图:组合位置_803B/@x_C606,'-') ">
              <!--liuyngyang 2015-03-19 修改属性为小数文档需要打开修复的问题 start-->
              <xsl:value-of select="round($xPostion)"/>
              <!--end liuyngyang 2015-03-19 修改属性为小数文档需要打开修复的问题-->
            </xsl:when>
            <xsl:when test="$xPostion > ./图:组合位置_803B/@x_C606">
              <xsl:value-of select="round($xPostion *12700)"/>
            </xsl:when>
            <xsl:when test="./图:组合位置_803B/@x_C606">
              <xsl:value-of select="round(./图:组合位置_803B/@x_C606 *12700)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end liuyin 20130315 修改bug_2763，需要修复 -->
        </xsl:attribute>
        <xsl:attribute name="y">
          <xsl:choose>
            <xsl:when test="./图:组合位置_803B/@y_C607">
              <xsl:value-of select="round(./图:组合位置_803B/@y_C607 * 12700)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </a:off>
      <!--end 2014-03-20, tangjiang -->

      <xsl:choose>
        <xsl:when test="./图:预定义图形_8018/图:属性_801D[图:艺术字_802D]/图:大小_8060">
          <a:ext>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604 * 18000)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*15000)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:ext>
        </xsl:when>
        <xsl:otherwise>
          <a:ext>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604">
                  <xsl:value-of select="round(./图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604 * 12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test="./图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605">
                  <xsl:value-of select="round(图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605*12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:ext>
        </xsl:otherwise>
      </xsl:choose>
    </a:xfrm>
  </xsl:template>
  <xsl:template name="grpXfrm">
    <a:xfrm>

      <!-- 09.10.30 added by myx -->
      <xsl:if test=".//图:旋转角度_804D!='0.0'">
        <xsl:attribute name="rot">
          <!--2011-5-26 luowentian-->
          <xsl:choose>
            <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
              <xsl:value-of select="round((360-.//图:旋转角度_804D)*60000)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="round(.//图:旋转角度_804D*60000)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <!--2011-3-20罗文甜，修改Bug-->
      <xsl:if test=".//图:翻转_803A">
        <xsl:call-template name="flip"/>
      </xsl:if>
      <a:off>
        <xsl:attribute name="x">
          <xsl:choose>
            <xsl:when test="contains(parent::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108,'-')">
              <xsl:value-of select="'0'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="round(parent::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108*12700)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="y">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108 * 12700)"/>
        </xsl:attribute>
      </a:off>
      <a:ext>
        <xsl:attribute name="cx">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:大小_C621/@宽_C605 * 12700)"/>
        </xsl:attribute>
        <xsl:attribute name="cy">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:大小_C621/@长_C604 * 12700)"/>
        </xsl:attribute>
      </a:ext>
      <a:chOff>
        <xsl:attribute name="x">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 * 12700)"/>
        </xsl:attribute>
        <xsl:attribute name="y">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108 * 12700)"/>
        </xsl:attribute>
      </a:chOff>
      <a:chExt>
        <xsl:attribute name="cx">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:大小_C621/@宽_C605 * 12700)"/>
        </xsl:attribute>
        <xsl:attribute name="cy">
          <xsl:value-of select="round(parent::uof:锚点_C644/uof:大小_C621/@长_C604 * 12700)"/>
        </xsl:attribute>
      </a:chExt>
    </a:xfrm>
  </xsl:template>
  <!--2011-1-5罗文甜，增加自定义曲线,关键点坐标-->
  <xsl:template name="custGeom">
    <a:custGeom>
      <a:avLst/>
      <a:rect l="l" t="t" r="r" b="b"/>
      <xsl:if test=".//图:路径_801C">
        <a:pathLst>
          
          <!--2014-05-08, tangjiang, 修复轴节点不一的自定义曲线的转换 start -->
          <xsl:choose>
            <xsl:when test="name(.)='uof:锚点_C644'">
              
              <!--2014-05-05, tangjiang, 修复多路径转换 start -->
              <xsl:for-each select="./图:图形_8062/图:预定义图形_8018/图:路径_801C">
                <a:path>

                  <!--2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 start -->
                  <xsl:if test="./图:路径值_8069/@宽度">
                    <xsl:attribute name="w">
                      <xsl:value-of select="./图:路径值_8069/@宽度"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="./图:路径值_8069/@高度">
                    <xsl:attribute name="h">
                      <xsl:value-of select="./图:路径值_8069/@高度"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="position() &lt; last()">
                    <xsl:attribute name="stroke">
                      <xsl:value-of select="'0'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="position()=last() and last() &gt; 1">
                    <xsl:attribute name="fill">
                      <xsl:value-of select="'none'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:attribute name="extrusionOk">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                  <!--end 2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 -->

                  <xsl:if test=".//图:大小_8001_8060/@宽_C605">
                    <xsl:attribute name="w">
                      <xsl:value-of select="round(.//图:大小_8001_8060/@宽_C605 * 12700)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test=".//图:大小_8001_8060/@长_C604">
                    <xsl:attribute name="h">
                      <xsl:value-of select="round(.//图:大小_8001_8060/@长_C604 * 12700)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:for-each select="./*">
                    <xsl:if test="name(.)='M'">
                      <a:moveTo>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                      </a:moveTo>
                    </xsl:if>
                    <xsl:if test="name(.)='C'">
                      <a:cubicBezTo>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[1]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[1]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[2]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[2]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[3]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[3]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                      </a:cubicBezTo>
                    </xsl:if>
                    <xsl:if test="name(.)='L'">
                      <a:lnTo>
                        <a:pt>
                          <xsl:attribute name="x">

                            <!--start liuyin 20130316 修改bug2774需要修复才能打开-->
                            <xsl:choose>
                              <xsl:when test="contains(pt/@x,'E')" >
                                <xsl:value-of select="round(substring-before( pt/@x,'E') * 12700 div 1.34)"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--<xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>-->
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:choose>
                              <xsl:when test="contains(pt/@y,'E')" >
                                <xsl:value-of select="round(substring-before( pt/@y,'E') * 12700 div 1.34)"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--<xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>-->
                            <!--end liuyin 20130316 修改bug2774需要修复才能打开-->

                          </xsl:attribute>
                        </a:pt>
                      </a:lnTo>
                    </xsl:if>
                  </xsl:for-each>
                </a:path>
              </xsl:for-each>
              <!--end 2014-05-05, tangjiang, 修复多路径转换 -->
              
            </xsl:when>
            <xsl:otherwise>
              
              <!--2014-05-05, tangjiang, 修复多路径转换 start -->
              <xsl:for-each select="./图:预定义图形_8018/图:路径_801C">
                <a:path>

                  <!--2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 start -->
                  <xsl:if test="./图:路径值_8069/@宽度">
                    <xsl:attribute name="w">
                      <xsl:value-of select="./图:路径值_8069/@宽度"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="./图:路径值_8069/@高度">
                    <xsl:attribute name="h">
                      <xsl:value-of select="./图:路径值_8069/@高度"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="position() &lt; last()">
                    <xsl:attribute name="stroke">
                      <xsl:value-of select="'0'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test="position()=last() and last() &gt; 1">
                    <xsl:attribute name="fill">
                      <xsl:value-of select="'none'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:attribute name="extrusionOk">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>
                  <!--end 2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 -->

                  <xsl:if test=".//图:大小_8001_8060/@宽_C605">
                    <xsl:attribute name="w">
                      <xsl:value-of select="round(.//图:大小_8001_8060/@宽_C605 * 12700)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test=".//图:大小_8001_8060/@长_C604">
                    <xsl:attribute name="h">
                      <xsl:value-of select="round(.//图:大小_8001_8060/@长_C604 * 12700)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:for-each select="./*">
                    <xsl:if test="name(.)='M'">
                      <a:moveTo>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                      </a:moveTo>
                    </xsl:if>
                    <xsl:if test="name(.)='C'">
                      <a:cubicBezTo>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[1]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[1]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[2]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[2]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                        <a:pt>
                          <xsl:attribute name="x">
                            <xsl:value-of select="round(pt[3]/@x * 12700 div 1.34)"/>
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:value-of select="round(pt[3]/@y * 12700 div 1.34)"/>
                          </xsl:attribute>
                        </a:pt>
                      </a:cubicBezTo>
                    </xsl:if>
                    <xsl:if test="name(.)='L'">
                      <a:lnTo>
                        <a:pt>
                          <xsl:attribute name="x">

                            <!--start liuyin 20130316 修改bug2774需要修复才能打开-->
                            <xsl:choose>
                              <xsl:when test="contains(pt/@x,'E')" >
                                <xsl:value-of select="round(substring-before( pt/@x,'E') * 12700 div 1.34)"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--<xsl:value-of select="round(pt/@x * 12700 div 1.34)"/>-->
                          </xsl:attribute>
                          <xsl:attribute name="y">
                            <xsl:choose>
                              <xsl:when test="contains(pt/@y,'E')" >
                                <xsl:value-of select="round(substring-before( pt/@y,'E') * 12700 div 1.34)"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--<xsl:value-of select="round(pt/@y * 12700 div 1.34)"/>-->
                            <!--end liuyin 20130316 修改bug2774需要修复才能打开-->

                          </xsl:attribute>
                        </a:pt>
                      </a:lnTo>
                    </xsl:if>
                  </xsl:for-each>
                </a:path>
              </xsl:for-each>
              <!--end 2014-05-05, tangjiang, 修复多路径转换 -->
              
            </xsl:otherwise>
          </xsl:choose>
          <!--end 2014-05-08, tangjiang, 修复轴节点不一的自定义曲线的转换 -->
          
        </a:pathLst>
      </xsl:if>
    </a:custGeom>
  </xsl:template>
  <xsl:template name="prstGeom">
    <!--liuyangyang 2015-04-27 添加控制点转换 start-->
    <a:prstGeom>
      <xsl:attribute name="prst">
        <xsl:choose>
          <xsl:when test=".//图:类别_8019|.//图:名称_801A">
            <xsl:call-template name="prstName"/>
          </xsl:when>
          <xsl:otherwise>rect</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <a:avLst>
        <!--<xsl:if
          test=".//图:名称_801A='Elbow Connector' or .//图:名称_801A='Elbow Arrow Connector' or .//图:名称_801A='Curved Arrow Connector' or .//图:名称_801A='Elbow Double-Arrow Connector' or .//图:名称_801A='Curved Double-Arrow Connector' or .//图:名称_801A='Curved Connector'">
          <a:gd name="adj1" fmla="val 50000"/>
        </xsl:if>-->

        <!--2014-02-24，wudi，预定义图形-控制点的转换，start-->
        <xsl:for-each select=".//图:预定义图形_8018">
          <xsl:choose>
            <!--liuyangyang 添加block arc控制点转换，adj1 adj2值有待改进-->
            <xsl:when test=".//图:名称_801A='Block Arc' and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="tmp">
                <xsl:choose>
                  <xsl:when test="$XC606 &lt; 0">
                    <xsl:value-of select="$XC606 div 11796480 + 1.5"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$XC606 div 11796480 - 0.5"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name="tmp1">
                <xsl:value-of select="round(($tmp + 0.5) * 21600000)"/>
              </xsl:variable>
              <xsl:variable name="tmp2">
                <xsl:value-of select="round(($tmp - 0.5) * 21600000)"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',$tmp1)"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',$tmp2)"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((1 - $YC607 div 10800) *  50000))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--椭圆形标注，圆角矩形标注，矩形标注，云形标注-->
            <xsl:when test="(.//图:名称_801A='Oval Callout' or .//图:名称_801A='Rounded Rectangular Callout' or .//图:名称_801A='Rectangular Callout' or .//图:名称_801A='Cloud Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($XC606 - 10800) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round(($YC607 - 10800) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <xsl:if test=".//图:名称_801A='Rounded Rectangular Callout'">
                <a:gd name="adj3">
                  <xsl:attribute name="fmla">
                    <xsl:value-of select="'val 16667'"/>
                  </xsl:attribute>
                </a:gd>
              </xsl:if>
            </xsl:when>
            <xsl:when test="(.//图:名称_801A='Line Callout1' or .//图:名称_801A='Line Callout1(Accent Bar)' or .//图:名称_801A='Line Callout1(No Border)' or .//图:名称_801A='Line Callout1(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--线形标注1，线形标注1（带强调线），线形标注1（无边框线），线形标注1（带边框和强调线）-->
            <xsl:when test="(.//图:名称_801A='Line Callout2' or .//图:名称_801A='Line Callout2(Accent Bar)' or .//图:名称_801A='Line Callout2(No Border)' or .//图:名称_801A='Line Callout2(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--线形标注2，线形标注2（带强调线），线形标注2（无边框线），线形标注2（带边框和强调线）-->
            <xsl:when test="(.//图:名称_801A='Line Callout3' or .//图:名称_801A='Line Callout3(Accent Bar)' or .//图:名称_801A='Line Callout3(No Border)' or .//图:名称_801A='Line Callout3(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X3C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y3C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($Y3C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($X3C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj5">
                <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj6">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj5">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj5"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj6">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj6"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--线形标注3，线形标注3（带强调线），线形标注3（无边框线），线形标注3（带边框和强调线）-->
            <xsl:when test="(.//图:名称_801A='Line Callout4' or .//图:名称_801A='Line Callout4(Accent Bar)' or .//图:名称_801A='Line Callout4(No Border)' or .//图:名称_801A='Line Callout4(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X3C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y3C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X4C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[4]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y4C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[4]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($Y4C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($X4C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($Y3C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X3C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj5">
                <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj6">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj7">
                <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj8">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj5">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj5"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj6">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj6"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj7">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj7"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj8">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj8"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--十字星，八角星，十六角星，二十四角星，三十二角星-->
            <!--五角星OOX有控制点，UOF无控制点-->
            <!--六角星，七角星转成八角星，十角星，十二角星转成十六角星-->
            <xsl:when test ="(.//图:名称_801A='4-Point Star' or .//图:名称_801A='8-Point Star' or .//图:名称_801A='16-Point Star' or .//图:名称_801A='24-Point Star' or .//图:名称_801A='32-Point Star') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上凸带形-->
            <xsl:when test="(.//图:名称_801A='Up Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((21600 - $YC607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 200000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--前凸带形-->
            <xsl:when test="(.//图:名称_801A='Down Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($YC607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 200000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上凸弯带形-->
            <xsl:when test="(.//图:名称_801A='Curved Up Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((21600 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:choose>
                  <xsl:when test ="$X1C606 = -13500">
                    <xsl:value-of select="100000"/>
                  </xsl:when>
                  <xsl:when test ="$X1C606 = -8099.568">
                    <xsl:value-of select ="25000"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round(($X2C606 - 675) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--前凸弯带形-->
            <xsl:when test="(.//图:名称_801A='Curved Down Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((21600 - $X2C606 - 675) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--竖卷形，横卷形-->
            <xsl:when test ="(.//图:名称_801A='Vertical Scroll' or .//图:名称_801A='Horizontal Scroll') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--波形，双波形-->
            <xsl:when test="(.//图:名称_801A='Wave' or .//图:名称_801A='Double Wave') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round(($YC607 - 10800) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--右箭头，虚尾箭头，燕尾形箭头-->
            <xsl:when test="(.//图:名称_801A='Right Arrow' or .//图:名称_801A='Striped Right Arrow' or .//图:名称_801A='Notched Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左箭头，左右箭头-->
            <xsl:when test="(.//图:名称_801A='Left Arrow' or .//图:名称_801A='Left-Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上箭头-->
            <xsl:when test="(.//图:名称_801A='Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上下箭头-->
            <xsl:when test="(.//图:名称_801A='Up-Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 100000 div 10800))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($YC607 * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--下箭头-->
            <xsl:when test="(.//图:名称_801A='Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--十字箭头，丁字箭头-->
            <xsl:when test="(.//图:名称_801A='Quad Arrow' or .//图:名称_801A='Left-Right-Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 10800))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 50000 div 10800))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 50000 div 10800))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--圆角右箭头，两边有差异-->
            <!--直角双向箭头，直角上箭头-->
            <xsl:when test="(.//图:名称_801A='Left-Up Arrow' or .//图:名称_801A='Bent-Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 50000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 50000 div 10800))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左弧形箭头，显示效果有差异-->
            <xsl:when test="(.//图:名称_801A='Curved Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 125 - $X1C606 div 2) * 200000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 + 250 - $X1C606) * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--右弧形箭头，显示效果有差异-->
            <xsl:when test="(.//图:名称_801A='Curved Left Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 125 - $X1C606 div 2) * 200000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 + 250 - $X1C606) * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--下弧形箭头，显示效果有差异-->
            <xsl:when test="(.//图:名称_801A='Curved Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 5 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 + 10 - $X1C606) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上弧形箭头，显示效果有差异-->
            <xsl:when test="(.//图:名称_801A='Curved Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 5 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((21600 + 10 - $X1C606) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--五边形，燕尾形-->
            <xsl:when test ="(.//图:名称_801A='Pentagon Arrow' or .//图:名称_801A='Chevron Arrow') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--右箭头标注-->
            <xsl:when test="(.//图:名称_801A='Right Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左箭头标注-->
            <xsl:when test="(.//图:名称_801A='Left Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--上箭头标注-->
            <xsl:when test="(.//图:名称_801A='Up Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--下箭头标注-->
            <xsl:when test="(.//图:名称_801A='Down Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左右箭头标注-->
            <xsl:when test="(.//图:名称_801A='Left-Right Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--十字箭头标注-->
            <xsl:when test="(.//图:名称_801A='Quad Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="X1C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y1C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="X2C606">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="Y2C607">
                <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <xsl:variable name="adj3">
                <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
              </xsl:variable>
              <xsl:variable name="adj4">
                <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj3"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj4">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj4"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--等腰三角形，八边形，十字型，缺角矩形，立方体，棱台，太阳形，新月形，双括号，双大括号-->
            <xsl:when test ="(.//图:名称_801A='Isosceles Triangle' or .//图:名称_801A='Octagon' or .//图:名称_801A='Cross' or .//图:名称_801A='Plaque' or .//图:名称_801A='Cube' or .//图:名称_801A='Bevel' or .//图:名称_801A='Sun' or .//图:名称_801A='Moon' or .//图:名称_801A='Double Bracket' or .//图:名称_801A='Double Brace') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--平行四边形，梯形，六边形，同心圆，禁止符-->
            <xsl:when test ="(.//图:名称_801A='Parallelogram' or .//图:名称_801A='Trapezoid' or .//图:名称_801A='Hexagon' or .//图:名称_801A='Donut' or .//图:名称_801A='No Symbol') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $width div 21600 div $height))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
              <xsl:if test=".//图:名称_801A='Hexagon'">
                <a:gd name="vf" fmla="val 115470"/>
              </xsl:if>
            </xsl:when>
            <!--圆柱形-->
            <xsl:when test ="(.//图:名称_801A='Can') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 200000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--折角形-->
            <xsl:when test ="(.//图:名称_801A='Folded Corner') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--笑脸-->
            <xsl:when test ="(.//图:名称_801A='Smiley Face') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round(($XC606 - 16200) * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左中括号，右中括号-->
            <xsl:when test ="(.//图:名称_801A='Left Bracket' or .//图:名称_801A='Right Bracket') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--左大括号，右大括号-->
            <xsl:when test ="(.//图:名称_801A='Left Bracket' or .//图:名称_801A='Right Bracket') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="YC607">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
              </xsl:variable>
              <xsl:variable name="width">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
              </xsl:variable>
              <xsl:variable name="height">
                <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
              </xsl:variable>
              <xsl:variable name="adj1">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
              </xsl:variable>
              <xsl:variable name="adj2">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj1">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj1"/>
                </xsl:attribute>
              </a:gd>
              <a:gd name="adj2">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj2"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
            <!--肘形连接符，肘形箭头连接符，肘形双箭头连接符，曲线连接符，曲线箭头连接符，曲线双箭头连接符-->
            <xsl:when test ="(.//图:名称_801A='Elbow Connector' or .//图:名称_801A='Curved Connector') and ancestor::图:图形_8062/图:控制点_8039">
              <xsl:variable name="XC606">
                <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
              </xsl:variable>
              <xsl:variable name="adj">
                <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
              </xsl:variable>
              <a:gd name="adj">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="$adj"/>
                </xsl:attribute>
              </a:gd>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
        <!--end-->

      </a:avLst>
      <!-- <xsl:if test=".//图:阴影">
        <xsl:call-template name="shadow"/>
      </xsl:if>-->
    </a:prstGeom>
    <!--end -->
    <!--<a:prstGeom>
      --><!--2011-01-11罗文甜，修改Bug--><!--
      <xsl:attribute name="prst">
        <xsl:choose>
          <xsl:when test=".//图:类别_8019|.//图:名称_801A">
            <xsl:call-template name="prstName"/>
          </xsl:when>
          <xsl:otherwise>rect</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <a:avLst>
        <xsl:if test=".//图:名称_801A='Elbow Connector' or .//图:名称_801A='Elbow Arrow Connector' or .//图:名称_801A='Curved Arrow Connector' or .//图:名称_801A='Elbow Double-Arrow Connector' or .//图:名称_801A='Curved Double-Arrow Connector' or .//图:名称_801A='Curved Connector'">
          <a:gd name="adj1" fmla="val 50000"/>
        </xsl:if>
      </a:avLst>
    </a:prstGeom>-->
  </xsl:template>
  <!--处理图形填充-->
  <xsl:template name="solidFill">
    <a:solidFill>
      <a:srgbClr>
        <xsl:attribute name="val">
          <xsl:choose>
            <xsl:when test=".//图:属性_801D/图:填充_804C/图:颜色_8004!='auto'">
              <xsl:variable name="RGB1">
                <xsl:value-of select=".//图:属性_801D/图:填充_804C/图:颜色_8004"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($RGB1,'#')"/>
            </xsl:when>
            <xsl:otherwise>bbe0e3</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test=".//图:属性_801D/图:透明度_8050">
          <a:alpha>
            <xsl:attribute name="val">
              <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
            </xsl:attribute>
          </a:alpha>
        </xsl:if>
      </a:srgbClr>
    </a:solidFill>
  </xsl:template>
  <!--处理线属性-->
  <xsl:template name="ln">
    <!--
    
    2010.4.6 黎美秀修改
    把值取整
     <xsl:value-of select=".//图:属性/图:线粗细*12700"/>
    -->
    <a:ln>
      <xsl:if test=".//图:线粗细_805C&gt;0">
        <xsl:attribute name="w">
          <xsl:value-of select="round(.//图:属性_801D/图:线_8057/图:线粗细_805C*12700)"/>
        </xsl:attribute>
      </xsl:if>
      <!--10.10.8罗文甜 修改线型对应-->
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='single'">
        <xsl:attribute name="cmpd">sng</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='double'">
        <xsl:attribute name="cmpd">dbl</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thin-thick'">
        <xsl:attribute name="cmpd">thickThin</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thick-thin'">
        <xsl:attribute name="cmpd">thinThick</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thick-between-thin'">
        <xsl:attribute name="cmpd">tri</xsl:attribute>
      </xsl:if>
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='none'">
      </xsl:if>
      <xsl:choose>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线颜色_8058">
          <a:solidFill>
            <a:srgbClr>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test=".//图:属性_801D/图:线_8057/图:线颜色_8058!='auto'">
                    <xsl:choose>
                      <xsl:when test=".//图:属性_801D/图:线_8057/图:线颜色_8058='#black'">000000</xsl:when>
                      <xsl:otherwise>
                        <xsl:variable name="RGB">
                          <xsl:value-of select=".//图:属性_801D/图:线_8057/图:线颜色_8058"/>
                        </xsl:variable>
                        <xsl:value-of select="substring-after($RGB,'#')"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test=".//图:属性_801D/图:线_8057/图:线颜色_8058='auto'">000000</xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </a:srgbClr>
          </a:solidFill>
        </xsl:when>
        <xsl:otherwise>
          <!-- liuyin 2012.10.30 文本框颜色有白色修改为无色-->
          <xsl:if test="not(uof:占位符_C626/@类型_C627) and 图:图形_8062/图:文本_803C/@是否为文本框_8046">
            <a:noFill/>
          </xsl:if>
          <!--<xsl:otherwise>
        
            <xsl:if test="not(uof:占位符_C626/@类型_C627) and 图:图形_8062/图:文本_803C/@是否为文本框_8046">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="'FFFFFF'"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>-->
        </xsl:otherwise>
      </xsl:choose>
      <!--10.10.8罗文甜 修改虚实对应-->
      <xsl:choose>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='solid'">
          <a:prstDash>
            <xsl:attribute name="val">solid</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='round-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='square-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash'">
          <a:prstDash>
            <xsl:attribute name="val">dash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash-dot'">
          <a:prstDash>
            <xsl:attribute name="val">dashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='long-dash'">
          <a:prstDash>
            <xsl:attribute name="val">lgDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='long-dash-dot'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash-dot-dot'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDotDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <!--xsl:when test=".//图:属性/图:虚实='dash-dot-heavy'">
          <a:prstDash>
            <xsl:attribute name="val">sysDashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:线型='dash-dot-dot'">
					<a:prstDash>
						<xsl:attribute name="val">lgDashDot</xsl:attribute>
					</a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性/图:虚实='dash-dot-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDashDotDot</xsl:attribute>
          </a:prstDash>
        </xsl:when-->
      </xsl:choose>
      <xsl:if test=".//图:前端箭头_805E">
        <a:headEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!--2011-3-20罗文甜，增添箭头大小-->
          <xsl:choose>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='1'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='2'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='3'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='4'">
              <xsl:attribute name="w">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='6'">
              <xsl:attribute name="w">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='7'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='8'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:前端箭头_805E/图:大小_8001='9'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>
            <!--xsl:when test=".//图:前端箭头/图:大小_8001='5'">
            </xsl:when-->
          </xsl:choose>
        </a:headEnd>
      </xsl:if>
      <xsl:if test=".//图:后端箭头_805F">
        <a:tailEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!--2011-3-20罗文甜，增添箭头大小-->
          <xsl:choose>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='1'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='2'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='3'">
              <xsl:attribute name="w">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='4'">
              <xsl:attribute name="w">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='6'">
              <xsl:attribute name="w">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='7'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'sm'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='8'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'med'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test=".//图:后端箭头_805F/图:大小_8001='9'">
              <xsl:attribute name="w">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
              <xsl:attribute name="len">
                <xsl:value-of select="'lg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>
            <!--xsl:when test=".//图:后端箭头_805F/图:大小_8001='5'">
            </xsl:when-->
          </xsl:choose>
        </a:tailEnd>
      </xsl:if>
    </a:ln>
  </xsl:template>
  <!--处理旋转-->
  <xsl:template name="rot">
    <a:scene3d>
      <a:camera>
        <xsl:attribute name="prst">orthographicFront</xsl:attribute>
        <a:rot>
          <xsl:attribute name="lat">0</xsl:attribute>
          <xsl:attribute name="lon">0</xsl:attribute>
          <xsl:attribute name="rev">
            <xsl:choose>
              <xsl:when test=".//图:属性_801D/图:旋转角度_804D!='0.0'">
                <!--2011-5-26 luowentian-->
                <xsl:choose>
                  <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                    <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="round(.//图:属性_801D/图:旋转角度_804D*60000)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:rot>
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">threePt</xsl:attribute>
        <xsl:attribute name="dir">t</xsl:attribute>
      </a:lightRig>
    </a:scene3d>
  </xsl:template>
  <!--style-->

  <xsl:template name="prstName">
    <xsl:choose>
      <!--xsl:when test="图:图形/@图:其他对象 and .//图:名称_801A='Rectangle'">actionButtonForwardNext</xsl:when-->
      <xsl:when test=".//图:名称_801A='Rectangle' or .//图:类别_8019='11'">rect</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='12'">parallelogram</xsl:when>
      <xsl:when test=".//图:名称_801A='Trapezoid' or .//图:类别_8019='13'">flowChartManualOperation</xsl:when>
      <xsl:when test=".//图:名称_801A='diamond' or .//图:类别_8019='14'">diamond</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangle' or .//图:类别_8019='15'">roundRect</xsl:when>
      <xsl:when test=".//图:名称_801A='Octagon'or .//图:类别_8019='16'">octagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Isosceles Triangle' or .//图:类别_8019='17'">triangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Triangle' or .//图:类别_8019='18'">rtTriangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval' or .//图:类别_8019='19'">ellipse</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Arrow' or .//图:类别_8019='21'">rightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow' or .//图:类别_8019='22'">leftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow' or .//图:类别_8019='23'">upArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow' or .//图:类别_8019='24'">downArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow' or .//图:类别_8019='25'">leftRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up-Down Arrow' or .//图:类别_8019='26'">upDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Quad Arrow' or .//图:类别_8019='27'">quadArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right-Up Arrow' or .//图:类别_8019='28'">leftRightUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent Arrow' or .//图:类别_8019='29'">bentArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Process' or .//图:类别_8019='31'">flowChartProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Alternate Process' or .//图:类别_8019='32'">flowChartAlternateProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Decision' or .//图:类别_8019='33'">flowChartDecision</xsl:when>
      <xsl:when test=".//图:名称_801A='Data' or .//图:类别_8019='34'">flowChartInputOutput</xsl:when>
      <!--      
      2010.4.26 黎美秀修改
         <xsl:when test=".//图:名称_801A='Predefined Process' or .//图:类别_8019='35'">flowChartProcess</xsl:when>
      -->
      <xsl:when test=".//图:名称_801A='Predefined Process' or .//图:类别_8019='35'">flowChartPredefinedProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Internal Storage' or .//图:类别_8019='36'">flowChartInternalStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Document' or .//图:类别_8019='37'">flowChartDocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Multidocument' or .//图:类别_8019='38'">flowChartMultidocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Terminator' or .//图:类别_8019='39'">flowChartTerminator</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 1' or .//图:类别_8019='41'">irregularSeal1</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 2' or .//图:类别_8019='42'">irregularSeal2</xsl:when>
      <xsl:when test=".//图:名称_801A='4-Point Star' or .//图:类别_8019='43'">star4</xsl:when>
      <xsl:when test=".//图:名称_801A='5-Point Star' or .//图:类别_8019='44'">star5</xsl:when>
      <xsl:when test=".//图:名称_801A='8-Point Star' or .//图:类别_8019='45'">star8</xsl:when>
      <xsl:when test=".//图:名称_801A='16-Point Star' or .//图:类别_8019='46'">star12</xsl:when>
      <xsl:when test=".//图:名称_801A='24-Point Star' or .//图:类别_8019='47'">star24</xsl:when>
      <xsl:when test=".//图:名称_801A='32-Point Star' or .//图:类别_8019='48'">star32</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Ribbon' or .//图:类别_8019='49'">ribbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Rectangular Callout' or .//图:类别_8019='51'">wedgeRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangular Callout' or .//图:类别_8019='52'">wedgeRoundRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval Callout' or .//图:类别_8019='53'">wedgeEllipseCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='54'">cloudCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2' or .//图:类别_8019='56'">borderCallout1</xsl:when>
      <!--4月13日线性标注3-->
      <xsl:when test=".//图:名称_801A='Line Callout3' or .//图:类别_8019='57'">borderCallout2</xsl:when>
      <!---->
      <xsl:when test=".//图:名称_801A='Line Callout4' or .//图:类别_8019='58'">borderCallout3</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Line Callout1(Accent Bar)' or .//图:类别_8019='59'">accentCallout1</xsl:when-->
      <xsl:when test=".//图:名称_801A='Line Callout2(Accent Bar)' or .//图:类别_8019='510'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Accent Bar)' or .//图:类别_8019='511'">accentCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Accent Bar)' or .//图:类别_8019='512'">accentCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1' or .//图:类别_8019='55'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(No Border)' or .//图:类别_8019='514'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(No Border)' or .//图:类别_8019='515'">callout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(No Border)' or .//图:类别_8019='516'">callout3</xsl:when>

      <!--2011-2-17罗文甜，修改uof中四个无对应标注,513,517对应borderCallout1。55,59对应callout1-->
      <xsl:when test=".//图:名称_801A='Line Callout1' or .//图:类别_8019='55'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Accent Bar)' or .//图:类别_8019='59'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(No Border)' or .//图:类别_8019='513'">borderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Border and Accent Bar)' or .//图:类别_8019='517'">borderCallout1</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Line Callout1(Border and Accent Bar)' or .//图:类别_8019='517'">accentBorderCallout1</xsl:when-->

      <xsl:when test=".//图:名称_801A='Line Callout2(Border and Accent Bar)' or .//图:类别_8019='518'">accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Border and Accent Bar)' or .//图:类别_8019='519'">accentBorderCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Border and Accent Bar)' or .//图:类别_8019='520'">accentBorderCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line' or .//图:类别_8019='61'">line</xsl:when>
      <xsl:when test=".//图:名称_801A='Arrow' or .//图:类别_8019='62'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Arrow' or .//图:类别_8019='63'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Curve' or .//图:类别_8019='64'">curvedConnector3</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Freeform' or .//图:类别_8019='65'">任意多边形 8</xsl:when>
                    <xsl:when test=".//图:名称_801A='Scribble' or .//图:类别_8019='66'">任意多边形 9</xsl:when-->
      <xsl:when test=".//图:名称_801A='Straight Connector' or .//图:类别_8019='71'">straightConnector1</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Straight Arrow Connector' or .//图:类别_8019='72'">straightConnector1</xsl:when>
                    <xsl:when test=".//图:名称_801A='Straight Double-Arrow Connector' or .//图:类别_8019='73'">straightConnector1</xsl:when-->
      <xsl:when test=".//图:名称_801A='Elbow Connector' or .//图:类别_8019='74'">bentConnector3</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Elbow Arrow Connector' or .//图:类别_8019='75'">bentConnector3</xsl:when>
                    <xsl:when test=".//图:名称_801A='Elbow Double-Arrow Connector' or .//图:类别_8019='76'">bentConnector3</xsl:when-->
      <xsl:when test=".//图:名称_801A='Curved Connector' or .//图:类别_8019='77'">curvedConnector3</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Curved Arrow Connector' or .//图:类别_8019='78'">curvedConnector3</xsl:when>
                    <xsl:when test=".//图:名称_801A='Curved Double-Arrow Connector' or .//图:类别_8019='79'">curvedConnector3</xsl:when-->
      <xsl:when test=".//图:名称_801A='Hexagon' or .//图:类别_8019='110'">hexagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Cross' or .//图:类别_8019='111'">plus</xsl:when>
      <xsl:when test=".//图:名称_801A='Regual Pentagon' or .//图:类别_8019='112'">pentagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Can' or .//图:类别_8019='113'">can</xsl:when>
      <xsl:when test=".//图:名称_801A='Cube' or .//图:类别_8019='114'">cube</xsl:when>
      <xsl:when test=".//图:名称_801A='Bevel' or .//图:类别_8019='115'">bevel</xsl:when>
      <xsl:when test=".//图:名称_801A='Folded Corner' or .//图:类别_8019='116'">foldedCorner</xsl:when>
      <xsl:when test=".//图:名称_801A='Smiley Face' or .//图:类别_8019='117'">smileyFace</xsl:when>
      <xsl:when test=".//图:名称_801A='Donut' or .//图:类别_8019='118'">donut</xsl:when>
      <xsl:when test=".//图:名称_801A='No Symbol' or .//图:类别_8019='119'">noSmoking</xsl:when>
      <xsl:when test=".//图:名称_801A='Block Arc' or .//图:类别_8019='120'">blockArc</xsl:when>
      <xsl:when test=".//图:名称_801A='Heart' or .//图:类别_8019='121'">heart</xsl:when>
      <xsl:when test=".//图:名称_801A='Lightning' or .//图:类别_8019='122'">lightningBolt</xsl:when>
      <xsl:when test=".//图:名称_801A='Sun' or .//图:类别_8019='123'">sun</xsl:when>
      <xsl:when test=".//图:名称_801A='Moon' or .//图:类别_8019='124'">moon</xsl:when>
      <xsl:when test=".//图:名称_801A='Arc' or .//图:类别_8019='125'">arc</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Bracket' or .//图:类别_8019='126'">bracketPair</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Brace' or .//图:类别_8019='127'">bracePair</xsl:when>
      <xsl:when test=".//图:名称_801A='Plaque' or .//图:类别_8019='128'">plaque</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Bracket' or .//图:类别_8019='129'">leftBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Bracket' or .//图:类别_8019='130'">rightBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Brace' or .//图:类别_8019='131'">leftBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Brace' or .//图:类别_8019='132'">rightBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='U-Turn Arrow' or .//图:类别_8019='210'">uturnArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Up Arrow' or .//图:类别_8019='211'">leftUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent-Up Arrow' or .//图:类别_8019='212'">bentUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Right Arrow' or .//图:类别_8019='213'">curvedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Left Arrow' or .//图:类别_8019='214'">curvedLeftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Arrow' or .//图:类别_8019='215'">curvedUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Arrow' or .//图:类别_8019='216'">curvedDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Striped Right Arrow' or .//图:类别_8019='217'">stripedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Notched Right Arrow' or .//图:类别_8019='218'">notchedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Pentagon Arrow' or .//图:类别_8019='219'">homePlate</xsl:when>
      <xsl:when test=".//图:名称_801A='Chevron Arrow' or .//图:类别_8019='220'">chevron</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Arrow Callout' or .//图:类别_8019='221'">rightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow Callout' or .//图:类别_8019='222'">leftArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow Callout' or .//图:类别_8019='223'">upArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow Callout' or .//图:类别_8019='224'">downArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow Callout' or .//图:类别_8019='225'">leftRightArrowCallout</xsl:when>
      <!--20110217Luowentian:箭头bug，软件中没有，但是转换过去软件可以识别-->
      <xsl:when test=".//图:名称_801A='Up-Down Arrow Callout' or .//图:类别_8019='226'">upDownArrowCallout</xsl:when>
      <!--xsl:when test=".//图:名称_801A='Up-Down Arrow Callout' or .//图:类别_8019='226'">upDownArrowCallout</xsl:when-->
      <xsl:when test=".//图:名称_801A='Quad Arrow Callout' or .//图:类别_8019='227'">quadArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Circular Arrow' or .//图:类别_8019='228'">circularArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Preparation' or .//图:类别_8019='310'">flowChartPreparation</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Input' or .//图:类别_8019='311'">flowChartManualInput</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Operation' or .//图:类别_8019='312'">flowChartManualOperation</xsl:when>
      <xsl:when test=".//图:名称_801A='Connector' or .//图:类别_8019='313'">flowChartConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Off-page Connector' or .//图:类别_8019='314'">flowChartOffpageConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Card' or .//图:类别_8019='315'">flowChartPunchedCard</xsl:when>
      <xsl:when test=".//图:名称_801A='Punched Tape' or .//图:类别_8019='316'">flowChartPunchedTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Summing Junction' or .//图:类别_8019='317' ">flowChartSummingJunction</xsl:when>
      <xsl:when test=".//图:名称_801A='Or' or .//图:类别_8019='318'">flowChartOr</xsl:when>
      <xsl:when test=".//图:名称_801A='Collate' or .//图:类别_8019='319'">flowChartCollate</xsl:when>
      <xsl:when test=".//图:名称_801A='Sort' or .//图:类别_8019='320'">flowChartSort</xsl:when>
      <xsl:when test=".//图:名称_801A='Extract' or .//图:类别_8019='321'">flowChartExtract</xsl:when>
      <xsl:when test=".//图:名称_801A='Merge' or .//图:类别_8019='322'">flowChartMerge</xsl:when>
      <xsl:when test=".//图:名称_801A='Stored Data' or .//图:类别_8019='323'">flowChartOnlineStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Delay' or .//图:类别_8019='324'">flowChartDelay</xsl:when>
      <xsl:when test=".//图:名称_801A='Sequential Access Storage' or .//图:类别_8019='325'">flowChartMagneticTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Magnetic Disk' or .//图:类别_8019='326'">flowChartMagneticDisk</xsl:when>
      <xsl:when test=".//图:名称_801A='Direct Access Storage' or .//图:类别_8019='327'">flowChartMagneticDrum</xsl:when>
      <xsl:when test=".//图:名称_801A='Display' or .//图:类别_8019='328'">flowChartDisplay</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Ribbon' or .//图:类别_8019='410'">ribbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Ribbon' or .//图:类别_8019='411'">ellipseRibbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Ribbon' or .//图:类别_8019='412'">ellipseRibbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Vertical Scroll' or .//图:类别_8019='413'">verticalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Horizontal Scroll' or .//图:类别_8019='414'">horizontalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Wave' or .//图:类别_8019='415'">wave</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Wave' or .//图:类别_8019='416'">doubleWave</xsl:when>
      <xsl:otherwise>rect</xsl:otherwise>
      <!--李娟添加 2012.02.28 预定义图形部分 默认转为矩形-->
    </xsl:choose>
  </xsl:template>
  <!--添加艺术字李娟 2012.02.29-->
  <!--<xsl:template match="图:艺术字_802D" mode="txbody">-->
  <xsl:template name="art">
    <p:txBody>
      <a:bodyPr>
        <xsl:if test="图:是否竖排文字_8033='true'">
          <xsl:attribute name="vert">eaVert</xsl:attribute>
        </xsl:if>

        <!--start liuyin 20130301 修改组合图形对象中的艺术字转换后效果不一致-->
        <xsl:if test="./图:艺术字文本_8036">
          <a:prstTxWarp prst="textDeflate"/>
        </xsl:if>
        <!--end liuyin 20130301 修改组合图形对象中的艺术字转换后效果不一致-->

        <a:spAutoFit/>
        <xsl:if test="../图:三维效果_8061">
          <xsl:for-each select="../图:三维效果_8061">

            <xsl:call-template name="图:三维效果_8061">

            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </a:bodyPr>
      <a:p>
        <a:pPr>
          <xsl:attribute name="algn">
            <xsl:choose>
              <xsl:when test="图:对齐方式_8031='left'">l</xsl:when>
              <xsl:when test="图:对齐方式_8031='right'">r</xsl:when>
              <xsl:when test="图:对齐方式_8031='center'">ctr</xsl:when>
              <xsl:when test="图:对齐方式_8031='justified'">just</xsl:when>
              <xsl:when test="图:对齐方式_8031='distributed'">dist</xsl:when>
              <xsl:otherwise>ctr</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:pPr>
        <a:r>
          <a:rPr>
            <xsl:if test="图:字体_802E/@中文字体引用_412A">
              <xsl:attribute name="lang">
                <xsl:variable name="temp" select="图:字体_802E/@中文字体引用_412A"/>
                <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="图:字体_802E/@西文字体引用_4129">
              <xsl:attribute name="altLang">
                <xsl:variable name="temp" select="图:字体_802E/@西文字体引用_4129"/>
                <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="sz">
              <xsl:choose>
                <xsl:when test="图:字体_802E/@字号_412D">
                  <xsl:value-of select="round(图:字体_802E/@字号_412D)*100"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'18*100'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <xsl:if test="../图:线_8057">
              <xsl:call-template name="artln"/>
            </xsl:if>
            <xsl:if test="../图:填充_804C">
              <xsl:for-each select="../图:填充_804C">
                <xsl:call-template name="artsolidFill">
                </xsl:call-template>
              </xsl:for-each>
            </xsl:if>

            <xsl:if test="../图:阴影_8051">
              <xsl:call-template name="artshadow">
              </xsl:call-template>
            </xsl:if>

          </a:rPr>
          <a:t>
            <xsl:value-of select="图:艺术字文本_8036"/>
          </a:t>
        </a:r>



      </a:p>
    </p:txBody>
  </xsl:template>
  <xsl:template name="artln">
    <a:ln>
      <xsl:if test="..//图:线粗细_805C&gt;0">
        <xsl:attribute name="w">
          <xsl:value-of select="round(..//图:线粗细_805C*12700)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="..//图:线类型_8059">
        <xsl:attribute name="cmpd">
          <xsl:choose>
            <xsl:when test="..//图:线类型_8059/@线型_805A='single'">sng</xsl:when>
            <xsl:when test="..//图:线类型_8059/@线型_805A='double'">dbl</xsl:when>
            <xsl:when test="..//图:线类型_8059/@线型_805A='thin-thick'">thickThin</xsl:when>
            <xsl:when test="..//图:线类型_8059/@线型_805A='thick-thin'">thinThick</xsl:when>
            <xsl:when test="..//图:线类型_8059/@线型_805A='thick-between-thin'">tri</xsl:when>
            <xsl:when test="..//图:线类型_8059/@线型_805A='none'"></xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="..//图:线颜色_8058">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:choose>
                <xsl:when test="..//图:线颜色_8058!='auto'">
                  <xsl:variable name="RGB">
                    <xsl:value-of select="..//图:线颜色_8058"/>
                  </xsl:variable>
                  <xsl:value-of select="substring-after($RGB,'#')"/>
                </xsl:when>
                <xsl:when test=".//图:线颜色_8058='auto'">000000</xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:if>
    </a:ln>
  </xsl:template>
  <xsl:template name="artsolidFill">
    <xsl:if test="图:颜色_8004 and substring(图:颜色_8004,2,6)!='ffffff'">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name="val">
            <xsl:choose>
              <xsl:when test="图:颜色_8004!='auto'">
                <xsl:variable name="RGB1">
                  <xsl:value-of select="图:颜色_8004"/>
                </xsl:variable>
                <xsl:value-of select="substring-after($RGB1,'#')"/>
              </xsl:when>
              <xsl:otherwise>bbe0e3</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:if test="..//图:透明度_8050">
            <a:alpha>
              <xsl:attribute name="val">
                <xsl:value-of select="concat((100-..//图:透明度_8050),'%')"/>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <xsl:if test="图:渐变_800D">
      <xsl:variable name="angle" select="图:渐变_800D/@渐变方向_8013"/>
      <a:gradFill flip="none" rotWithShape="0">
        <a:gsLst>
          <xsl:choose>
            <xsl:when test="图:渐变_800D/@种子类型_8010='radar'">
              <a:gs pos="0%">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                  <!--2011-4-4罗文甜-->
                  <xsl:if test="../图:属性_801D/图:透明度_8050">
                    <a:alpha>
                      <xsl:attribute name="val">
                        <xsl:value-of select="concat((100-../图:属性_801D/图:透明度_8050),'%')"/>
                      </xsl:attribute>
                    </a:alpha>
                  </xsl:if>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="50%">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(图:渐变_800D/@终止色_800F,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100%">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
                  <a:gs pos="100%">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                      <!--2011-4-4罗文甜-->
                      <xsl:if test="../图:属性_801D/图:透明度_8050">
                        <a:alpha>
                          <xsl:attribute name="val">
                            <xsl:value-of select="concat((100-../图:属性_801D/图:透明度_8050),'%')"/>
                          </xsl:attribute>
                        </a:alpha>
                      </xsl:if>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="0%">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(../图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                </xsl:when>
                <xsl:otherwise>
                  <a:gs pos="0%">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(/图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                      <!--2011-4-4罗文甜-->
                      <xsl:if test="../图:属性_801D/图:透明度_8050">
                        <a:alpha>
                          <xsl:attribute name="val">
                            <xsl:value-of select="concat((100-../图:属性_801D/图:透明度_8050),'%')"/>
                          </xsl:attribute>of 
                        </a:alpha>
                      </xsl:if>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="100">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </a:gsLst>
        <xsl:choose>
          <xsl:when test="图:渐变_800D/@种子类型_8010='square'">
            <a:path>
              <!--<xsl:attribute name="path">shape</xsl:attribute>-->
              <xsl:attribute name="path">rect</xsl:attribute>
              <xsl:variable name="x" select="图:渐变_800D/@种子X位置_8015"/>
              <xsl:variable name="y" select="图:渐变_800D/@种子Y位置_8016"/>
              <!--<xsl:choose>
                <xsl:when test =".//图:渐变/@图:种子X位置='100'and .//图:渐变/@图:种子Y位置='100'">
                  <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                </xsl:when>
                <xsl:when test =".//图:渐变/@图:种子X位置='30'and .//图:渐变/@图:种子Y位置='30'">
                  <a:fillToRect l="0" t="0" r="85000" b="85000"/>
                </xsl:when>
                <xsl:otherwise>
                  <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                </xsl:otherwise>
              </xsl:choose>-->
              <xsl:choose>
                <xsl:when test="$x='30' and $y='30'">
                  <a:fillToRect r="100%" b="100%"/>
                </xsl:when>
                <xsl:when test="$x='30' and $y='60'">
                  <a:fillToRect t="100%" r="100%"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='30'">
                  <a:fillToRect l="100%" b="100%"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='60'">
                  <a:fillToRect l="100%" t="100%"/>
                </xsl:when>
              </xsl:choose>
            </a:path>
          </xsl:when>
          <xsl:otherwise>
            <a:lin>
              <xsl:attribute name="ang">
                <!--4月16日jjy改渐变颜色-->
                <xsl:choose>
                  <xsl:when test="$angle='0'">
                    <xsl:value-of select="90*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='45'">
                    <xsl:value-of select="45*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='90'">
                    <xsl:value-of select="0*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='135'">
                    <xsl:value-of select="315*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='180'">
                    <xsl:value-of select="270*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='225'">
                    <xsl:value-of select="225*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='270'">
                    <xsl:value-of select="180*60000"/>
                  </xsl:when>
                  <xsl:when test="$angle='315'">
                    <xsl:value-of select="135*60000"/>
                  </xsl:when>
                </xsl:choose>

              </xsl:attribute>
              <xsl:attribute name="scaled">1</xsl:attribute>
            </a:lin>
          </xsl:otherwise>
        </xsl:choose>
        <a:tileRect/>
      </a:gradFill>
    </xsl:if>
  </xsl:template>
  <xsl:template name="artshadow">
    <xsl:if test="..//图:阴影_8051/@是否显示阴影_C61C='true'">
      <a:effectLst>
        <a:outerShdw>
          <xsl:if test="..//图:阴影_8051/uof:偏移量_C61B/dist">
            <xsl:attribute name="dist">
              <xsl:value-of select="round(..//图:阴影_8051/uof:偏移量_C61B/dist)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="..//图:阴影_8051/uof:偏移量_C61B/dir">
            <xsl:attribute name="dir">
              <xsl:value-of select="round(..//图:阴影_8051/uof:偏移量_C61B/dir)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:choose>
            <xsl:when test="..//图:阴影_8051/@类型_C61D='3'or .//图:阴影/@类型_C61D='11'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-1200000'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="..//图:阴影_8051/@类型_C61D='shaperelative'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'1200000'"/>
              </xsl:attribute>
            </xsl:when>
            <!--<xsl:when test=".//图:阴影_8051/@类型_C61D=''">
              <xsl:attribute name="kx">
                <xsl:value-of select="'800400'"/>
              </xsl:attribute>
            </xsl:when>-->
            <xsl:when test="..//图:阴影_8051/@类型_C61D='single'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-800400'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>
          </xsl:choose>
          <a:srgbClr>
            <xsl:choose>
              <xsl:when test="contains(..//图:阴影_8051/@颜色_C61E,'#')">
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(..//图:阴影_8051/@颜色_C61E,'#')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="..//图:阴影_8051/@颜色_C61E='auto' or not(..//图:阴影_8051/@颜色_C61E)">
                <xsl:attribute name="val">
                  <xsl:value-of select="'808080'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:if test="..//图:阴影_8051/@透明度_C61F">
              <a:alpha>
                <xsl:attribute name="val">
                  <xsl:value-of select="concat(round((..//图:阴影_8051/@透明度_C61F div 256) * 100),'%')"/>
                </xsl:attribute>
              </a:alpha>
            </xsl:if>
          </a:srgbClr>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>
  </xsl:template>

  <!--2014-03-20, tangjiang, 添加媒体标识符转换方法，修复标识符重复引起的文件需要修复 start -->
  <xsl:template name="mediaIdConvetor">
    <xsl:param name="mediaId" select="'silde100rId100'"/>
    <xsl:choose>
      <!--5.0以前的标识符转换方式，UOF->OOX采用的处理方式-->
      <xsl:when test="contains($mediaId,'Obj')">
        <xsl:value-of select="substring-after($mediaId,'Obj')"/>
      </xsl:when>
      <xsl:when test="contains($mediaId,'image')">
        <xsl:value-of select="100+substring-after($mediaId,'image')"/>
      </xsl:when>
      <xsl:when test="contains($mediaId,'chart')">
        <xsl:value-of select="400+substring-after($mediaId,'chart')"/>
      </xsl:when>
      <!--5.0及以后的标识符转换方式，互操作采用的处理方式-->
      <xsl:when test="contains($mediaId,'slideLayout')">
        <xsl:value-of select="200+concat(substring-before(substring-after($mediaId,'slideLayout'),'rId'),substring-after($mediaId,'rId'))"/>
      </xsl:when>
      <xsl:when test="contains($mediaId,'slideMaster')">
        <xsl:value-of select="400+concat(substring-before(substring-after($mediaId,'slideMaster'),'rId'),substring-after($mediaId,'rId'))"/>
      </xsl:when>
      <xsl:when test="contains($mediaId,'slide')">
        <xsl:value-of select="600+concat(substring-before(substring-after($mediaId,'slide'),'rId'),substring-after($mediaId,'rId'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$mediaId"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- end 2014-03-20, tangjiang, 添加媒体标识符转换方法，修复标识符重复引起的文件需要修复 -->
</xsl:stylesheet>
