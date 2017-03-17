<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:图形="http://schemas.uof.org/cn/2009/graphics">
 
  <!--start liuyin 20121211 修改音视频丢失-->
  <!--修复图形和音频文件在一个幻灯片的情况下音频不能正确转换  liqiuling 2013-03-29 -->
  <!--200+substring-after(图:图形_8062/图:其他对象引用_8038,'media') 音频和图片文件编号冲突-->

  <!--2014-03-20, tangjiang, 引入媒体标识符转换方法，并修改本文件中所有标识符的转换方式 start -->
  <xsl:import href="cSld.xsl"/>
  <!--end 2014-03-20, tangjiang, 引入媒体标识符转换方法-->
  
  <xsl:template name="animation_other">
    <xsl:if test="演:动画_6B1A/演:序列_6B1B">
      <p:timing>
        <p:tnLst>
          <p:par>
            <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
              <p:childTnLst>
                <xsl:for-each select="uof:锚点_C644[图:图形_8062/图:其他对象引用_8038]">
                    <p:seq concurrent="1" nextAc="seek">
                      <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                        <xsl:if test="../演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34">
                          <p:stCondLst>
                            <p:cond evt="onClick" delay="0">
                              <p:tgtEl>
                                <p:spTgt>
                                  <xsl:variable name="mediaId" select="./图:图形_8062/图:其他对象引用_8038"/>
                                  <xsl:attribute name="spid">
                                    <xsl:call-template name="mediaIdConvetor">
                                      <xsl:with-param name="mediaId" select="$mediaId"/>
                                    </xsl:call-template>
                                  </xsl:attribute>
                                </p:spTgt>
                              </p:tgtEl>
                            </p:cond>
                          </p:stCondLst>
                          <p:endSync evt="end" delay="0">
                            <p:rtn val="all"/>
                          </p:endSync>
                        </xsl:if>

                        <p:childTnLst>
                          <xsl:call-template name="timingPar_other">
                            <xsl:with-param name="pos" select="position()"/>
                          </xsl:call-template>
                        </p:childTnLst>
                      </p:cTn>
                      <xsl:choose>
                        <xsl:when test="../演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34">
                          <p:nextCondLst>
                            <p:cond evt="onClick" delay="0">
                              <p:tgtEl>
                                <p:spTgt>
                                  <xsl:variable name="mediaId" select="./图:图形_8062/图:其他对象引用_8038"/>
                                  <xsl:attribute name="spid">
                                    <xsl:call-template name="mediaIdConvetor">
                                      <xsl:with-param name="mediaId" select="$mediaId"/>
                                    </xsl:call-template>
                                  </xsl:attribute>
                                </p:spTgt>
                              </p:tgtEl>
                            </p:cond>
                          </p:nextCondLst>
                        </xsl:when>
                        <xsl:otherwise>
                          <p:prevCondLst>
                            <p:cond evt="onPrev" delay="0">
                              <p:tgtEl>
                                <p:sldTgt/>
                              </p:tgtEl>
                            </p:cond>
                          </p:prevCondLst>
                          <p:nextCondLst>
                            <p:cond evt="onNext" delay="0">
                              <p:tgtEl>
                                <p:sldTgt/>
                              </p:tgtEl>
                            </p:cond>
                          </p:nextCondLst>
                        </xsl:otherwise>
                      </xsl:choose>
                    </p:seq>
                    <p:video>
                      <p:cMediaNode vol="80000">
                        <p:cTn id="7" fill="hold" display="0">
                          <p:stCondLst>
                            <p:cond delay="indefinite"/>
                          </p:stCondLst>
                        </p:cTn>
                        <p:tgtEl>
                          <p:spTgt>
                            <xsl:variable name="mediaId" select="./图:图形_8062/图:其他对象引用_8038"/>
                            <xsl:attribute name="spid">
                              <xsl:call-template name="mediaIdConvetor">
                                <xsl:with-param name="mediaId" select="$mediaId"/>
                              </xsl:call-template>
                            </xsl:attribute>
                          </p:spTgt>
                        </p:tgtEl>
                      </p:cMediaNode>
                    </p:video>
                </xsl:for-each>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:tnLst>

        <xsl:if test="演:动画_6B1A/演:序列_6B1B">
          <xsl:if test="./演:增强_6B35/演:动画文本_6B3A/@组合文本_6B3F">
            <p:bldLst>
              <p:bldP uiExpand="1" animBg="1" autoUpdateAnimBg="0" grpId="0">
                <xsl:call-template name="sptgt"/>

                <xsl:variable name="bldTxt" select=".//演:增强_6B35/演:动画文本_6B3A/@组合文本_6B3F"/>
                <xsl:choose>
                  <xsl:when test="$bldTxt='All-paragraphs-at-once'">
                    <xsl:attribute name="build">allAtOnce</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-1st-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">1</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-2nd-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">2</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-3rd-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">3</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-4th-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">4</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-5th-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">5</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='As-one-object'">
                    <xsl:attribute name="build">Whole</xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="build">p</xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <xsl:if test="./演:增强_6B35/演:动画文本_6B3A/@是否使用相反顺序_6B3E">
                  <xsl:attribute name="rev">
                    <xsl:choose>
                      <xsl:when test=".//演:增强_6B35/演:动画文本_6B3A/@是否使用相反顺序_6B3E='true'">1</xsl:when>
                      <xsl:otherwise>0</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>
              </p:bldP>
            </p:bldLst>
          </xsl:if>
        </xsl:if>
      </p:timing>
    </xsl:if>
  </xsl:template>
  <xsl:template name="timingPar_other">
    <xsl:param name="pos"/>
    <p:par>
      <p:cTn id="3" fill="hold">
        <p:stCondLst>
          <p:cond delay="0"/>
          <xsl:if test="$pos='1' and 演:序列_6B1B[1]/演:定时_6B2E/@事件_6B2F!='on-click'">
            <p:cond evt="onBegin" delay="0">
              <p:tn val="2"/>
            </p:cond>
          </xsl:if>
        </p:stCondLst>
        <p:childTnLst>
          <p:par>
            <p:cTn id="4" fill="hold">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="5" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect">
                    <p:stCondLst>
                      <p:cond delay="0"/>
                    </p:stCondLst>
                    <p:childTnLst>
                      <p:cmd type="call" cmd="togglePause">
                        <p:cBhvr>
                          <p:cTn id="6" dur="1" fill="hold"/>
                          <p:tgtEl>
                            <p:spTgt>
                              <xsl:variable name="mediaId" select="./图:图形_8062/图:其他对象引用_8038"/>
                              <xsl:attribute name="spid">
                                <xsl:call-template name="mediaIdConvetor">
                                  <xsl:with-param name="mediaId" select="$mediaId"/>
                                </xsl:call-template>
                              </xsl:attribute>
                            </p:spTgt>
                          </p:tgtEl>
                        </p:cBhvr>
                      </p:cmd>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </xsl:template>
  <!--end liuyin 20121211 修改音视频丢失-->
  
  <xsl:template name="animation">
    <!--修改动画 标签 李娟 2012.01.06-->
    <!--2010.04.22<xsl:if test="演:动画/演:序列">-->
    <xsl:if test="演:动画_6B1A/演:序列_6B1B">
      <p:timing>
        <p:tnLst>
          <p:par>
            <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
              <p:childTnLst>
                <p:seq concurrent="1" nextAc="seek">
                  <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                    <!--2010-11-2罗文甜：增加触发器-->
                    <xsl:if test="演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34/text()!=''">
                      <p:stCondLst>
                        <p:cond evt="onClick" delay="0">
                          <p:tgtEl>
                            <p:spTgt>
                              <xsl:attribute name="spid">
                                <xsl:value-of select="translate(演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34,':','')"/>
                              </xsl:attribute>
                            </p:spTgt>
                          </p:tgtEl>
                        </p:cond>
                      </p:stCondLst>
                      <p:endSync evt="end" delay="0">
                        <p:rtn val="all"/>
                      </p:endSync>
                    </xsl:if>

                    <p:childTnLst>
                      <xsl:for-each select="演:动画_6B1A/演:序列_6B1B">
                        <xsl:call-template name="timingPar">
                          <!--xsl:apply-templates select="." mode="level1"-->
                          <xsl:with-param name="pos" select="position()"/>
                        </xsl:call-template>
                      </xsl:for-each>
                    </p:childTnLst>
                  </p:cTn>
                  <!--2010-11-2罗文甜：增加触发器-->
                  <xsl:choose>
                    <xsl:when test="演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34/text()!=''">
                      <p:nextCondLst>
                        <p:cond evt="onClick" delay="0">
                          <p:tgtEl>
                            <p:spTgt>
                              <xsl:attribute name="spid">
                                <xsl:value-of select="translate(演:动画_6B1A/演:序列_6B1B/演:定时_6B2E/@触发对象引用_6B34,':','')"/>
                              </xsl:attribute>
                            </p:spTgt>
                          </p:tgtEl>
                        </p:cond>
                      </p:nextCondLst>
                    </xsl:when>
                    <xsl:otherwise>
                      <p:prevCondLst>
                        <p:cond evt="onPrev" delay="0">
                          <p:tgtEl>
                            <p:sldTgt/>
                          </p:tgtEl>
                        </p:cond>
                      </p:prevCondLst>
                      <p:nextCondLst>
                        <p:cond evt="onNext" delay="0">
                          <p:tgtEl>
                            <p:sldTgt/>
                          </p:tgtEl>
                        </p:cond>
                      </p:nextCondLst>
                    </xsl:otherwise>
                  </xsl:choose>
                </p:seq>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:tnLst>
        <!--2010.04.22<xsl:if test="演:动画/演:序列">-->
        <!--2010-12-29罗文甜。修改组合文本bug-->
        <xsl:if test="演:动画_6B1A/演:序列_6B1B">
          <!--p:bldLst>          
              <p:bldP uiExpand="1" animBg="1" autoUpdateAnimBg="0" grpId="0">
                <xsl:call-template name="sptgt"/-->
          <!--2010.04.30 myx add >
                <xsl:if test="@演:段落引用">
                  <xsl:attribute name="build">p</xsl:attribute>
                </xsl:if-->

          <!--2010.10.25 罗文甜增加 组合文本 build ,buildlel,rev-->
          <xsl:if test="./演:增强_6B35/演:动画文本_6B3A/@组合文本_6B3F">
            <p:bldLst>
              <p:bldP uiExpand="1" animBg="1" autoUpdateAnimBg="0" grpId="0">
                <xsl:call-template name="sptgt"/>

                <xsl:variable name="bldTxt" select=".//演:增强_6B35/演:动画文本_6B3A/@组合文本_6B3F"/>
                <xsl:choose>
                  <xsl:when test="$bldTxt='All-paragraphs-at-once'">
                    <xsl:attribute name="build">allAtOnce</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-1st-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">1</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-2nd-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">2</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-3rd-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">3</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-4th-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">4</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='By-5th-paragraph'">
                    <xsl:attribute name="build">p</xsl:attribute>
                    <xsl:attribute name="bldLvl">5</xsl:attribute>
                  </xsl:when>
                  <xsl:when test="$bldTxt='As-one-object'">
                    <xsl:attribute name="build">Whole</xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="build">p</xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
                <xsl:if test="./演:增强_6B35/演:动画文本_6B3A/@是否使用相反顺序_6B3E">
                  <xsl:attribute name="rev">
                    <xsl:choose>
                      <xsl:when test=".//演:增强_6B35/演:动画文本_6B3A/@是否使用相反顺序_6B3E='true'">1</xsl:when>
                      <xsl:otherwise>0</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:if>
              </p:bldP>
            </p:bldLst>
          </xsl:if>
        </xsl:if>
      </p:timing>
    </xsl:if>
  </xsl:template>
  <!--<xsl:for-each select="演:动画/演:序列[演:序列/演:效果/*]">
                              <xsl:apply-templates select="." mode="outter"/>
                            </xsl:for-each>-->
  <xsl:template name="timingPar">
    <xsl:param name="pos"/>
    <p:par>
      <p:cTn fill="hold">
        <p:stCondLst>
          <p:cond delay="indefinite"/>
          <xsl:if test="$pos='1' and 演:序列_6B1B[1]/演:定时_6B2E/@事件_6B2F!='on-click'">
            <p:cond evt="onBegin" delay="0">
              <p:tn val="2"/>
            </p:cond>
          </xsl:if>
        </p:stCondLst>
        <xsl:if test="演:序列_6B1B/演:效果_6B40/*">
          <p:childTnLst>
            <xsl:for-each select="演:序列_6B1B">
              <p:par>
                <p:cTn fill="hold">
                  <p:stCondLst>
                    <p:cond>
                      <xsl:if test="@delay">
                        <xsl:attribute name="delay">
                          <xsl:value-of select="@delay"/>
                        </xsl:attribute>
                      </xsl:if>
                    </p:cond>
                  </p:stCondLst>
                  <p:childTnLst>
                    <xsl:apply-templates select=".">
                      <xsl:with-param name="id" select="@id"/>
                    </xsl:apply-templates>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </xsl:for-each>
          </p:childTnLst>
        </xsl:if>
      </p:cTn>
    </p:par>
  </xsl:template>

  <xsl:template match="演:序列_6B1B">
    <xsl:param name="id"/>
    <p:par>
      <p:cTn fill="hold">
        <xsl:attribute name="id">
          <xsl:value-of select="$id"/>
        </xsl:attribute>
        <xsl:attribute name="nodeType">
          <xsl:choose>
            <xsl:when test="演:定时_6B2E/@事件_6B2F='on-click'">clickEffect</xsl:when>
            <xsl:when test="演:定时_6B2E/@事件_6B2F='with-previous'">withEffect</xsl:when>
            <xsl:when test="演:定时_6B2E/@事件_6B2F='after-previous'">afterEffect</xsl:when>
            <xsl:otherwise>clickEffect</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="演:定时_6B2E/@重复_6B32!='none'">
          <xsl:attribute name="repeatCount">
            <xsl:choose>
              <xsl:when test="演:定时_6B2E/@重复_6B32='until-next-click' or 演:定时_6B2E/@重复_6B32='until-end-of-slide'">indefinite</xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="演:定时_6B2E/@重复_6B32*1000"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="演:定时_6B2E/@是否回卷_6B33='true'">
          <xsl:attribute name="fill">remove</xsl:attribute>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41">
            <xsl:call-template name="present_id_type_entr_exit">
              <xsl:with-param name="animNode" select="演:效果_6B40"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="演:效果_6B40/演:强调_6B07">
            <xsl:call-template name="present_id_type_emph"/>
          </xsl:when>
          <xsl:when test="演:效果_6B40/演:退出_6BBE">
            <xsl:call-template name="present_id_type_entr_exit">
              <!--2013-03-19  唐江 去掉select语句中的 “ 演:效果_6B40/ ” 删除 “ 演:退出_6BBE/演:基本型_6C23 ” -->
              <xsl:with-param name="animNode" select="演:效果_6B40"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="演:效果_6B40/演:动作路径_6BD1">
            <xsl:attribute name="presetID">13</xsl:attribute>
            <!--<xsl:attribute name="p:animMotion">0</xsl:attribute>-->
            <!--永中Bug-->
          </xsl:when>
        </xsl:choose>
        <xsl:attribute name="presetClass">
          <xsl:choose>
            <xsl:when test="演:效果_6B40/演:进入_6B41">entr</xsl:when>
            <xsl:when test="演:效果_6B40/演:强调_6B07">emph</xsl:when>
            <xsl:when test="演:效果_6B40/演:退出_6BBE">exit</xsl:when>
            <xsl:when test="演:效果_6B40/演:动作路径_6BD1">path</xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <!--2011-2-18罗文甜：增加属性grpId-->
        <xsl:attribute name="grpId">0</xsl:attribute>
        <!--2010-11-1罗文甜：增加decel accel,aotoRev-->
        <xsl:if test="演:效果_6B40/演:动作路径_6BD1/@是否平稳开始_6BD8='true'">
          <xsl:attribute name="accel">
            <xsl:value-of select="'50000'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="演:效果_6B40/演:动作路径_6BD1/@是否平稳结束_6BD9='true'">
          <xsl:attribute name="decel">
            <xsl:value-of select="'50000'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="演:效果_6B40/演:动作路径_6BD1/@是否自动翻转_6BDA='true'">
          <xsl:attribute name="autoRev">1</xsl:attribute>
        </xsl:if>

        <p:stCondLst>
          <p:cond>
            <xsl:if test="演:定时_6B2E/@延时_6B30">
              <xsl:attribute name="delay">
                <!--luowentian 考虑金山软件和标准-->
                <xsl:choose>
                  <!--start liuyin 20130107 修改动画序列的计时中的延迟转换错误。-->
                  <xsl:when test="contains(演:定时_6B2E/@延时_6B30,'P0Y0M0DT0H0M')">
                    <!--<xsl:value-of select="substring-before(substring-after(演:定时_6B2E/@延时_6B30,'H0M'),'S')*1000"/>-->
					           <xsl:value-of select="substring-before(substring-after(演:定时_6B2E/@延时_6B30,'P0Y0M0DT0H0M'),'S')*1000"/>
                  <!--end liuyin 20130107 修改动画序列的计时中的延迟转换错误。-->
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="演:定时_6B2E/@延时_6B30*1000"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </p:cond>
        </p:stCondLst>
        
        <!--start liuyin 20130107 修改动画序列的计时中的重复转换错误。-->
        <!--<xsl:if test="演:效果_6B40/演:强调_6B07//@期间_6B98='until-next-click'">-->
        <xsl:if test="演:定时_6B2E/@重复_6B32='until-next-click'">
          <!--start liuyin 20130107 修改动画序列的计时中的重复转换错误。-->
          <p:endCondLst>
            <p:cond evt="onNext" delay="0">
              <p:tgtEl>
                <p:sldTgt/>
              </p:tgtEl>
            </p:cond>
          </p:endCondLst>
        </xsl:if>
        <!--2011-4-12 luowentian add duration by WPS-->
        <xsl:if test="演:增强_6B35/演:动画文本_6B3A/@发送_6B3B!='all-at-once'">
          <xsl:choose>
            <xsl:when test="演:增强_6B35/演:动画文本_6B3A/@发送_6B3B='by-word'">
              <p:iterate type="wd">
                <p:tmPct>
                  <xsl:attribute name="val">
                    <xsl:choose>
                      <xsl:when test="演:增强_6B35/演:动画文本_6B3A/@间隔_6B3C">
                        <xsl:value-of select="substring-before(substring-after(演:增强_6B35/演:动画文本_6B3A/@间隔_6B3C,'PT'),'S')*100000 div 2"/>
                      </xsl:when>
                      <xsl:otherwise>10000</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </p:tmPct>
              </p:iterate>
            </xsl:when>
            <xsl:otherwise>
              <p:iterate type="lt">
                <p:tmPct>
                  <xsl:attribute name="val">
                    <xsl:choose>
                      <xsl:when test="演:增强_6B35/演:动画文本_6B3A/@间隔_6B3C">
                        <xsl:value-of select="substring-before(substring-after(演:增强_6B35/演:动画文本_6B3A/@间隔_6B3C,'PT'),'S')*100000 div 2"/>
                      </xsl:when>
                      <xsl:otherwise>10000</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </p:tmPct>
              </p:iterate>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
     
          <xsl:choose>
            <xsl:when test="演:效果_6B40/演:进入_6B41">
              <xsl:call-template name="animMain_entr">
                <xsl:with-param name="id" select="$id"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="演:效果_6B40/演:强调_6B07">
              <xsl:call-template name="animMain_emph">
                <xsl:with-param name="id" select="$id"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="演:效果_6B40/演:退出_6BBE">
              <!--2011-5-16 luowentian-->
              <xsl:call-template name="animMain_exit">
                <xsl:with-param name="id" select="$id"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="演:效果_6B40/演:动作路径_6BD1">
              <xsl:call-template name="animMain_path">
                <xsl:with-param name="id" select="$id"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
      
        
      </p:cTn>
    </p:par>
    <xsl:for-each select="./演:序列_6B1B">
      <xsl:apply-templates select=".">
        <xsl:with-param name="id" select="@id"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>

	<!--动画进入效果 李娟 2012.01.06-->
  <xsl:template name="animMain_entr">
    <xsl:param name="id"/>
    <p:childTnLst>
    <xsl:choose>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:闪烁一次_6B51">
          <p:set>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="visible"/>
            </p:to>
          </p:set>
      </xsl:when>
      <!--唐江 2013-03-22  增加细微型中的渐变式回旋、弹跳效果 修改bug_2710  start-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:出现_6B46">
        <p:set>
          <p:cBhvr>
            <p:cTn dur="1" fill="hold">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.visibility</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="visible"/>
          </p:to>
        </p:set>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变式回旋_6B69">
          <p:set>
          <p:cBhvr>
            <p:cTn dur="1" fill="hold">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.visibility</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="visible"/>
          </p:to>
        </p:set>
        <p:animEffect transition="in" filter="fade">
          <p:cBhvr>
            <p:cTn dur="2000"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="2000" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_w</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0" fmla="#ppt_w*sin(2.5*pi*$)">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="1"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="2000" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_h</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal val="#ppt_h"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_h"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:华丽型_6B7C/演:弹跳_6B7D">
        <p:set>
          <p:cBhvr>
            <p:cTn dur="1" fill="hold">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.visibility</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="visible"/>
          </p:to>
        </p:set>
        <p:animEffect transition="in" filter="wipe(down)">
          <p:cBhvr>
            <p:cTn dur="580">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn  dur="1822" tmFilter="0,0; 0.14,0.36; 0.43,0.73; 0.71,0.91; 1.0,1.0">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_x</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal val="#ppt_x-0.25"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_x"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="664" tmFilter="0.0,0.0; 0.25,0.07; 0.50,0.2; 0.75,0.467; 1.0,1.0">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0" fmla="#ppt_y-sin(pi*$)/3">
              <p:val>
                <p:fltVal val="0.5"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="1"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="664" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
              <p:stCondLst>
                <p:cond delay="664"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0" fmla="#ppt_y-sin(pi*$)/9">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="1"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="332" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
              <p:stCondLst>
                <p:cond delay="1324"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0" fmla="#ppt_y-sin(pi*$)/27">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="1"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn  dur="164" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
              <p:stCondLst>
                <p:cond delay="1656"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0" fmla="#ppt_y-sin(pi*$)/81">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="1"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="26">
              <p:stCondLst>
                <p:cond delay="650"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="60000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn  dur="166" decel="50000">
              <p:stCondLst>
                <p:cond delay="676"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="100000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="26">
              <p:stCondLst>
                <p:cond delay="1312"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="80000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="166" decel="50000">
              <p:stCondLst>
                <p:cond delay="1338"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="100000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn  dur="26">
              <p:stCondLst>
                <p:cond delay="1642"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="90000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="166" decel="50000">
              <p:stCondLst>
                <p:cond delay="1668"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="100000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="26">
              <p:stCondLst>
                <p:cond delay="1808"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="95000"/>
        </p:animScale>
        <p:animScale>
          <p:cBhvr>
            <p:cTn dur="166" decel="50000">
              <p:stCondLst>
                <p:cond delay="1834"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
          <p:to x="100000" y="100000"/>
        </p:animScale>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:温和型_6B6D/演:上升_6B76">
        <p:set>
          <p:cBhvr>
            <p:cTn dur="1" fill="hold">
              <p:stCondLst>
                <p:cond delay="0"/>
              </p:stCondLst>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.visibility</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="visible"/>
          </p:to>
        </p:set>
        <p:animEffect transition="in" filter="fade">
          <p:cBhvr>
            <p:cTn  dur="1000"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="1000" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_x</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal val="#ppt_x"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_x"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn dur="1000" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal val="#ppt_y+.1"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_y"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
      </xsl:when>
      <!--唐江 2013-03-22  增加细微型中的渐变式回旋、弹跳效果 修改bug_2710 end-->
      <xsl:otherwise>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="visible"/>
            </p:to>
          </p:set>
          <xsl:call-template name="animEffect_entr"/>
      </xsl:otherwise>
    </xsl:choose>
    </p:childTnLst>
	  <!--先注销 aferAmin 调用 李娟  2012.02.16-->
	  
    <xsl:call-template name="afterAnim">
      <xsl:with-param name="id" select="$id"/>
    </xsl:call-template>
	 
  </xsl:template>

  <xsl:template name="animEffect_entr">
	  
    <xsl:choose>
      <!--2011-4-07罗文甜，修改动画效果：切入-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60">
        <xsl:variable name="dir" select="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58"/>
        <xsl:variable name="ppt">
          <xsl:choose>
            <xsl:when test="contains($dir,'left') or contains($dir,'right')">ppt_x</xsl:when>
            <xsl:when test="contains($dir,'top') or contains($dir,'bottom')">ppt_y</xsl:when>
            <xsl:otherwise>ppt_x</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="pptval">
          <xsl:choose>
            <xsl:when test="contains($dir,'right')">#ppt_x+#ppt_w*1.125000</xsl:when>
            <xsl:when test="contains($dir,'left')">#ppt_x-#ppt_w*1.125000</xsl:when>
            <xsl:when test="contains($dir,'bottom')">#ppt_y+#ppt_h*1.125000</xsl:when>
            <xsl:when test="contains($dir,'top')">#ppt_y-#ppt_h*1.125000</xsl:when>           
          </xsl:choose>
        </xsl:variable>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr additive="base">
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>
                <xsl:value-of select="$ppt"/>
              </p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$pptval"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="concat('#',$ppt)"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:animEffect transition="in">
          <xsl:attribute name="filter">
            <xsl:call-template name="animEffect_filter_entr"/>
          </xsl:attribute>
          <p:cBhvr>
            <p:cTn>
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
      </xsl:when>
      
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43 or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57   or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47 or 
                演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49   or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:菱形_6B5D       or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B or 
                演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:圆形扩展_6B56 or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61   or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:向内溶解_6B64 or 
                演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:随机线条_6B62 or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:十字形扩展_6B53 or 
                演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:棋盘_6B4E     or 演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E ">
        <p:animEffect transition="in">
          <xsl:attribute name="filter">
            <xsl:call-template name="animEffect_filter_entr"/>
          </xsl:attribute>
          <p:cBhvr>
            <p:cTn>
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:飞入_6B59">
        <xsl:variable name="dir" select="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:飞入_6B59/@方向_6B5A"/>
        <xsl:variable name="x">
          <xsl:choose>
            <xsl:when test="contains($dir,'left')">0-#ppt_w/2</xsl:when>
            <xsl:when test="contains($dir,'right')">1+#ppt_w/2</xsl:when>
            <xsl:otherwise>#ppt_x</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="y">
          <xsl:choose>
            <xsl:when test="contains($dir,'top')">0-#ppt_h/2</xsl:when>
            <xsl:when test="contains($dir,'bottom')">1+#ppt_h/2</xsl:when>
            <xsl:otherwise>#ppt_y</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr additive="base">
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_x</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$x"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_x"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr additive="base">
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$y"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_y"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B">
        <xsl:variable name="dir" select="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58"/>
        <xsl:variable name="x">
          <xsl:choose>
            <xsl:when test="contains($dir,'left')">0-#ppt_w/2</xsl:when>
            <xsl:when test="contains($dir,'right')">1+#ppt_w/2</xsl:when>
            <xsl:otherwise>#ppt_x</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="y">
          <xsl:choose>
            <xsl:when test="contains($dir,'top')">0-#ppt_h/2</xsl:when>
            <xsl:when test="contains($dir,'bottom')">1+#ppt_h/2</xsl:when>
            <xsl:otherwise>#ppt_x</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr additive="base">
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_x</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$x"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_x"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr additive="base">
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_y</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:strVal>
                  <xsl:attribute name="val">
                    <xsl:value-of select="$y"/>
                  </xsl:attribute>
                </p:strVal>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_y"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
      </xsl:when>
      <!--luowentian:转换为 翻转由远及近-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:随机效果_6B55">
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_w</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_w"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>ppt_h</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:strVal val="#ppt_h"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:anim calcmode="lin" valueType="num">
          <p:cBhvr>
            <p:cTn fill="hold">
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.rotation</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:tavLst>
            <p:tav tm="0">
              <p:val>
                <p:fltVal val="90"/>
              </p:val>
            </p:tav>
            <p:tav tm="100000">
              <p:val>
                <p:fltVal val="0"/>
              </p:val>
            </p:tav>
          </p:tavLst>
        </p:anim>
        <p:animEffect transition="in" filter="fade">
          <p:cBhvr>
            <p:cTn>
              <xsl:call-template name="dur"/>
            </p:cTn>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
          </p:cBhvr>
        </p:animEffect>
      </xsl:when>
	  <!--添加细微型 李娟 2012 05 07-->
	<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变式缩放_6B6A">
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_w</p:attrName>
					</p:attrNameLst>
					
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:fltVal val="0"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_w"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_h</p:attrName>
					</p:attrNameLst>
					
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:fltVal val="0"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_h"/>
						</p:val>
					</p:tav>
				</p:tavLst>
				</p:anim>
		</xsl:when>
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变_6B67">
			<p:animEffect transition="in" filter="fade">
				<p:cBhvr>
					<p:cTn>
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
				</p:cBhvr>
			</p:animEffect>
		</xsl:when>
		
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:展开_6B6B">
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_w</p:attrName>
					</p:attrNameLst>
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:strVal val="#ppt_w*0.70"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_w"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_h</p:attrName>
					</p:attrNameLst>
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:strVal val="#ppt_h"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_h"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:animEffect transition="in" filter="fade">
				<p:cBhvr>
					<p:cTn>
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
				</p:cBhvr>
			</p:animEffect>
		</xsl:when>
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:温和型_6B6D/演:翻转时由远及近_6B6E">
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_w</p:attrName>
					</p:attrNameLst>
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:fltVal val="0"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_w"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>ppt_h</p:attrName>
					</p:attrNameLst>
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:fltVal val="0"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:strVal val="#ppt_h"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:anim calcmode="lin" valueType="num">
				<p:cBhvr>
					<p:cTn fill="hold">
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
					<p:attrNameLst>
						<p:attrName>style.rotation</p:attrName>
					</p:attrNameLst>
				</p:cBhvr>
				<p:tavLst>
					<p:tav tm="0">
						<p:val>
							<p:fltVal val="90"/>
						</p:val>
					</p:tav>
					<p:tav tm="100000">
						<p:val>
							<p:fltVal val="0"/>
						</p:val>
					</p:tav>
				</p:tavLst>
			</p:anim>
			<p:animEffect transition="in" filter="fade">
				<p:cBhvr>
					<p:cTn>
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
						</p:tgtEl>
				</p:cBhvr>
			</p:animEffect>
		</xsl:when>
		<xsl:otherwise>
			<p:animEffect transition="in">
				<xsl:attribute name="filter">
					<xsl:call-template name="animEffect_filter_entr"/>
				</xsl:attribute>
				<p:cBhvr>
					<p:cTn>
						<xsl:call-template name="dur"/>
					</p:cTn>
					<p:tgtEl>
						<p:spTgt>
							<xsl:call-template name="sptgt"/>
							<xsl:call-template name="parpref"/>
						</p:spTgt>
					</p:tgtEl>
				</p:cBhvr>
			</p:animEffect>
		</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="animEffect_filter_entr">
    <xsl:choose>
		<!--添加细微型  李娟 2012 050 7-->
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变式缩放_6B6A">fade</xsl:when>
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变_6B67">fade</xsl:when>
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:渐变式回旋_6B69">fade</xsl:when>
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:细微型_6B66/演:展开_6B6B">fade</xsl:when>
		<!--添加温和型 李娟 2012 0507-->
		<xsl:when test="演:效果_6B40/演:进入_6B41/演:温和型_6B6D/演:翻转时由远及近_6B6E">fade</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@方向_6B45='horizontal'">blinds(horizontal)</xsl:when>
          <xsl:otherwise>blinds(vertical)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!--2011=4-7罗文甜，修改擦除效果-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-bottom'">wipe(down)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-top'">wipe(up)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-left'">wipe(left)</xsl:when>
          <!--<xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from bottom'">wipe(down)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from top'">wipe(up)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from left'">wipe(left)</xsl:when>-->
          <xsl:otherwise>wipe(right)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@方向_6B48='in'">box(in)</xsl:when>
          <xsl:otherwise>box(out)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left-down'">strips(downLeft)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left-up'">strips(upLeft)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='right-down'">strips(downRight)</xsl:when>
          <!--<xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left down'">strips(downLeft)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left up'">strips(upLeft)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='right down'">strips(downRight)</xsl:when>-->
          <xsl:otherwise>strips(upRight)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:菱形_6B5D">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:菱形_6B5D/@方向_6B48='out'">diamond(out)</xsl:when>
          <xsl:otherwise>diamond(in)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!--2010-10-25罗文甜：修改辐射状属性为轮辐-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B/@轮辐_6B4D='1'">wheel(1)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B/@轮辐_6B4D='2'">wheel(2)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B/@轮辐_6B4D='3'">wheel(3)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B/@轮辐_6B4D='4'">wheel(4)</xsl:when>
          <xsl:otherwise>wheel(8)</xsl:otherwise>
        </xsl:choose>
        <!--xsl:choose>
          <xsl:when test="演:效果/演:进入/演:轮子/@演:轮辐='1'">wheel(1)</xsl:when>
          <xsl:when test="演:效果/演:进入/演:轮子/@演:轮辐='2'">wheel(2)</xsl:when>
          <xsl:when test="演:效果/演:进入/演:轮子/@演:轮辐='3'">wheel(3)</xsl:when>
          <xsl:when test="演:效果/演:进入/演:轮子/@演:轮辐='4'">wheel(4)</xsl:when>
          <xsl:otherwise>wheel(8)</xsl:otherwise>
        </xsl:choose-->
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:圆形扩展_6B56">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:圆形扩展_6B56/@方向_6B48='in'">circle(in)</xsl:when>
          <xsl:otherwise>circle(out)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61">wedge</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:向内溶解_6B64">dissolve</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:随机线条_6B62">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:随机线条_6B62/@方向_6B45='horizontal'">randombar(horizontal)</xsl:when>
          <xsl:otherwise>randombar(vertical)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:十字形扩展_6B53">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:十字形扩展_6B53/@方向_6B48='in'">plus(in)</xsl:when>
          <xsl:otherwise>plus(out)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <!--2011=4-7罗文甜，修改切入效果-->
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from-bottom'">wipe(up)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from-top'">wipe(down)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from-left'">wipe(right)</xsl:when>
          <!--<xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from bottom'">wipe(up)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from top'">wipe(down)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58='from left'">wipe(right)</xsl:when>-->
          <xsl:otherwise>wipe(left)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:棋盘_6B4E">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:棋盘_6B4E/@方向_6B50='down'">checkerboard(down)</xsl:when>
          <xsl:otherwise>checkerboard(across)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E">
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical-out'">barn(outVertical)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical-in'">barn(inVertical)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal-out'">barn(outHorizontal)</xsl:when>
          <!--<xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical out'">barn(outVertical)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical in'">barn(inVertical)</xsl:when>
          <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal out'">barn(outHorizontal)</xsl:when>-->
          <xsl:otherwise>barn(inHorizontal)</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
		<xsl:otherwise>fade</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="present_id_type_entr_exit">
    <xsl:param name="animNode"/>
    <xsl:choose>
		<!--退出 效果-->
      <xsl:when test="$animNode/演:退出_6BBE/演:华丽型_6C26/演:弹跳_6B7D">
        <xsl:attribute name="presetID">26</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:细微型_6C24/演:渐变式回旋_6B69">
        <xsl:attribute name="presetID">45</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:细微型_6C24/演:渐变_6B67">
        <xsl:attribute name="presetID">10</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:温和型_6C25/演:下降_6B78">
        <xsl:attribute name="presetID">42</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-17  增加温和型中的上升效果 修改bug_2710  start-->
      <xsl:when test="$animNode/演:退出_6BBE/演:基本型_6C23/演:消失_6BC7">
        <xsl:attribute name="presetID">1</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:基本型_6C23/演:飞出_6BBF">
        <xsl:attribute name="presetID">2</xsl:attribute>
        <xsl:attribute name="presetSubtype">4</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:基本型_6C23/演:劈裂_6B5E">
        <xsl:attribute name="presetID">16</xsl:attribute>
        <xsl:attribute name="presetSubtype">21</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B">
        <xsl:attribute name="presetID">21</xsl:attribute>
        <xsl:attribute name="presetSubtype">1</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:细微型_6C24/演:渐变_6B67">
        <xsl:attribute name="presetID">10</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:退出_6BBE/演:细微型_6C24/演:渐变式缩放_6B6A">
        <xsl:attribute name="presetID">53</xsl:attribute>
        <xsl:attribute name="presetSubtype">32</xsl:attribute>
      </xsl:when>
      
      <!--进入 效果-->
      <xsl:when test="$animNode/演:进入_6B41/演:温和型_6B6D/演:上升_6B76">
        <xsl:attribute name="presetID">42</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:温和型_6B6D/演:翻转时由远及近_6B6E">
        <xsl:attribute name="presetID">31</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:华丽型_6B7C/演:弹跳_6B7D">
        <xsl:attribute name="presetID">26</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-17  增加温和型中的上升效果 修改bug_2710  end-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43">
        <xsl:attribute name="presetID">3</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@方向_6B45='horizontal'">10</xsl:when>
            <xsl:otherwise>5</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--2011-5-9罗文甜 修改方向bug。-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57">
        <xsl:attribute name="presetID">22</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <!--<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from right'">2</xsl:when>-->
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:擦除_6B57/@方向_6B58='from-right'">2</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--<xsl:when test="$animNode/演:出现 or $animNode/演:消失">--><!--注销消失效果 李娟 2012.01.06-->
		<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:出现_6B46 or $animNode/演:消失_6BC7">
        <xsl:attribute name="presetID">1</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:飞入_6B59 or $animNode/演:进入_6B41/演:基本型_6B42/演:飞出_6BBF"><!--注销飞出效果 李娟 2012.01.06-->
      <!--<xsl:when test="$animNode/演:基本型_6B42/演:飞入_6B59 or $animNode/演:飞出">-->
        <xsl:attribute name="presetID">2</xsl:attribute>
        <xsl:attribute name="presetSubtype">
			<xsl:variable name="dir" select="$animNode/演:进入_6B41/演:基本型_6B42/演:飞入_6B59/@方向_6B5A"/>
			<!--<xsl:variable name="dir" select="$animNode//@演:方向"/>-->
          <xsl:choose>
            <!--start liuyin 20130327 修改bug2810 动画效果不正确。注此处由于代码原来是“from bottom”更改为“from-bottom”，存在多个bug-->
            <!--<xsl:when test="$dir='to bottom' or $dir='from bottom'"            >4</xsl:when>-->
            <!--<xsl:when test="$dir='to left' or $dir='from left'"                >8</xsl:when>
            <xsl:when test="$dir='to right' or $dir='from bottom'"             >2</xsl:when>
            <xsl:when test="$dir='to top' or $dir='from bottom'"               >1</xsl:when>
            <xsl:when test="$dir='to bottom-left' or $dir='from bottom-left'"  >12</xsl:when>
            <xsl:when test="$dir='to bottom-right' or $dir='from bottom-right'">6</xsl:when>
            <xsl:when test="$dir='to top-left' or $dir='from top-left'"        >9</xsl:when>-->

            <xsl:when test="$dir='to-bottom' or $dir='from-bottom'"            >4</xsl:when>
            <xsl:when test="$dir='to-left' or $dir='from-left'"                >8</xsl:when>
            <xsl:when test="$dir='to-right' or $dir='from-bottom'"             >2</xsl:when>
            <xsl:when test="$dir='to-top' or $dir='from-bottom'"               >1</xsl:when>
            <xsl:when test="$dir='to-bottom-left' or $dir='from-bottom-left'"  >12</xsl:when>
            <xsl:when test="$dir='to-bottom-right' or $dir='from-bottom-right'">6</xsl:when>
            <xsl:when test="$dir='to-top-left' or $dir='from-top-left'"        >9</xsl:when>
            <!--start liuyin 20130327 修改bug2810 动画效果不正确。-->
            
            <xsl:otherwise>3</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:盒状_6B47">
        <xsl:attribute name="presetID">4</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@方向_6B48='in'">16</xsl:when>
            <xsl:otherwise>32</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B">
        <xsl:attribute name="presetID">7</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <!--<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from right'">2</xsl:when>-->

            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from-bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from-left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢进入_6B5B/@方向_6B58='from-right'">2</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--2011-5-9罗文甜，修改缓慢移出-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1">
        <xsl:attribute name="presetID">7</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <!--<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to right'">2</xsl:when>-->
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to-bottom'">4</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to-left'">8</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:缓慢移出_6BC1/@方向_6BC2='to-right'">2</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49">
        <xsl:attribute name="presetID">18</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <!--<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left down'">12</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left up'">9</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='right down'">6</xsl:when>-->
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left-down'">12</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='left-up'">9</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:阶梯状_6B49/@方向_6B4A='right-down'">6</xsl:when>
            <xsl:otherwise>3</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:盒状_6B47">
        <xsl:attribute name="presetID">8</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@方向_6B48='in'">16</xsl:when>
            <xsl:otherwise>32</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--2010-10-25罗文甜：修改辐射状属性为轮辐-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B">
        <xsl:attribute name="presetID">21</xsl:attribute>
        <!--xsl:attribute name="presetSubtype">
          <xsl:value-of select="$animNode/演:轮子/@演:轮辐"/>
        </xsl:attribute-->
        <xsl:attribute name="presetSubtype">
          <xsl:value-of select="$animNode/演:进入_6B41/演:基本型_6B42/演:轮子_6B4B/@轮辐_6B4D"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E">
        <xsl:attribute name="presetID">16</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <!--2011-4-13 luowentian-->
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal-in'">26</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical-in'">42</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal-out'">21</xsl:when>
            <!--<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal in'">26</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='vertical in'">42</xsl:when>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:劈裂_6B5E/@方向_6B5F='horizontal out'">21</xsl:when>-->
            <xsl:otherwise>37</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:棋盘_6B4E">
        <xsl:attribute name="presetID">5</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:棋盘_6B4E/@方向_6B50='across'">10</xsl:when>
            <xsl:otherwise>5</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:切入_6B60">
        <xsl:attribute name="presetID">12</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:variable name="dir" select="$animNode/演:进入_6B41/演:基本型_6B42/演:切入_6B60/@方向_6B58"/>
          <xsl:choose>
            <xsl:when test="$dir='from bottom'">4</xsl:when>
            <xsl:when test="$dir='from left'">8</xsl:when>
            <xsl:when test="$dir='from right'">2</xsl:when>
            <!--luowentian-->
            <xsl:when test="$dir='to bottom'">4</xsl:when>
            <xsl:when test="$dir='to left'">8</xsl:when>
            <xsl:when test="$dir='to right'">2</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:切出_6BC4">
			<xsl:attribute name="presetID">12</xsl:attribute>
			<xsl:attribute name="presetSubtype">
				<xsl:variable name="dir" select="$animNode/演:进入_6B41/演:基本型_6B42/演:切出_6BC4/@方向_6BC2"/>
				<xsl:choose>
					<xsl:when test="$dir='from bottom'">4</xsl:when>
					<xsl:when test="$dir='from left'">8</xsl:when>
					<xsl:when test="$dir='from right'">2</xsl:when>
					<!--luowentian-->
					<xsl:when test="$dir='to bottom'">4</xsl:when>
					<xsl:when test="$dir='to left'">8</xsl:when>
					<xsl:when test="$dir='to right'">2</xsl:when>
					<xsl:otherwise>1</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</xsl:when>
      <!--luowentian：转换为出现-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:闪烁一次_6B51">
        <xsl:attribute name="presetID">1</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:十字形扩展_6B53">
        <xsl:attribute name="presetID">13</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:十字形扩展_6B53/@方向_6B48='in'">16</xsl:when>
            <xsl:otherwise>32</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--2011-5-9罗文甜，修改方向-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:随机线条_6B62">
        <xsl:attribute name="presetID">14</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:随机线条_6B62/@方向_6B45='horizontal'">10</xsl:when>
            <xsl:otherwise>5</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--luowentian :转换为翻转由原及进-->
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:随机效果_6B55">
        <xsl:attribute name="presetID">31</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:向内溶解_6B64">
        <xsl:attribute name="presetID">9</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:圆形扩展_6B56">
        <xsl:attribute name="presetID">6</xsl:attribute>
        <xsl:attribute name="presetSubtype">
          <xsl:choose>
            <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:圆形扩展_6B56/@方向_6B48='in'">16</xsl:when>
            <xsl:otherwise>32</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="$animNode/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61">
        <xsl:attribute name="presetID">20</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!-- 细微型 温和型 华丽型 永中Bug 暂转为百叶窗 新添加 李娟 2012 05 07-->
		
		<xsl:when test="$animNode/演:进入_6B41/演:细微型_6B66/演:渐变式缩放_6B6A">
			<xsl:attribute name="presetID">53</xsl:attribute>
			<xsl:attribute name="presetSubtype">16</xsl:attribute>
		</xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:细微型_6B66/演:渐变_6B67">
			<xsl:attribute name="presetID">10</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:细微型_6B66/演:渐变式回旋_6B69">
			<xsl:attribute name="presetID">45</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:细微型_6B66/演:展开_6B6B">
			<xsl:attribute name="presetID">55</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		<xsl:when test="$animNode/演:进入_6B41/演:温和型_6B6D/演:翻转时由远及近_6B6E">
			<xsl:attribute name="presetID">31</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		
      <xsl:otherwise>
        <xsl:attribute name="presetID">3</xsl:attribute>
        <xsl:attribute name="presetSubtype">10</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="animMain_emph">
    <xsl:param name="id"/>
    <p:childTnLst>
      <xsl:choose>
		  <!--细微型-->
		  <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:添加下划线_6BB4">
			  <p:set>
				  <p:cBhvr override="childStyle">
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:添加下划线_6BB4/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'500'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>style.textDecorationUnderline</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:to>
					  <p:strVal val="true"/>
				  </p:to>
			  </p:set>
		  </xsl:when>
        <!--唐江 2013-03-17  增加细微型中的忽明忽暗效果 修改bug_2710  start-->
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:忽明忽暗_6BB2 or 演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:加深_6BAC">
          <p:animClr clrSpc="hsl" dir="cw">
            <p:cBhvr override="childStyle">
              <p:cTn dur="500" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.color</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:by>
              <p:hsl h="0" s="-12549" l="-25098"/>
            </p:by>
          </p:animClr>
          <p:animClr clrSpc="hsl" dir="cw">
            <p:cBhvr>
              <p:cTn dur="500" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>fillcolor</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:by>
              <p:hsl h="0" s="-12549" l="-25098"/>
            </p:by>
          </p:animClr>
          <p:animClr clrSpc="hsl" dir="cw">
            <p:cBhvr>
              <p:cTn dur="500" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>stroke.color</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:by>
              <p:hsl h="0" s="-12549" l="-25098"/>
            </p:by>
          </p:animClr>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="500" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>fill.type</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="solid"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:闪动_6BB6">
          <p:animClr clrSpc="rgb" dir="cw">
          <p:cBhvr override="childStyle">
            <p:cTn dur="250" autoRev="1" fill="remove"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.color</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <a:schemeClr val="bg1"/>
          </p:to>
        </p:animClr>
        <p:animClr clrSpc="rgb" dir="cw">
          <p:cBhvr>
            <p:cTn dur="250" autoRev="1" fill="remove"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>fillcolor</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <a:schemeClr val="bg1"/>
          </p:to>
        </p:animClr>
        <p:set>
          <p:cBhvr>
            <p:cTn dur="250" autoRev="1" fill="remove"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>fill.type</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="solid"/>
          </p:to>
        </p:set>
        <p:set>
          <p:cBhvr>
            <p:cTn dur="250" autoRev="1" fill="remove"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>fill.on</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="true"/>
          </p:to>
        </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:补色_6BAE">
          <p:animClr clrSpc="hsl" dir="cw">
          <p:cBhvr override="childStyle">
            <p:cTn dur="500" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>style.color</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:by>
            <p:hsl h="7200000" s="0" l="0"/>
          </p:by>
        </p:animClr>
        <p:animClr clrSpc="hsl" dir="cw">
          <p:cBhvr>
            <p:cTn id="33" dur="500" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>fillcolor</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:by>
            <p:hsl h="7200000" s="0" l="0"/>
          </p:by>
        </p:animClr>
        <p:animClr clrSpc="hsl" dir="cw">
          <p:cBhvr>
            <p:cTn id="34" dur="500" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>stroke.color</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:by>
            <p:hsl h="7200000" s="0" l="0"/>
          </p:by>
        </p:animClr>
        <p:set>
          <p:cBhvr>
            <p:cTn id="35" dur="500" fill="hold"/>
            <p:tgtEl>
              <p:spTgt>
                <xsl:call-template name="sptgt"/>
                <xsl:call-template name="parpref"/>
              </p:spTgt>
            </p:tgtEl>
            <p:attrNameLst>
              <p:attrName>fill.type</p:attrName>
            </p:attrNameLst>
          </p:cBhvr>
          <p:to>
            <p:strVal val="solid"/>
          </p:to>
        </p:set>
        </xsl:when>
        <!--唐江 2013-03-17  增加细微型中的忽明忽暗效果 修改bug_2710 end-->
					
		  <!--华丽型-->
		  <xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC">
			  <p:animMotion origin="layout" path="M 0.0 0.0 L 0.0 -0.07213" pathEditMode="relative" ptsTypes="">
				  <p:cBhvr>
					  <p:cTn  dur="250" accel="50000" decel="50000" autoRev="1" fill="hold">
						  <p:stCondLst>
							  <p:cond delay="0"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_x</p:attrName>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animMotion>
			  <p:animRot by="1500000">
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'125'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="0"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot by="-1500000">
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'125'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="125"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot by="-1500000">
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'125'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="250"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot by="1500000">
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'125'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="375"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
		  </xsl:when>
		  
		 <!--温和型 李娟 2012 0508-->
		  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7">
			  <p:animRot>
				  <xsl:attribute name="by">120000</xsl:attribute>
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  
							  <xsl:choose>
								  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7/@速度_6B44='medium'">
									  <xsl:attribute name="dur">
										  <xsl:value-of select="'200'"/>
									  </xsl:attribute>
								  </xsl:when>
								  <xsl:otherwise>
									  <xsl:call-template name="dur"/>
								  </xsl:otherwise> 
							  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="0"/>
						  </p:stCondLst>
						  
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot>
				  <xsl:attribute name="by">-240000</xsl:attribute>
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'200'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="200"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot>
				  <xsl:attribute name="by">240000</xsl:attribute>
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'200'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  
							  <p:stCondLst>
								  <p:cond delay="400"/>
							  </p:stCondLst>
						 
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot>
				  <xsl:attribute name="by">-240000</xsl:attribute>
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'200'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="600"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
			  <p:animRot>
				  <xsl:attribute name="by">120000</xsl:attribute>
				  <p:cBhvr>
					  <p:cTn fill="hold">
						  <xsl:choose>
							  <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7/@速度_6B44='medium'">
								  <xsl:attribute name="dur">
									  <xsl:value-of select="'200'"/>
								  </xsl:attribute>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:call-template name="dur"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <p:stCondLst>
							  <p:cond delay="800"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>r</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
			  </p:animRot>
		  </xsl:when>
        <!-- 基本型 -->
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B">
          <xsl:variable name="angle">
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B/@预定义角度_6B9D='quarter spin'">5400000</xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B/@预定义角度_6B9D='half spin'">10800000</xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B/@预定义角度_6B9D='full spin'">21600000</xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B/@预定义角度_6B9D='two spins'">43200000</xsl:when>
              <xsl:otherwise>21600000</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:animRot>
            <xsl:attribute name="by">
              <xsl:choose>
                <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B/@是否顺时针方向_6B9C='false'">
                  <xsl:value-of select="concat('-',$angle)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$angle"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>r</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
          </p:animRot>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3">
          <xsl:variable name="dur">
            <xsl:choose>
              <xsl:when test="string(演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3/@期间_6BA6*1000) != 'NaN'">
                <xsl:value-of select="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3/@期间_6BA6*1000"/>
              </xsl:when>
              <xsl:otherwise>indefinite</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="opacity">
            <!--2011-4-13 luowentian-->
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3/@预定义透明度_6BA4='25'">
                <xsl:value-of select="'0.75'"/>
              </xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3/@预定义透明度_6BA4='75'">
                <xsl:value-of select="'0.25'"/>
              </xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3/@预定义透明度_6BA4='100'">
                <xsl:value-of select="'0'"/>
              </xsl:when>
              <xsl:otherwise>0.50</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:set>
            <p:cBhvr rctx="PPT">
              <p:cTn>
                <xsl:attribute name="dur">
                  <xsl:value-of select="$dur"/>
                </xsl:attribute>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.opacity</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal>
                <xsl:attribute name="val">
                  <xsl:value-of select="$opacity"/>
                </xsl:attribute>
              </p:strVal>
            </p:to>
          </p:set>
          <p:animEffect filter="image">
            <xsl:attribute name="prLst">
              <xsl:value-of select="contains('opacity: ',$opacity)"/>
            </xsl:attribute>
            <p:cBhvr rctx="IE">
              <p:cTn>
                <xsl:attribute name="dur">
                  <xsl:value-of select="$dur"/>
                </xsl:attribute>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
        </xsl:when>
        <!--luowentian,转换为加粗显示-->
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字形_6B96">
          <!--p:animClr clrSpc="rgb">
            <p:cBhvr override="childStyle">
              <p:cTn dur="2000" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.color</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <a:schemeClr val="accent2"/>
            </p:to>
          </p:animClr-->
          <!--永中Bug 转为 PPTX 默认-->
          <p:set>
            <p:cBhvr override="childStyle">
              <p:cTn dur="indefinite"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.fontWeight</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="bold"/>
            </p:to>
          </p:set>         
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字体颜色_6BA2">
          <p:animClr clrSpc="rgb">
            <p:cBhvr override="childStyle">
              <p:cTn dur="2000" fill="hold"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.color</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <a:srgbClr>
                <!--2011-5-9罗文甜修改字体颜色-->
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字体颜色_6BA2/@颜色_6B95,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </p:to>
          </p:animClr>
        </xsl:when>
        <!--luowentian 转换放大，缩小-->
        <!--xsl:when test="演:效果/演:强调/演:更改字号">
          <xsl:variable name="to">
            <xsl:choose>
              <xsl:when test="演:效果/演:强调/演:更改字号/@演:预定义尺寸='tiny'">0.25</xsl:when>
              <xsl:when test="演:效果/演:强调/演:更改字号/@演:预定义尺寸='larger'">1.5</xsl:when>
              <xsl:when test="演:效果/演:强调/演:更改字号/@演:预定义尺寸='huge'">4</xsl:when>
              <xsl:otherwise>0.50</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:anim calcmode="lin" valueType="num">
            <xsl:attribute name="to">
              <xsl:value-of select="$to"/>
            </xsl:attribute>
            <p:cBhvr override="childStyle">
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.fontSize</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
          </p:anim>
        </xsl:when-->
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改线条颜色_6B94">
          <p:animClr clrSpc="rgb">
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>stroke.color</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改线条颜色_6B94/@颜色_6B95,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </p:to>
          </p:animClr>
          <p:set>
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>stroke.on</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="true"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改填充颜色_6B9F">
          <p:animClr clrSpc="rgb">
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>fillcolor</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改填充颜色_6B9F/@颜色_6B95,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </p:to>
          </p:animClr>
          <p:set>
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>fill.type</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="solid"/>
            </p:to>
          </p:set>
          <p:set>
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>fill.on</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="true"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!-- 默认 放大缩小 -->
        <xsl:otherwise>
          <xsl:variable name="size">
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:缩放_6B72/@预定义尺寸_6B92='tiny'"   >25000</xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:缩放_6B72/@预定义尺寸_6B92='smaller'">50000</xsl:when>
              <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:缩放_6B72/@预定义尺寸_6B92='huge'"   >400000</xsl:when>
              <xsl:otherwise>150000</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:animScale>
            <p:cBhvr>
              <p:cTn fill="hold">
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
            <p:by>
              <!--2011-4-13 luowentian-->
              <xsl:choose>
                <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:缩放_6B72/@方向_6B91='horizontal'">
                  <xsl:attribute name="x">
                    <xsl:value-of select="$size"/>
                  </xsl:attribute>
                  <xsl:attribute name="y">100000</xsl:attribute>
                </xsl:when>
                <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:缩放_6B72/@方向_6B91='vertical'">
                  <xsl:attribute name="x">100000</xsl:attribute>
                  <xsl:attribute name="y">
                    <xsl:value-of select="$size"/>
                  </xsl:attribute>               
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="x">
                    <xsl:value-of select="$size"/>
                  </xsl:attribute>
                  <xsl:attribute name="y">
                    <xsl:value-of select="$size"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </p:by>
          </p:animScale>
        </xsl:otherwise>
      </xsl:choose>
    </p:childTnLst>
	  <!--显著小这块调用，否则报错 李娟 2012.02.20-->
    <xsl:call-template name="afterAnim">
      <xsl:with-param name="id" select="$id"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="present_id_type_emph">
    <xsl:choose>
		<!--华丽型-->
		<xsl:when test="演:效果_6B40/演:强调_6B07/演:华丽型_6C22/演:波浪形_6BBC">
			<xsl:attribute name="presetID">34</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		<xsl:when test="演:效果_6B40/演:退出_6BBE/演:华丽型_6C22/演:弹跳_6B7D">
			<xsl:attribute name="presetID">26</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
		
		<!--细微型-->
		<xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:添加下划线_6BB4">
			<xsl:attribute name="presetID">18</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
      <!--唐江 2013-03-19  增加细微型中的忽明忽暗、补色、加深效果 修改bug_2710 start-->
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:忽明忽暗_6BB2">
        <xsl:attribute name="presetID">24</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:补色_6BAE">
        <xsl:attribute name="presetID">21</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:细微型_6C20/演:加深_6BAC">
        <xsl:attribute name="presetID">24</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-19  增加细微型中的忽明忽暗、补色、加深效果 修改bug_2710  end-->
      
      
      <!-- 基本型 -->
      
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:透明_6BA3">
        <xsl:attribute name="presetID">9</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字形_6B96">
        <xsl:attribute name="presetID">5</xsl:attribute>
        <xsl:attribute name="presetSubtype">1</xsl:attribute>
      </xsl:when-->
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字体颜色_6BA2 or 演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字形_6B96">
        <xsl:attribute name="presetID">3</xsl:attribute>
        <xsl:attribute name="presetSubtype">2</xsl:attribute>
      </xsl:when>
      <!--xsl:when test="演:效果/演:强调/演:更改字号">
        <xsl:attribute name="presetID">4</xsl:attribute>
        <xsl:attribute name="presetSubtype">2</xsl:attribute>
      </xsl:when-->
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改线条颜色_6B94">
        <xsl:attribute name="presetID">7</xsl:attribute>
        <xsl:attribute name="presetSubtype">2</xsl:attribute>
      </xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改填充颜色_6B9F">
        <xsl:attribute name="presetID">1</xsl:attribute>
        <xsl:attribute name="presetSubtype">2</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-19  增加基本型中的陀螺旋效果 修改bug_2710 start-->
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:陀螺旋_6B9B">
        <xsl:attribute name="presetID">8</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-19  增加基本型中的陀螺旋效果 修改bug_2710  end-->
      
		<!--温和型-->
		<xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:跷跷板_6BB7">
			<xsl:attribute name="presetID">32</xsl:attribute>
			<xsl:attribute name="presetSubtype">0</xsl:attribute>
		</xsl:when>
      <!--唐江 2013-03-19  增加温和型中的闪动效果 修改bug_2710 start-->
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:温和型_6C21/演:闪动_6BB6">
        <xsl:attribute name="presetID">27</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:when>
      <!--唐江 2013-03-19  增加温和型中的闪动效果 修改bug_2710  end-->
		
      <!-- 默认 放大缩小 -->
      <xsl:otherwise>
        <xsl:attribute name="presetID">6</xsl:attribute>
        <xsl:attribute name="presetSubtype">0</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="animMain_exit">
    <xsl:param name="id"/>
    <p:childTnLst>
      <xsl:choose>
		  <!--温和型-->
		  <xsl:when test="演:效果_6B40/演:退出_6BBE/演:温和型_6C25/演:下降_6B78">
			  <p:animEffect transition="out" filter="fade">
				  <p:cBhvr>
					  <p:cTn  dur="1000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
			  </p:animEffect>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="1000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_x</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_x"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_x"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="1000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y+.1"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:set>
				  <p:cBhvr>
					  <p:cTn  dur="1" fill="hold">
						  <p:stCondLst>
							  <p:cond delay="999"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>style.visibility</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:to>
					  <p:strVal val="hidden"/>
				  </p:to>
			  </p:set>
					  
		  </xsl:when>
		  <!--细微型-->
		  <xsl:when test="演:效果_6B40/演:退出_6BBE/演:细微型_6C24/演:渐变_6B67">
			  <p:animEffect transition="out" filter="fade">
				  <p:cBhvr>
					  <p:cTn  dur="500"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
			  </p:animEffect>
			  <p:set>
				  <p:cBhvr>
					  <p:cTn  dur="1" fill="hold">
						  <p:stCondLst>
							  <p:cond delay="499"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>style.visibility</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:to>
					  <p:strVal val="hidden"/>
				  </p:to>
			  </p:set>
		  </xsl:when>
		  <xsl:when test="演:效果_6B40/演:退出_6BBE/演:细微型_6C24/演:渐变式回旋_6B69">
			  <p:animEffect transition="out" filter="fade">
				  <p:cBhvr>
					  <p:cTn  dur="2000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
			  </p:animEffect>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="2000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_w</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="5000">
						  <p:val>
							  <p:strVal val="0.92*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="10000">
						  <p:val>
							  <p:strVal val="0.71*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="15000">
						  <p:val>
							  <p:strVal val="0.38*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="20000">
						  <p:val>
							  <p:fltVal val="0"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="25000">
						  <p:val>
							  <p:strVal val="-0.38*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="30000">
						  <p:val>
							  <p:strVal val="-0.71*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="35000">
						  <p:val>
							  <p:strVal val="-0.92*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="40000">
						  <p:val>
							  <p:strVal val="-ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="45000">
						  <p:val>
							  <p:strVal val="-0.92*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="50000">
						  <p:val>
							  <p:strVal val="-0.71*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="55000">
						  <p:val>
							  <p:strVal val="-0.38*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="60000">
						  <p:val>
							  <p:fltVal val="0"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="65000">
						  <p:val>
							  <p:strVal val="0.38*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="70000">
						  <p:val>
							  <p:strVal val="0.71*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="75000">
						  <p:val>
							  <p:strVal val="0.92*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="80000">
						  <p:val>
							  <p:strVal val="ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="85000">
						  <p:val>
							  <p:strVal val="0.92*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="90000">
						  <p:val>
							  <p:strVal val="0.71*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="95000">
						  <p:val>
							  <p:strVal val="0.38*ppt_w"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:fltVal val="0"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="2000"/>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_h</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_h"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_h"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:set>
				  <p:cBhvr>
					  <p:cTn  dur="1" fill="hold">
						  <p:stCondLst>
							  <p:cond delay="1999"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>style.visibility</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:to>
					  <p:strVal val="hidden"/>
				  </p:to>
			  </p:set>
		  </xsl:when>
		  
		  <!--华丽型-->
		  <xsl:when test="演:效果_6B40/演:退出_6BBE/演:华丽型_6C26/演:弹跳_6B7D">
			  <p:animEffect transition="out" filter="wipe(down)">
				  <p:cBhvr>
					  <p:cTn  dur="180" accel="50000">
						  <p:stCondLst>
							  <p:cond delay="1820"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
			  </p:animEffect>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="1822" tmFilter="0,0; 0.14,0.31; 0.43,0.73; 0.71,0.91; 1.0,1.0">
						  <p:stCondLst>
							  <p:cond delay="0"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_x</p:attrName>
					  </p:attrNameLst>
					  
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_x"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="#ppt_x+0.25"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="178">
						  <p:stCondLst>
							  <p:cond delay="1822"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_x</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_x"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_x"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="664" tmFilter="0.0,0.0;0.25,0.07;0.50,0.2;0.75,0.467;1.0,1.0">
						  <p:stCondLst>
							  <p:cond delay="1822"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="5000">
						  <p:val>
							  <p:strVal val="ppt_y+0.026"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="10000">
						  <p:val>
							  <p:strVal val="ppt_y+0.052"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="15000">
						  <p:val>
							  <p:strVal val="ppt_y+0.078"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="20000">
						  <p:val>
							  <p:strVal val="ppt_y+0.103"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="30000">
						  <p:val>
							  <p:strVal val="ppt_y+0.151"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="40000">
						  <p:val>
							  <p:strVal val="ppt_y+0.196"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="50000">
						  <p:val>
							  <p:strVal val="ppt_y+0.236"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="60000">
						  <p:val>
							  <p:strVal val="ppt_y+0.270"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="70000">
						  <p:val>
							  <p:strVal val="ppt_y+0.297"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="80000">
						  <p:val>
							  <p:strVal val="ppt_y+0.317"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="90000">
						  <p:val>
							  <p:strVal val="ppt_y+0.329"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y+0.333"/>
						  </p:val>
					  </p:tav>
					  
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="664" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
						  <p:stCondLst>
							  <p:cond delay="664"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="10000">
						  <p:val>
							  <p:strVal val="ppt_y-0.034"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="20000">
						  <p:val>
							  <p:strVal val="ppt_y-0.065"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="30000">
						  <p:val>
							  <p:strVal val="ppt_y-0.090"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="40000">
						  <p:val>
							  <p:strVal val="ppt_y-0.106"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="50000">
						  <p:val>
							  <p:strVal val="ppt_y-0.111"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="60000">
						  <p:val>
							  <p:strVal val="ppt_y-0.106"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="70000">
						  <p:val>
							  <p:strVal val="ppt_y-0.090"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="80000">
						  <p:val>
							  <p:strVal val="ppt_y-0.065"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="90000">
						  <p:val>
							  <p:strVal val="ppt_y-0.034"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="332" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
						  <p:stCondLst>
							  <p:cond delay="1324"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="10000">
						  <p:val>
							  <p:strVal val="ppt_y-0.011"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="20000">
						  <p:val>
							  <p:strVal val="ppt_y-0.022"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="30000">
						  <p:val>
							  <p:strVal val="ppt_y-0.030"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="40000">
						  <p:val>
							  <p:strVal val="ppt_y-0.035"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="50000">
						  <p:val>
							  <p:strVal val="ppt_y-0.037"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="60000">
						  <p:val>
							  <p:strVal val="ppt_y-0.035"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="70000">
						  <p:val>
							  <p:strVal val="ppt_y-0.030"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="80000">
						  <p:val>
							  <p:strVal val="ppt_y-0.022"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="90000">
						  <p:val>
							  <p:strVal val="ppt_y-0.011"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="164" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">
						  <p:stCondLst>
							  <p:cond delay="1656"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="10000">
						  <p:val>
							  <p:strVal val="ppt_y-0.004"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="20000">
						  <p:val>
							  <p:strVal val="ppt_y-0.007"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="30000">
						  <p:val>
							  <p:strVal val="ppt_y-0.010"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="40000">
						  <p:val>
							  <p:strVal val="ppt_y-0.012"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="50000">
						  <p:val>
							  <p:strVal val="ppt_y-0.0123"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="60000">
						  <p:val>
							  <p:strVal val="ppt_y-0.012"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="70000">
						  <p:val>
							  <p:strVal val="ppt_y-0.010"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="80000">
						  <p:val>
							  <p:strVal val="ppt_y-0.007"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="90000">
						  <p:val>
							  <p:strVal val="ppt_y-0.004"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
			  </p:anim>
			  <p:anim calcmode="lin" valueType="num">
				  <p:cBhvr>
					  <p:cTn  dur="180" accel="50000">
						  <p:stCondLst>
							  <p:cond delay="1820"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>ppt_y</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:tavLst>
					  <p:tav tm="0">
						  <p:val>
							  <p:strVal val="ppt_y"/>
						  </p:val>
					  </p:tav>
					  <p:tav tm="100000">
						  <p:val>
							  <p:strVal val="ppt_y+ppt_h"/>
						  </p:val>
					  </p:tav>
				  </p:tavLst>
				  
			  </p:anim>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="26">
						  <p:stCondLst>
							  <p:cond delay="620"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="60000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="166" decel="50000">
						  <p:stCondLst>
							  <p:cond delay="646"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="100000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="26">
						  <p:stCondLst>
							  <p:cond delay="1312"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="80000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="166" decel="50000">
						  <p:stCondLst>
							  <p:cond delay="1338"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="100000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="26">
						  <p:stCondLst>
							  <p:cond delay="1642"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="90000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="166" decel="50000">
						  <p:stCondLst>
							  <p:cond delay="1668"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="100000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="26">
						  <p:stCondLst>
							  <p:cond delay="1808"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="95000"/>
			  </p:animScale>
			  <p:animScale>
				  <p:cBhvr>
					  <p:cTn  dur="166" decel="50000">
						  <p:stCondLst>
							  <p:cond delay="1834"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
				  </p:cBhvr>
				  <p:to x="100000" y="100000"/>
			  </p:animScale>
			  <p:set>
				  <p:cBhvr>
					  <p:cTn  dur="1" fill="hold">
						  <p:stCondLst>
							  <p:cond delay="1999"/>
						  </p:stCondLst>
					  </p:cTn>
					  <p:tgtEl>
						  <p:spTgt>
							  <xsl:call-template name="sptgt"/>
							  <xsl:call-template name="parpref"/>
						  </p:spTgt>
					  </p:tgtEl>
					  <p:attrNameLst>
						  <p:attrName>style.visibility</p:attrName>
					  </p:attrNameLst>
				  </p:cBhvr>
				  <p:to>
					  <p:strVal val="hidden"/>
				  </p:to>
			  </p:set>
		  </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:圆形扩展_6B56">
          <p:animEffect transition="out" filter="circle(in)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:圆形扩展_6B56">
              <xsl:attribute name="filter">circle(out)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="1999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:扇形展开_6B61">
          <p:animEffect transition="out" filter="wedge">
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn  dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="1999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:消失_6BC7">
          <p:set>
            <p:cBhvr>
              <p:cTn  dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:向外溶解_6BC5">
          <p:animEffect transition="out" filter="dissolve">
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--luowentian:翻转旋转-->
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:随机效果_6B55">
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr>
              <p:cTn dur="1000"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_w</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_w"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:fltVal val="0"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr>
              <p:cTn dur="1000"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_h</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_h"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:fltVal val="0"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr>
              <p:cTn dur="1000"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.rotation</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:fltVal val="0"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:fltVal val="90"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:animEffect transition="out" filter="fade">
            <p:cBhvr>
              <p:cTn dur="1000"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:随机线条_6B62">
          <p:animEffect transition="out" filter="randombar(vertical)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:随机线条_6B62/@方向_6B45='horizontal'">
              <xsl:attribute name="filter">randombar(horizontal)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:十字形扩展_6B53">
          <p:animEffect transition="out" filter="plus(in)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/@方向_6B48='out'">
              <xsl:attribute name="filter">plus(out)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="1999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--luowentian :消失-->
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:闪烁一次_6B51">
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="0"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--2011-4-7-luowentian:修改切出效果-->
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4">
          <xsl:variable name="dir" select="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4/@方向_6BC2"/>
          <xsl:variable name="ppt">
            <xsl:choose>
              <xsl:when test="contains($dir,'left') or contains($dir,'right')">ppt_x</xsl:when>
              <xsl:when test="contains($dir,'top') or contains($dir,'bottom')">ppt_y</xsl:when>
              <xsl:otherwise>ppt_x</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="pptval">
            <xsl:choose>
              <xsl:when test="contains($dir,'right')">#ppt_x+#ppt_w*1.125000</xsl:when>
              <xsl:when test="contains($dir,'left')">#ppt_x-#ppt_w*1.125000</xsl:when>
              <xsl:when test="contains($dir,'bottom')">#ppt_y+#ppt_h*1.125000</xsl:when>
              <xsl:when test="contains($dir,'top')">#ppt_y-#ppt_h*1.125000</xsl:when>
            </xsl:choose>
          </xsl:variable>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr additive="base">
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>
                  <xsl:value-of select="$ppt"/>
                </p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal>
                    <xsl:attribute name="val">
                      <xsl:value-of select="concat('#',$ppt)"/>
                    </xsl:attribute>
                  </p:strVal>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:strVal>
                    <xsl:attribute name="val">
                      <xsl:value-of select="$pptval"/>
                    </xsl:attribute>
                  </p:strVal>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:animEffect transition="out">
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4/@方向_6BC2='to left'">
                <xsl:attribute name="filter">wipe(right)</xsl:attribute>
              </xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4/@方向_6BC2='to right'">
                <xsl:attribute name="filter">wipe(left)</xsl:attribute>
              </xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4/@方向_6BC2='to top'">
                <xsl:attribute name="filter">wipe(up)</xsl:attribute>
              </xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:切出_6BC4/@方向_6BC2='to bottom'">
                <xsl:attribute name="filter">wipe(down)</xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:棋盘_6B4E">
          <p:animEffect transition="out" filter="checkerboard(across)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:棋盘_6B4E/@方向_6B48='down'">
              <xsl:attribute name="filter">checkerboard(down)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:劈裂_6B5E">
          <p:animEffect transition="out" filter="barn(inHorizontal)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:劈裂_6B5E/@方向_6B5F='vertical out'">
              <xsl:attribute name="filter">barn(outVertical)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:劈裂_6B5E/@方向_6B5F='vertical in'">
              <xsl:attribute name="filter">barn(inVertical)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:劈裂_6B5E/@方向_6B5F='horizontal out'">
              <xsl:attribute name="filter">barn(outHorizontal)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--2010-10-25罗文甜：修改辐射状属性为轮辐-->
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B">
          <p:animEffect transition="out" filter="wheel(4)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@轮辐_6B4D='1'">
              <xsl:attribute name="filter">wheel(1)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@轮辐_6B4D='2'">
              <xsl:attribute name="filter">wheel(2)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@轮辐_6B4D='3'">
              <xsl:attribute name="filter">wheel(3)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@轮辐_6B4D='8'">
              <xsl:attribute name="filter">wheel(8)</xsl:attribute>
            </xsl:if>
            <!--xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@演:轮辐='1'">
              <xsl:attribute name="filter">wheel(1)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@演:轮辐='2'">
              <xsl:attribute name="filter">wheel(2)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@演:轮辐='3'">
              <xsl:attribute name="filter">wheel(3)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:轮子_6B4B/@演:轮辐='8'">
              <xsl:attribute name="filter">wheel(8)</xsl:attribute>
            </xsl:if-->
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="1999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:菱形_6B5D">
          <p:animEffect transition="out" filter="diamond(in)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:菱形_6B5D/@方向_6B48='out'">diamond(out)</xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="1999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:阶梯状_6B49">
          <p:animEffect transition="out" filter="strips(downLeft)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:阶梯状_6B49/@方向_6B4A='left up'">
              <xsl:attribute name="filter">strips(upLeft)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:阶梯状_6B49/@方向_6B4A='right down'">
              <xsl:attribute name="filter">strips(downRight)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:阶梯状_6B49/@方向_6B4A='right up'">
              <xsl:attribute name="filter">strips(upRight)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1">
          <xsl:variable name="x">
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to top'">ppt_x</xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to left'">0-ppt_w/2</xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to right'">1+ppt_w/2</xsl:when>
              <xsl:otherwise>ppt_x</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:choose>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to top'">0-ppt_h/2</xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to left'">ppt_y</xsl:when>
              <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:缓慢移出_6BC1/@方向_6BC2='to right'">ppt_y</xsl:when>
              <xsl:otherwise>1+ppt_h/2</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr additive="base">
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_x</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_x"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:strVal>
                    <xsl:attribute name="val">
                      <xsl:value-of select="$x"/>
                    </xsl:attribute>
                  </p:strVal>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr additive="base">
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_y</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_y"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:strVal>
                    <xsl:attribute name="val">
                      <xsl:value-of select="$y"/>
                    </xsl:attribute>
                  </p:strVal>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:set>
            <p:cBhvr>
              <p:cTn id="8" dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="4999"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:盒状_6B47">
          <p:animEffect transition="out" filter="box(in)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:盒状_6B47/@方向_6B48='out'">
              <xsl:attribute name="filter">box(out)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:飞出_6BBF">
          <!--xsl:when test="演:效果/演:退出/飞出 or 演:效果/演:退出//@uof:locID='p0101'"-->
          <!--<xsl:variable name="dir" select="演:效果/演:退出//@演:方向"/>-->
			<xsl:variable name="dir" select="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:飞出_6BBF/@方向_6BC0"/>
          <xsl:variable name="x">
            <xsl:choose>
              <xsl:when test="contains($dir,'left')">0-#ppt_w/2</xsl:when>
              <xsl:when test="contains($dir,'right')">1+#ppt_w/2</xsl:when>
              <xsl:otherwise>#ppt_x</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:choose>
              <xsl:when test="contains($dir,'top')">0-#ppt_h/2</xsl:when>
              <xsl:when test="contains($dir,'bottom')">1+#ppt_h/2</xsl:when>
              <xsl:otherwise>#ppt_x</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr additive="base">
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_x</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_x"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:strVal val="$x"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr additive="base">
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_y</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_y"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:strVal val="$y"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:set>
            <p:cBhvr>
              <p:cTn id="8" dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:擦除_6B57">
          <p:animEffect transition="out" filter="wipe(down)">
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:擦除_6B57/@方向_6B58='from top'">
              <xsl:attribute name="filter">wipe(up)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:擦除_6B57/@方向_6B58='from left'">
              <xsl:attribute name="filter">wipe(left)</xsl:attribute>
            </xsl:if>
            <xsl:if test="演:效果_6B40/演:退出_6BBE/演:基本型_6C23/演:擦除_6B57/@方向_6B58='from right'">
              <xsl:attribute name="filter">wipe(right)</xsl:attribute>
            </xsl:if>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--唐江 2013-03-22  增加细微型中的渐变式缩放效果 修改bug_2710  start-->
        <xsl:when test="演:效果_6B40/演:退出_6BBE/演:细微型_6C24/演:渐变式缩放_6B6A">
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr>
              <p:cTn dur="500"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_w</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_w"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:fltVal val="0"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:anim calcmode="lin" valueType="num">
            <p:cBhvr>
              <p:cTn id="14" dur="500"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>ppt_h</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:tavLst>
              <p:tav tm="0">
                <p:val>
                  <p:strVal val="ppt_h"/>
                </p:val>
              </p:tav>
              <p:tav tm="100000">
                <p:val>
                  <p:fltVal val="0"/>
                </p:val>
              </p:tav>
            </p:tavLst>
          </p:anim>
          <p:animEffect transition="out" filter="fade">
            <p:cBhvr>
              <p:cTn id="15" dur="500"/>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn id="16" dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:when>
        <!--唐江 2013-03-22  增加细微型中的渐变式缩放效果 修改bug_2710  end-->
        <xsl:otherwise>
          <p:animEffect transition="out">
            <!--2011-4-8 luowentian -->
            <xsl:choose>
              <xsl:when test="演:效果/演:退出//@演:方向='vertical'">
                <xsl:attribute name="filter">blinds(vertical)</xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="filter">blinds(horizontal)</xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
            <p:cBhvr>
              <p:cTn>
                <xsl:call-template name="dur"/>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
            </p:cBhvr>
          </p:animEffect>
          <p:set>
            <p:cBhvr>
              <p:cTn  dur="1" fill="hold">
                <p:stCondLst>
                  <p:cond delay="499"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt>
                  <xsl:call-template name="sptgt"/>
                  <xsl:call-template name="parpref"/>
                </p:spTgt>
              </p:tgtEl>
              <p:attrNameLst>
                <p:attrName>style.visibility</p:attrName>
              </p:attrNameLst>
            </p:cBhvr>
            <p:to>
              <p:strVal val="hidden"/>
            </p:to>
          </p:set>
        </xsl:otherwise>
      </xsl:choose>
    </p:childTnLst>
    <!--2011-5-16 luowentian-->
	  <!--注销这块引用 否则 报错 打不开 李娟 2012.02.20-->
    <xsl:call-template name="afterAnim">
      <xsl:with-param name="id" select="$id"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="animMain_path">
    <xsl:param name="id"/>
    <p:childTnLst>
      <p:animMotion origin="layout" ptsTypes="">
        <xsl:attribute name="path">
          <!--2011-3-15罗文甜，修改动画路径的转换，按照金山来实现-->
          <!--2015.5.11guoyongbin  修改修改动画路径的转换-->
          <!--xsl:value-of select="演:效果_6B40/演:动作路径_6BD1/@路径_6BD6"/-->
          <xsl:value-of select="'M -3.95833E-6 4.07407E-6 L -0.00117 0.13611'"/>
          <!--end,2015.5.11guoyongbin.-->
          <!--xsl:call-template name="path">
            <xsl:with-param name="val" select="演:效果/演:动作路径/@演:路径"/>
            <xsl:with-param name="w" select="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页面设置集/演:页面设置/演:纸张/@uof:宽度"/>
            <xsl:with-param name="h" select="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页面设置集/演:页面设置/演:纸张/@uof:高度"/>
            <xsl:with-param name="step" select="1"/>
          </xsl:call-template-->
        </xsl:attribute>
        <!--2010-11-1罗文甜：增加路径是否锁定-->
        <xsl:choose>
          <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@是否锁定_6BD7='true'">
            <xsl:attribute name="pathEditMode">fixed</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="pathEditMode">relative</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <p:cBhvr>
          <p:cTn fill="hold">
            <!--2010-11-1罗文甜：增加反转路径方向-->
            <xsl:if test="演:效果_6B40/演:动作路径_6BD1/@是否反转路径方向_6BDC='true'">
              <xsl:attribute name="spd">
                <xsl:value-of select="'-100000'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:call-template name="dur"/>
          </p:cTn>
          <p:tgtEl>
            <p:spTgt>
              <xsl:call-template name="sptgt"/>
              <xsl:call-template name="parpref"/>
            </p:spTgt>
          </p:tgtEl>
          <p:attrNameLst>
            <p:attrName>ppt_x</p:attrName>
            <p:attrName>ppt_y</p:attrName>
          </p:attrNameLst>
        </p:cBhvr>
      </p:animMotion>
    </p:childTnLst>
	  <!--先注销这部分 否则 报错李娟 2012.02.20-->
    <xsl:call-template name="afterAnim">
      <xsl:with-param name="id" select="$id"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="dur">
    <xsl:attribute name="dur">
      <xsl:choose><!--dingshi-->
        <!--<xsl:when test="演:定时/@演:速度='very slow'">5000</xsl:when>
        <xsl:when test="演:定时/@演:速度='slow'"     >3000</xsl:when>
        <xsl:when test="演:定时/@演:速度='medium'">2000</xsl:when>
        <xsl:when test="演:定时/@演:速度='fast'">1000</xsl:when>
        <xsl:when test="演:定时/@演:速度='very fast'">500</xsl:when>-->

      <!--start liuyin 20130107 修改动画序列的计时转换错误。-->
      <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@速度_6B44='very-slow'">5000</xsl:when>
		  <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@速度_6B44='slow'">3000</xsl:when>
		  <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@速度_6B44='medium'">2000</xsl:when>
		  <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@速度_6B44='fast'">1000</xsl:when>
		  <xsl:when test="演:效果_6B40/演:动作路径_6BD1/@速度_6B44='very-fast'">500</xsl:when>
      <!--end liuyin 20130107 修改动画序列的计时转换错误。-->
        
      <!--start liuyin 20130327 修改bug2810 动画速度错误。-->
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='very-fast'">500</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:百叶窗_6B43/@速度_6B44='very-fast'">500</xsl:when>

      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='very-fast'">500</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:盒状_6B47/@速度_6B44='very-fast'">500</xsl:when>

      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='very-fast'">500</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:效果_6B40/演:进入_6B41/演:基本型_6B42/演:扇形展开_6B61/@速度_6B44='very-fast'">500</xsl:when>
      
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:序列_6B1B/演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='very-fast'">500</xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='very-slow'">5000</xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='slow'">3000</xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='medium'">2000</xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='fast'">1000</xsl:when>
      <xsl:when test="演:效果_6B40/演:强调_6B07/演:基本型_6C19/演:更改字号_6BA0/@速度_6B44='very-fast'">500</xsl:when>
      <!--end liuyin 20130327 修改bug2810 动画速度错误。-->
     
       <xsl:otherwise>2000</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="afterAnim">
    <xsl:param name="id"/>
    <!--2011-5-16 luowentian-->
    <xsl:if test="演:增强_6B35/演:声音_6B22 or 演:增强_6B35/演:动画播放后_6B36/演:是否播放后隐藏_6B38='true' or 演:增强_6B35/演:动画播放后_6B36/演:是否单击后隐藏_6B39='true' or 
                 演:增强_6B35/演:动画播放后_6B36/演:颜色_6B37!=''">
	<!--注销and 演:增强_6B35/演:动画播放后_6B36/@演:变暗='false' 李娟 2012.01.07-->
      <p:subTnLst>
        <xsl:if test="演:增强_6B35/演:声音_6B22">
          <xsl:choose>
			  <xsl:when test="演:增强_6B35/演:声音_6B22/@预定义声音_C631='none'">
            <!--<xsl:when test="演:增强_6B35/演:声音_6B22/@预定义声音_C631='stop-previous-sound'">-->
              <p:cmd type="evt" cmd="onstopaudio">
                <p:cBhvr>
                  <p:cTn display="0" masterRel="sameClick">
                    <p:stCondLst>
                      <p:cond evt="begin" delay="0">
                        <p:tn>
                          <xsl:attribute name="val">
                            <xsl:value-of select="$id"/>
                          </xsl:attribute>
                        </p:tn>
                      </p:cond>
                    </p:stCondLst>
                  </p:cTn>
                  <p:tgtEl>
                    <p:sldTgt/>
                  </p:tgtEl>
                </p:cBhvr>
              </p:cmd>
            </xsl:when>
            <xsl:otherwise>
              <p:audio>
                <p:cMediaNode>
                  <p:cTn display="0" masterRel="sameClick">
                    <p:stCondLst>
                      <p:cond evt="begin" delay="0">
                        <p:tn>
                          <xsl:attribute name="val">
                            <xsl:value-of select="$id"/>
                          </xsl:attribute>
                        </p:tn>
                      </p:cond>
                    </p:stCondLst>
                    <p:endCondLst>
                      <p:cond evt="onStopAudio" delay="0">
                        <p:tgtEl>
                          <p:sldTgt/>
                        </p:tgtEl>
                      </p:cond>
                    </p:endCondLst>
                  </p:cTn>
                  <p:tgtEl>
                    <p:sndTgt>
                      <xsl:attribute name="r:embed">
                        <xsl:value-of select="演:增强_6B35/演:声音_6B22/@自定义声音_C632"/>
                      </xsl:attribute>
                    </p:sndTgt>
                  </p:tgtEl>
                </p:cMediaNode>
              </p:audio>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>

        <xsl:choose>
          <xsl:when test="演:增强_6B35/演:动画播放后_6B36/演:是否播放后隐藏_6B38='true'">
            <p:set>
              <p:cBhvr override="childStyle">
                <p:cTn dur="1" fill="hold" display="0" masterRel="sameClick" afterEffect="1">
                  <p:stCondLst>
                    <p:cond evt="end" delay="0">
                      <p:tn>
                        <xsl:attribute name="val">
                          <xsl:value-of select="$id"/>
                        </xsl:attribute>
                      </p:tn>
                    </p:cond>
                  </p:stCondLst>
                </p:cTn>
                <p:tgtEl>
                  <p:spTgt>
                    <xsl:call-template name="sptgt"/>
                    <xsl:call-template name="parpref"/>
                  </p:spTgt>
                </p:tgtEl>
                <p:attrNameLst>
                  <p:attrName>style.visibility</p:attrName>
                </p:attrNameLst>
              </p:cBhvr>
              <p:to>
                <p:strVal val="hidden"/>
              </p:to>
            </p:set>
          </xsl:when>
          <xsl:when test="演:增强_6B35/演:动画播放后_6B36/演:是否单击后隐藏_6B39='true'">
            <p:set>
              <p:cBhvr override="childStyle">
                <p:cTn dur="1" fill="hold" display="0" masterRel="nextClick" afterEffect="1">
                  <p:stCondLst>
                    <p:cond evt="end" delay="0">
                      <p:tn val="5"/>
                    </p:cond>
                  </p:stCondLst>
                </p:cTn>
                <p:tgtEl>
                  <p:spTgt>
                    <xsl:call-template name="sptgt"/>
                    <xsl:call-template name="parpref"/>
                  </p:spTgt>
                </p:tgtEl>
                <p:attrNameLst>
                  <p:attrName>style.visibility</p:attrName>
                </p:attrNameLst>
              </p:cBhvr>
              <p:to>
                <p:strVal val="hidden"/>
              </p:to>
            </p:set>
          </xsl:when>
          <xsl:when test="演:增强_6B35/演:动画播放后_6B36/演:颜色_6B37!=''">
			  <!--住消and 演:增强/演:动画播放后/@演:变暗='false' 李娟 2012.01.07-->

            <p:animClr>
              <p:cBhvr override="childStyle">
                <p:cTn dur="1" fill="hold" display="0" masterRel="nextClick" afterEffect="1"/>
                <p:tgtEl>
                  <p:spTgt>
                    <xsl:call-template name="sptgt"/>
                    <xsl:call-template name="parpref"/>
                  </p:spTgt>
                </p:tgtEl>
                <p:attrNameLst>
                  <p:attrName>ppt_c</p:attrName>
                </p:attrNameLst>
              </p:cBhvr>
              <p:to>
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring-after(演:增强_6B35/演:动画播放后_6B36/演:颜色_6B37,'#')"/>
                  </xsl:attribute>
                </a:srgbClr>
              </p:to>
            </p:animClr>
          </xsl:when>
        </xsl:choose>

      </p:subTnLst>
    </xsl:if>
  </xsl:template>

  <!--2014-04-19, tangjiang,  修复图片动画sptgt元素的id值错误 start -->
  <!--2015-03-23,liuyangyang,修改添加ole对象动画时，导致文档转换需要修复的错误-->
  <xsl:template name="sptgt">
    <xsl:variable name="mediaId" select="@对象引用_6C28"/>
    <xsl:variable name="graphicId" select="//图:图形_8062[@标识符_804B=$mediaId]/图:图片数据引用_8037"/>
    <xsl:choose>
      <xsl:when test="$graphicId!=''">
        <xsl:choose>
        <xsl:when test="//图:图形_8062[@标识符_804B=$mediaId]/ole/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr/@id">
          <xsl:attribute name="spid">
          <xsl:value-of select="//图:图形_8062[@标识符_804B=$mediaId]/ole/p:graphicFrame/p:nvGraphicFramePr/p:cNvPr/@id"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
        <xsl:attribute name="spid">
          <xsl:call-template name="mediaIdConvetor">
            <xsl:with-param name="mediaId" select="$graphicId"/>
          </xsl:call-template>
        </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
            <xsl:attribute name="spid">
              <xsl:call-template name="mediaIdConvetor">
                <xsl:with-param name="mediaId" select="$mediaId"/>
              </xsl:call-template>
            </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--end 2014-04-19, tangjiang,  修复图片动画sptgt元素的id值错误 -->

  <xsl:template name="parpref">
    <xsl:if test="./@段落引用_6C27">
      <xsl:variable name="tgt" select="@对象引用_6C28"/>
      <!--2010-1-5罗文甜：按照永中修改动画的段落引用-->
      <xsl:variable name="pararef" select="substring-after(@段落引用_6C27,substring-before(@段落引用_6C27,'graphc'))"/>
      <p:txEl>
        <p:pRg>
          <xsl:for-each select="//图:图形_8062[@标识符_804B=$tgt]//字:段落_416B">
            <xsl:variable name="paraid" select="substring-after(@标识符_4169,substring-before(@标识符_4169,'graphc'))"/>
            <xsl:if test="$paraid=$pararef">
              <xsl:attribute name="st">
                <xsl:value-of select="position()-1"/>
              </xsl:attribute>
              <xsl:attribute name="end">
                <xsl:value-of select="position()-1"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </p:pRg>
      </p:txEl>
      
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
