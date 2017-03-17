<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
      

                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph">
  <xsl:import href="fill.xsl"/>
	<!--11.11.10 修改动画 进入 强调 退出都是只写了基本型,细微 温和和华丽都没有写 李娟-->
  <xsl:template match="p:timing">
	  <演:动画_6B1A>

      <!--start liuyin 20121211 修改音视频丢失-->
       <!-- <xsl:apply-templates select="." mode="par"/>-->
      
      <!--2011-01-04 罗文甜：修改BUG，增加p:tnLst/标识
    <xsl:for-each select="p:tnLst//p:par[p:cTn[@presetID and @presetClass!='mediacall']]">-->
      <!--2013-1-10  修改动画延迟时间  liqiuling  start-->
      <xsl:for-each select="p:tnLst//p:par[p:cTn[@presetID]]">
        <xsl:apply-templates select="." mode="par"/>
      </xsl:for-each>
      <!--end liuyin 20121211 修改音视频丢失-->
      <!--2013-1-10  修改动画延迟时间  liqiuling  end-->
      
    </演:动画_6B1A>
  </xsl:template>

  <xsl:template match="p:par" mode="par">
    <xsl:variable name="spid" select="ancestor::p:timing//p:cond/p:tgtEl/p:spTgt/@spid"/>
	  <演:序列_6B1B>
    <!--<演:序列 uof:locID="p0043" uof:attrList="段落引用 动画对象">-->
      <xsl:call-template name="target"/>
		  <演:定时_6B2E>
		  <!--<演:定时 uof:locID="p0067" uof:attrList="事件 延时 速度 重复 回卷 触发器">-->
        <xsl:attribute name="事件_6B2F">
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@nodeType='withEffect'">with-previous</xsl:when>
            <xsl:when test="descendant::p:cTn[@presetID]/@nodeType='afterEffect'">after-previous</xsl:when>
            <xsl:otherwise>on-click</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="延时_6B30">
          <xsl:choose>
            <!--2010-11-19 罗文甜：修改BUG-->
            <!--2013-1-10 修改动画延时不正确 liqiuling start-->
            <xsl:when test="descendant::p:cTn[@presetID]/p:stCondLst/p:cond/@delay!='0'and string(descendant::p:cTn[@presetID]/p:stCondLst/p:cond/@delay div 1000)!='NaN'">
              <!--xsl:when test="string(descendant::p:cTn[@presetID]/p:stCondLst/p:cond/@delay div 1000)!='NaN'"
              <xsl:value-of select="descendant::p:cTn[@presetID]/p:stCondLst/p:cond/@delay div 1000"/>-->
              <xsl:variable name="delaytime">
                <xsl:value-of select="descendant::p:cTn[@presetID]/p:stCondLst/p:cond/@delay div 1000"/>
              </xsl:variable>
              <xsl:value-of select="concat('P0Y0M0DT0H0M',$delaytime,'S')"/>
              
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="'P0Y0M0DT0H0M0S'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!--2013-1-10 修改动画延时不正确 liqiuling start-->
			  <!--11.11.10 李娟 注释这部分-->
        <!--<xsl:attribute name="演:速度">
          <xsl:call-template name="dur">
            <xsl:with-param name="loc" select="演:速度"/>
            <xsl:with-param name="dur" select=".//*[@dur!='1']/@dur"/>
          </xsl:call-template>
        </xsl:attribute>-->
        <xsl:attribute name="重复_6B32">
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]//@repeatCount='indefinite'">
              <xsl:choose>
                <xsl:when test="descendant::p:cTn[@presetID]/p:endCondLst">until-next-click</xsl:when>
                <xsl:otherwise>until-end-of-slide</xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string(descendant::p:cTn[@presetID]/@repeatCount div 1000) != 'NaN'">
              <xsl:value-of select="descendant::p:cTn[@presetID]/@repeatCount div 1000"/>
            </xsl:when>
            <!--2010-11-19 罗文甜：修改bug-->
            <xsl:otherwise>none</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="是否回卷_6B33">
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@fill='remove'">true</xsl:when>
            <xsl:otherwise>false</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!--2010-11-2罗文甜 增加触发器-->
        <xsl:if test="ancestor::p:timing//p:cond/p:tgtEl/p:spTgt/@spid">
          <xsl:attribute name="触发对象引用_6B34">           
            <xsl:value-of select="translate(ancestor::p:timing/preceding-sibling::p:cSld/p:spTree/p:sp[p:nvSpPr/p:cNvPr/@id=$spid]/@id,':','')"/>
          </xsl:attribute>
        </xsl:if>
      <!--</演:定时>-->
		  </演:定时_6B2E>
		  <演:增强_6B35>
        <!--2013-1-10 修改动画增强效果 liqiuling start-->
		  <演:动画播放后_6B36>
        <xsl:choose>
          <!--<演:动画播放后 uof:locID="p0070" uof:attrList="颜色 变暗 播放后隐藏 单击后隐藏">-->

          <xsl:when test="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr">
            <演:颜色_6B37>
              <xsl:choose>
                <xsl:when test="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr">
                  <xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
                </xsl:when>
                <xsl:when test="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:srgbClr">
                  <xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
                </xsl:when>
                <xsl:otherwise>#ffffff</xsl:otherwise>
              </xsl:choose>
       
          </演:颜色_6B37>
          </xsl:when>
          <!--<xsl:attribute name="演:变暗">
            <xsl:choose>
              <xsl:when test="not(descendant::p:cTn[@presetID]/p:subTnLst/p:animClr)">true</xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>-->
          <!-- 李娟 11.11.17-->

          <xsl:when test="descendant::p:cTn[@presetID]/p:subTnLst/p:set/p:cBhvr[p:attrNameLst/p:attrName='style.visibility']/p:cTn/@masterRel='sameClick'">
            <演:是否播放后隐藏_6B38>true</演:是否播放后隐藏_6B38>
          </xsl:when>
          <xsl:when test="descendant::p:cTn[@presetID]/p:subTnLst/p:set/p:cBhvr[p:attrNameLst/p:attrName='style.visibility']/p:cTn/@masterRel='nextClick'">
            <演:是否单击后隐藏_6B39>true</演:是否单击后隐藏_6B39>
          </xsl:when>
        </xsl:choose>
        
				  
				</演:动画播放后_6B36>
        <!--2013-1-10 修改动画增强效果 liqiuling end-->
			  <演:动画文本_6B3A>
        <!--<演:动画文本 uof:locID="p0071" uof:attrList="发送 间隔 动画形状 相反顺序 组合文本">-->
          <!--2010-10-25罗文甜：变量定义提前-->
          <xsl:variable name="sId" select="p:cTn[@presetID]/descendant::p:spTgt[@spid]/@spid"/>
          <xsl:variable name="grpId" select="p:cTn[@presetID]/@grpId"/>
          <xsl:variable name="bldP" select="ancestor::p:timing/p:bldLst/p:bldP[@spid=$sId and @grpId=$grpId]"/>
          <!--2011-4-12 luowentian-->
          <xsl:attribute name="发送_6B3B">
            <xsl:choose>
              <xsl:when test="p:cTn[@presetID]/p:iterate/@type='lt'">by-letter</xsl:when>
              <xsl:when test="p:cTn[@presetID]/p:iterate/@type='wd'">by-word</xsl:when>
              <xsl:otherwise>all-at-once</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!--2011-4-12 luowentian add duration by WPS-->
          <xsl:if test="p:cTn[@presetID]/p:iterate/p:tmPct">
            <xsl:variable name="dur" select="(substring-before(p:cTn[@presetID]/p:iterate/p:tmPct/@val,'%')*1000) div 100000 * 2"/>
            <xsl:attribute name="间隔_6B3C">
              <xsl:value-of select="concat('PT',$dur,'S')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:attribute name="是否使用动画形状_6B3D">
            <xsl:choose>
              <xsl:when test="descendant::p:charRg[@st='4294967295']">false</xsl:when>
              <xsl:otherwise>true</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="是否使用相反顺序_6B3E">
            <!--2010-10-25罗文甜：变量定义提前xsl:variable name="spid" select="p:cTn[@presetID]/descendant::p:spTgt[@spid]/@spid"/>
            <xsl:variable name="grpId" select="p:cTn[@presetID]/@grpId"/>
            <xsl:variable name="bldP" select="ancestor::p:timing/p:bldLst/p:bldP[@spid=$spid and @grpId=$grpId]"/-->
            <xsl:choose>
              <xsl:when test="$bldP/@rev='1'">true</xsl:when>
              <xsl:otherwise>false</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!--2010-10-25罗文甜：按金山来实现组合文本-->
          <xsl:attribute name="组合文本_6B3F">
            <xsl:choose>
              <xsl:when test="$bldP/@build='allAtOnce'">All-paragraphs-at-once</xsl:when>
              <xsl:when test="($bldP/@build='p' and not($bldP/@bldLvl))or($bldP/@build='p' and $bldP/@bldLvl='1')">by-1st-paragraph</xsl:when>
              <xsl:when test="$bldP/@build='p' and $bldP/@bldLvl='2'">by-2nd-paragraph</xsl:when>
              <xsl:when test="$bldP/@build='p' and $bldP/@bldLvl='3'">by-3rd-paragraph</xsl:when>
              <xsl:when test="$bldP/@build='p' and $bldP/@bldLvl='4'">by-4th-paragraph</xsl:when>
              <xsl:when test="$bldP/@build='p' and $bldP/@bldLvl='5'">by-5th-paragraph</xsl:when>
              <xsl:when test="$bldP/@build='Whole'or not($bldP/@build)">as-one-object</xsl:when>
              <!--xsl:otherwise>As one object</xsl:otherwise-->
            </xsl:choose>
          </xsl:attribute>
			  </演:动画文本_6B3A>
        <!-- 2010.03.29 myx add 动画声音 -->
        <xsl:if test="p:cTn/p:subTnLst[p:audio]">
			<演:声音_6B22>
          <!--<演:声音 uof:locID="p0061" uof:attrList="预定义声音 自定义声音 是否循环播放">-->
				<!--预定义声音_C631,是否循环播放_C633-->
            <xsl:attribute name="自定义声音_C632">
              <xsl:variable name="tgt">
                <xsl:value-of select="p:cTn/p:subTnLst[p:audio]/p:audio/p:cMediaNode/p:tgtEl/p:sndTgt/@r:embed"/>
              </xsl:variable>
              <!--2010-12-29 罗文甜：修改ancestor::p:sld；增加母版ID的判断-->
              <xsl:variable name="sldid">
                <xsl:choose>
                  <xsl:when test="ancestor::p:sld">
                    <xsl:value-of select="ancestor::p:sld/@id"/>
                  </xsl:when>
                  <xsl:when test="ancestor::p:sldMaster">
                    <xsl:value-of select="ancestor::p:sldMaster/@id"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:variable>
              <xsl:value-of select="concat(substring-before($sldid,'.xml'), $tgt)"/>
            </xsl:attribute>
			</演:声音_6B22>
        </xsl:if>
        <!--2013-1-11 修改增强效果的声音效果  liqiuling start-->
        <xsl:if test="p:cTn/p:subTnLst/p:cmd[@cmd='onstopaudio']">
			<演:声音_6B22>
          <!--<演:声音 uof:locID="p0061" uof:attrList="预定义声音 自定义声音 是否循环播放">-->
       
            <xsl:attribute name="预定义声音_C631">stop-previous-sound</xsl:attribute>
          </演:声音_6B22>
        </xsl:if>
        <!--2013-1-11 修改增强效果的声音效果  liqiuling end-->
      </演:增强_6B35>
	<演:效果_6B40>
        <xsl:choose>
          <xsl:when test="descendant::p:cTn[@presetID]/@presetClass='entr'">
			  <演:进入_6B41>
              <xsl:call-template name="entr"/>
				  <!--11.11.10 新的动画进入 强调 退出 有基本型温和型华丽型和细微型 李娟-->
            </演:进入_6B41>
          </xsl:when>
			
         <xsl:when test="descendant::p:cTn[@presetID]/@presetClass='emph'">
			  <演:强调_6B07>
              <xsl:call-template name="emph"/>
            </演:强调_6B07>
          </xsl:when>
          <xsl:when test="descendant::p:cTn[@presetID]/@presetClass='exit'">
			  <演:退出_6BBE>
              <xsl:call-template name="exit"/>
            </演:退出_6BBE>
          </xsl:when>
          <xsl:when test="descendant::p:cTn[@presetID]/@presetClass='path'">
			  <演:动作路径_6BD1>
          
          <!--2013-12-24, tangjiang, 修改Strict OOXML到动画动作路径的转换 start -->
				  <xsl:attribute name="路径_6BD6">
            <xsl:variable name="pathValue">
              <xsl:value-of select="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@path"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="not(contains($pathValue,'C')) and descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@pathEditMode='relative'">
                <xsl:choose>
                  <xsl:when test="contains($pathValue,'M') and contains($pathValue,'L')">
                    <xsl:variable name="pathValueM">
                      <xsl:value-of  select="substring-before($pathValue,' L ')"/>
                    </xsl:variable>
                    <xsl:variable name="pathValueL">
                      <xsl:choose>
                        <xsl:when test="contains($pathValue,' E')">
                          <xsl:value-of select="substring-before(substring-after($pathValue,' L '),' E')"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="substring-after($pathValue,' L ')"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:variable>
                    <xsl:variable name="pathValueLX">
                      <!-- 指数处理,这里均按‘E-6’和‘E+6’处理，有待完善-->
                      <xsl:variable name="pathValueLXStr" select="substring-before($pathValueL,' ')"/>
                      <xsl:choose>
                        <xsl:when test="contains($pathValueLXStr,'E+')">
                          <xsl:value-of select="round(number(substring-before($pathValueLXStr,'E+') * 1000000) *1000 )"/>
                        </xsl:when>
                        <xsl:when test="contains($pathValueLXStr,'E-')">
                          <xsl:value-of select="round(number(substring-before($pathValueLXStr,'E-') div 1000000) *1000 )"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="round(number(substring-before($pathValueL,' ')) *1000 )"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:variable>
                    <xsl:variable name="pathValueLY">
                      <xsl:variable name="pathValueLYStr" select="substring-after($pathValueL,' ')"/>
                      <xsl:choose>
                        <xsl:when test="contains($pathValueLYStr,'E+')">
                          <xsl:value-of select="round(number(substring-before($pathValueLYStr,'E+') * 1000000) *1000 )"/>
                        </xsl:when>
                        <xsl:when test="contains($pathValueLYStr,'E-')">
                          <xsl:value-of select="round(number(substring-before($pathValueLYStr,'E-') div 1000000) *1000 )"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="round(number(substring-after($pathValueL,' ')) *1000 )"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:variable>
                    <xsl:value-of select="concat($pathValueM,  ' L ', $pathValueLX,' ', $pathValueLY,' ')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$pathValue"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <!--end 2013-12-24, tangjiang, 修改Strict OOXML到动画动作路径的转换-->
              <xsl:otherwise>
                <!-- start 2015.05.19 guoyongbin,修改bug#3658 动画效果转换错误-->
                <!--<xsl:value-of select="$pathValue"/>-->
                <xsl:variable name="pathValueL">
                  <xsl:choose>
                    <xsl:when test="contains($pathValue,' E')">
                      <xsl:value-of select="substring-before(substring-after($pathValue,' L '),' E')"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:variable>
                <xsl:variable name="pathValueLX">
                  <xsl:variable name="pathValueLXStr" select="substring-before($pathValueL,' ')"/>
                  <xsl:choose>
                    <xsl:when test="contains($pathValueLXStr,'E+')">
                      <xsl:value-of select="round(number(substring-before($pathValueLXStr,'E+') * 1000000) *1000 )"/>
                    </xsl:when>
                    <xsl:when test="contains($pathValueLXStr,'E-')">
                      <xsl:value-of select="round(number(substring-before($pathValueLXStr,'E-') div 1000000) *1000 )"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(number(substring-before($pathValueL,' ')) *800 )"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:variable name="pathValueLY">
                  <xsl:variable name="pathValueLYStr" select="substring-after($pathValueL,' ')"/>
                  <xsl:choose>
                    <xsl:when test="contains($pathValueLYStr,'E+')">
                      <xsl:value-of select="round(number(substring-before($pathValueLYStr,'E+') * 1000000) *1000 )"/>
                    </xsl:when>
                    <xsl:when test="contains($pathValueLYStr,'E-')">
                      <xsl:value-of select="round(number(substring-before($pathValueLYStr,'E-') div 1000000) *1000 )"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(number(substring-after($pathValueL,' ')) *1000 )"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <!-- <xsl:value-of select="round(number(substring-after($pathValueL,' ')) *1000 )"/>-->
                <xsl:value-of select="concat(substring-before($pathValue,' L '),  ' L ',substring-before($pathValueL,' '),' ', $pathValueLY,' ')"/>
                <!-- end 2015.05.19 guoyongbin,修改bug#3658 动画效果转换错误-->
              </xsl:otherwise>
            </xsl:choose>
				  </xsl:attribute>
         

          <xsl:call-template name="effectdur"/>
				  <xsl:choose>
						<xsl:when test="descendant::p:cTn[@presetID]/@accel">
							<xsl:attribute name="是否平稳开始_6BD8">true</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="是否平稳开始_6BD8">false</xsl:attribute>
					</xsl:otherwise>
					</xsl:choose> 
				  <xsl:choose>
					 <xsl:when test="descendant::p:cTn[@presetID]/@decel">
						<xsl:attribute name="是否平稳结束_6BD9">true</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="是否平稳结束_6BD9">false</xsl:attribute>
					</xsl:otherwise>
				 </xsl:choose>
				  <xsl:choose>
				<xsl:when test="descendant::p:cTn[@presetID]/@autoRev='1'">
					 <xsl:attribute name="是否自动翻转_6BDA">true</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
						<xsl:attribute name="是否自动翻转_6BDA">false</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose> 
				  <xsl:choose>
						<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@pathEditMode='fixed'">
							<xsl:attribute name="是否锁定_6BD7">true</xsl:attribute>
						</xsl:when>
			<xsl:otherwise>
			  <xsl:attribute name="是否锁定_6BD7">false</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<xsl:choose>
			<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/p:cBhvr/p:cTn/@spd='-100%'">
				<xsl:attribute name="是否反转路径方向_6BDC">true</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="是否反转路径方向_6BDC">false</xsl:attribute>
			</xsl:otherwise>
        </xsl:choose>
          <!--2013-1-11 动作路径直线消失 liqiuling 其他路径需要补充 start-->
          <演:形状_6BD2>
           <xsl:call-template name="pathshape"/>
          </演:形状_6BD2>
				  
			  </演:动作路径_6BD1>
				  </xsl:when>
	  </xsl:choose>
  </演:效果_6B40>
	  </演:序列_6B1B>
  </xsl:template>


  <xsl:template name="pathshape">
    <xsl:choose>
      <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@path = 'M 0 0 L 0 -0.25 E'">
        <演:直线和曲线_6BD4>up</演:直线和曲线_6BD4>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@path = 'M 0 0 L 0 0.25 E'">
        <演:直线和曲线_6BD4>down</演:直线和曲线_6BD4>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@path = 'M 0 0 L -0.25 0 E'">
        <演:直线和曲线_6BD4>left</演:直线和曲线_6BD4>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animMotion/@path = 'M 0 0 L 0.25 0 E'">
        <演:直线和曲线_6BD4>right</演:直线和曲线_6BD4>
      </xsl:when>
      
    </xsl:choose>
    
  </xsl:template>
  <!--2013-1-11 动作路径直线消失 liqiuling 其他路径需要补充 end-->
  
	<xsl:template name="entr">
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='22' or descendant::p:cTn[@presetID]/@presetID='1'or descendant::p:cTn[@presetID]/@presetID='2' or descendant::p:cTn[@presetID]/@presetID='4' or descendant::p:cTn[@presetID]/@presetID='7' or descendant::p:cTn[@presetID]/@presetID='18' or descendant::p:cTn[@presetID]/@presetID='8'
				or descendant::p:cTn[@presetID]/@presetID='16' or descendant::p:cTn[@presetID]/@presetID='5' or descendant::p:cTn[@presetID]/@presetID='12'  or descendant::p:cTn[@presetID]/@presetID='13' or descendant::p:cTn[@presetID]/@presetID='14' or descendant::p:cTn[@presetID]/@presetID='9'
				or descendant::p:cTn[@presetID]/@presetID='21' or descendant::p:cTn[@presetID]/@presetID='6' or descendant::p:cTn[@presetID]/@presetID='3'">
		<演:基本型_6B42>
			<xsl:choose>
        <!--2013-1-11 修复擦除方向不正确  liqiuling start-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='22'">
					<演:擦除_6B57>
						<xsl:attribute name="方向_6B58">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">from-bottom</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">from-left</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">from-right</xsl:when>
                <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='1'">from-top</xsl:when>
								<xsl:otherwise>from-left</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
						<xsl:call-template name="effectdur"/>
					</演:擦除_6B57>
				</xsl:when>
        <!--2013-1-11 修复擦除方向不正确  liqiuling end-->
				<!--2011-01-03罗文甜 淡出10 转换为 出现-->
				<!--去除淡出 效果 李娟 11.12.01-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='1'">
					<演:出现_6B46/>
				</xsl:when>
				<!--2011-01-03罗文甜 上浮42  下浮47  升起37  字幕式28 曲线向上52 空翻56 浮动30 下拉38 螺旋飞入15 转换为 飞入-->
				<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='2' or descendant::p:cTn[@presetID]/@presetID='42' or descendant::p:cTn[@presetID]/@presetID='47' or descendant::p:cTn[@presetID]/@presetID='37'
                or descendant::p:cTn[@presetID]/@presetID='28' or descendant::p:cTn[@presetID]/@presetID='52' or descendant::p:cTn[@presetID]/@presetID='56' or descendant::p:cTn[@presetID]/@presetID='30'
                or descendant::p:cTn[@presetID]/@presetID='38' or descendant::p:cTn[@presetID]/@presetID='15'">-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='2'">
					<演:飞入_6B59>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="方向_6B5A">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">from-bottom</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">from-left</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">from-right</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='1'">from-top</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='12'">from-bottom-left</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='6'">from-bottom-right</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='9'">from-top-left</xsl:when>
								<xsl:otherwise>from-top-right</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:飞入_6B59>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='4'">
					<演:盒状_6B47>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
								<xsl:attribute name="方向_6B48">out</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B48">in</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:盒状_6B47>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='7'">
					<演:缓慢进入_6B5B>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="方向_6B58">
							<xsl:choose>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='4'">from-bottom</xsl:when>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='8'">from-left</xsl:when>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='2'">from-right</xsl:when>
								<xsl:otherwise>from-top</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:缓慢进入_6B5B>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='18'">
					<演:阶梯状_6B49>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="方向_6B4A">
							<xsl:choose>
								<!--luowentian -->
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='12'">left-down</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='9'">left-up</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='6'">right-down</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='3'">right-up</xsl:when>
								<xsl:otherwise>left-down</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:阶梯状_6B49>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='8'">
					<演:菱形_6B5D>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
								<xsl:attribute name="方向_6B48">out</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B48">in</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:菱形_6B5D>
				</xsl:when>
				<!--2010-10-25罗文甜：修改辐射状属性为轮辐-->
				<!--2011-01-03罗文甜 旋转45 回旋49 飞旋25 转换为 轮子-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='21'">
					<演:轮子_6B4B>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="轮辐_6B4D">
							<xsl:value-of select="descendant::p:cTn[@presetID]/@presetSubtype"/>
						</xsl:attribute>
					</演:轮子_6B4B>

				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='16'">
					<演:劈裂_6B5E>
						<!--<演:劈裂 uof:locID="p0094" uof:attrList="速度 方向">-->
						<xsl:call-template name="effectdur"/>
						<!--2011-4-13 luowentian-->
						<xsl:attribute name="方向_6B5F">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='26'">horizontal-in</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='21'">horizontal-out</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='42'">vertical-in</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='37'">vertical-out</xsl:when>
							</xsl:choose>
						</xsl:attribute>
						<!--</演:劈裂>-->
					</演:劈裂_6B5E>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='5'">
					<演:棋盘_6B4E>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='10'">
								<xsl:attribute name="方向_6B50">across</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B50">down</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:棋盘_6B4E>
				</xsl:when>
				<!--2011-01-03罗文甜 缩放53 基本缩放23 转换为 切入-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='12'">
					<演:切入_6B60>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="方向_6B58">
							<xsl:choose>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='4'">from-bottom</xsl:when>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='8'">from-left</xsl:when>
								<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='2'">from-right</xsl:when>
								<xsl:otherwise>from-top</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:切入_6B60>
				</xsl:when>
				<!-- 注销闪烁一次 李娟 11.12.01-->
				<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='11'">
					<演:闪烁一次_6B51>
						<xsl:call-template name="effectdur"/>
					</演:闪烁一次_6B51>
				</xsl:when>-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='13'">
					<演:十字形扩展_6B53>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
								<xsl:attribute name="方向_6B48">out</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B48">in</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:十字形扩展_6B53>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='14'">
					<演:随机线条_6B62>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='5'">
								<xsl:attribute name="方向_6B45">vertical</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B45">horizontal</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:随机线条_6B62>
				</xsl:when>
				<!--xsl:when test="descendant::p:cTn[@presetID]/"@presetID='24'"-->
				<!--2011-01-03罗文甜 翻转式由远及近31 弹跳26 挥鞭式41 转换为 -->
				<!--修改随机效果 李娟11.12。01-->
				<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='26' or descendant::p:cTn[@presetID]/@presetID='41'">
					<演:随机效果_6B55/>
				</xsl:when>-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='9'">
					<演:向内溶解_6B64>
						<xsl:call-template name="effectdur"/>
					</演:向内溶解_6B64>
				</xsl:when>
				<!-- 注销 此部分 李娟 11.12.01-->
				<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='20'">
					<演:扇形展开_6B61>
						<xsl:call-template name="effectdur"/>
					</演:扇形展开_6B61>
				</xsl:when>-->
				<!--2011-01-03罗文甜 中心旋转43 玩具风车35 基本旋转19 转换为 圆形扩展-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='6'">
					<演:圆形扩展_6B56>
						<xsl:call-template name="effectdur"/>
						<!--luowentian-->
						<xsl:choose>
							<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='32'">
								<xsl:attribute name="方向_6B48">out</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B48">in</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</演:圆形扩展_6B56>
				</xsl:when>
				<!--2011-01-03罗文甜 展开55 转换为 百叶窗-->
				<!--<xsl:otherwise>-->
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='3'">
					<演:百叶窗_6B43>
						<!--luowentian-->
						<xsl:call-template name="effectdur"/>

            <!--2014-03-09, pengxin, 修改Strict OOXML到UOF母版-动画效果改变 start -->
						<xsl:choose>
							<xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='10'">
								<xsl:attribute name="方向_6B45">horizontal</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="方向_6B45">vertical</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
            <!-- end 2014-03-09, pengxin -->
            
					</演:百叶窗_6B43>
					<!--</xsl:otherwise>-->
				</xsl:when>
			</xsl:choose>
		</演:基本型_6B42>
		</xsl:if>
		<!--添加细微型 李娟11.12.01-->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='55' or descendant::p:cTn[@presetID]/@presetID='10'or descendant::p:cTn[@presetID]/@presetID='45' or descendant::p:cTn[@presetID]/@presetID='53'">
			<演:细微型_6B66>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='55'">
						<演:展开_6B6B>
							<xsl:call-template name="effectdur"/>
						</演:展开_6B6B>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='10'">
						<演:渐变_6B67>
							<xsl:call-template name="effectdur"/>
						</演:渐变_6B67>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='45'">
						<演:渐变式回旋_6B69>
							<xsl:call-template name="effectdur"/>
						</演:渐变式回旋_6B69>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='53'">
						<演:渐变式缩放_6B6A>
							<xsl:call-template name="effectdur"/>
						</演:渐变式缩放_6B6A>
					</xsl:when>
				</xsl:choose>
			</演:细微型_6B66>
		</xsl:if>
		<!--李娟 温和型翻转式31，上浮42，下浮47，基本缩放23，升起37，回旋49，中心旋转43-->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='31' or descendant::p:cTn[@presetID]/@presetID='42' or descendant::p:cTn[@presetID]/@presetID='47' or descendant::p:cTn[@presetID]/@presetID='37'
				 or descendant::p:cTn[@presetID]/@presetID='23'">
			<演:温和型_6B6D>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='31'">
						<演:翻转时由远及近_6B6E>
							<xsl:call-template name="effectdur"/>
						</演:翻转时由远及近_6B6E>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='23'">
						<演:缩放_6B72>
							<xsl:call-template name="effectdur"/>
						</演:缩放_6B72>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='42'">
						<演:上升_6B76>
							<xsl:call-template name="effectdur"/>
						</演:上升_6B76>
				</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='47'">
						<演:下降_6B78>
							<xsl:call-template name="effectdur"/>
						</演:下降_6B78>
			   </xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='37'">
						<演:升起_6B77>
							<xsl:call-template name="effectdur"/>
						</演:升起_6B77>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='43'">
						<演:中心旋转_6B7B>
							<xsl:call-template name="effectdur"/>
						</演:中心旋转_6B7B>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='49'">
						<演:回旋_6B75>
							<xsl:call-template name="effectdur"/>
						</演:回旋_6B75>
					</xsl:when>
				</xsl:choose>
			</演:温和型_6B6D>
		</xsl:if>
		<!--李娟 添加动画华丽型 弹跳26，挥鞭式41，浮动30，飞旋25，基本旋转19转为旋转，下拉38转为挥舞，空翻56，字幕式28，螺旋飞入15，曲线向上52  -->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='26' or descendant::p:cTn[@presetID]/@presetID='41' or descendant::p:cTn[@presetID]/@presetID='30'or descendant::p:cTn[@presetID]/@presetID='25'or descendant::p:cTn[@presetID]/@presetID='19'or descendant::p:cTn[@presetID]/@presetID='56' or descendant::p:cTn[@presetID]/@presetID='28' or descendant::p:cTn[@presetID]/@presetID='38' or descendant::p:cTn[@presetID]/@presetID='35'">
			<演:华丽型_6B7C>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='26'">
						<演:弹跳_6B7D>
							<xsl:call-template name="effectdur"/>
						</演:弹跳_6B7D>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='41'">
						<演:挥鞭式_6B80>
							<xsl:call-template name="effectdur"/>
						</演:挥鞭式_6B80>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='30'">
						<演:浮动_6B8A>
							<xsl:call-template name="effectdur"/>
						</演:浮动_6B8A>
				</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='25'">
						<演:飞旋_6B7E>
							<xsl:call-template name="effectdur"/>
						</演:飞旋_6B7E>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='38'">
						<演:旋转_6B84>
							<xsl:call-template name="effectdur"/>
						</演:旋转_6B84>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='19'">
						<演:挥舞_6B8C>
							<xsl:call-template name="effectdur"/>
						</演:挥舞_6B8C>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='56'">
						<演:空翻_6B81>
							<xsl:call-template name="effectdur"/>
						</演:空翻_6B81>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='28'">
						<演:字幕式_6B87>
							<xsl:call-template name="effectdur"/>
						</演:字幕式_6B87>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='15'">
						<演:螺旋飞入_6B8D>
							<xsl:call-template name="effectdur"/>
						</演:螺旋飞入_6B8D>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='52'">
						<演:曲线向上_6B82>
							<xsl:call-template name="effectdur"/>
						</演:曲线向上_6B82>
					</xsl:when>
				</xsl:choose>
			</演:华丽型_6B7C>
		</xsl:if>
	</xsl:template>

	<xsl:template name= "emph">
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='1' or descendant::p:cTn[@presetID]/@presetID='7' or descendant::p:cTn[@presetID]/@presetID='3' or descendant::p:cTn[@presetID]/@presetID='9' or descendant::p:cTn[@presetID]/@presetID='8' or descendant::p:cTn[@presetID]/@presetID='6'">
			<演:基本型_6C19>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='1'">
						<演:更改填充颜色_6B9F>
							<xsl:call-template name="effectdur"/>
							<xsl:attribute name="颜色_6B95">
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
										<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
										<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
									</xsl:when>
									<xsl:otherwise>#ffffff</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:更改填充颜色_6B9F>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='7'">
						<演:更改线条颜色_6B94>
							<xsl:call-template name="effectdur"/>
							<xsl:attribute name="颜色_6B95">
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
										<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
										<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
									</xsl:when>
									<xsl:otherwise>#ffffff</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:更改线条颜色_6B94>
					</xsl:when>
					<!--xsl:when test="descendant::p:cTn[@presetID]/@presetID='4'">
        <演:更改字号 uof:locID="p0125" uof:attrList="速度 预定义尺寸 自定义尺寸" 演:预定义尺寸="huge">
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="演:预定义尺寸">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:anim/@to='0.25'">tiny</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:anim/@to='1.5'">larger</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:anim/@to='4'">huge</xsl:when>
              <xsl:otherwise>smaller</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:更改字号>
      </xsl:when-->
					<!--2011-01-03 罗文甜：画笔颜色16 补色21 补色2（22） 对比色23 对象颜色19 加深24 彩色延伸28 转换为-->
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='3'">
						<演:更改字体颜色_6BA2>
							<xsl:call-template name="effectdur"/>
							<xsl:attribute name="颜色_6B95">
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
										<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
										<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr"/>
									</xsl:when>
									<xsl:otherwise>#ffffff</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:更改字体颜色_6BA2>
					</xsl:when>
					<!--xsl:when test="descendant::p:cTn[@presetID]/"@presetID='5'"-->
					<!--2011-01-03 罗文甜：加粗闪烁10 转换为-->
					<!--李娟 11.12.01-->
					<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='10'">
		  <演:更改字形_6B96>
			  <xsl:attribute name="期间_6B98">2.0</xsl:attribute>
			  -->
					<!--<xsl:attribute name="字形_6B97"></xsl:attribute>-->
					<!--
		  </演:更改字形_6B96>
      </xsl:when>-->

					<!--2011-01-03 罗文甜：下划线18 转换为 更改字体，待修改>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='18'">
        <演:更改字体/>
      </xsl:when-->
					<!--2011-01-03 罗文甜：变淡30 彩色脉冲27 闪烁35 转换为-->
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='9'">
						<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='9' or descendant::p:cTn[@presetID]/@presetID='30' or descendant::p:cTn[@presetID]/@presetID='27'
                or descendant::p:cTn[@presetID]/@presetID='35'">-->
						<演:透明_6BA3>
							<xsl:attribute name="预定义透明度_6BA4">
								<!--2011-4-13 luowentian-->
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:set/p:to/p:strVal/@val='0'">
										<xsl:value-of select="'100'"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:set/p:to/p:strVal/@val='0.75'">
										<xsl:value-of select="'25'"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:set/p:to/p:strVal/@val='0.25'">
										<xsl:value-of select="'75'"/>
									</xsl:when>
									<xsl:otherwise>50</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<!--<xsl:attribute name="自定义透明度_6BA5">有待扩充 李娟 11.11.10</xsl:attribute>-->
							<xsl:attribute name="期间_6BA6">
								<xsl:variable name="dur" select="descendant::p:cTn[@presetID]/p:childTnLst/p:par/descendant::p:cTn[@presetID]/p:childTnLst/p:par/descendant::p:cTn[@presetID]/p:childTnLst/p:set/p:cBhvr/descendant::p:cTn[@presetID]/@durn/p:childTnLst/p:set/p:cBhvr/descendant::p:cTn[@presetID]/@dur"/>
								<xsl:choose>
									<xsl:when test="string($dur div 1000) != 'NaN'">
										<xsl:value-of select="$dur div 1000"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:choose>
											<xsl:when test="descendant::p:cTn[@presetID]/p:endCondLst">until-next-click</xsl:when>
											<!--永中 bug-->
											<xsl:otherwise>until-next-click</xsl:otherwise>
										</xsl:choose>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:透明_6BA3>
					</xsl:when>
					<!--2011-01-03 罗文甜：跷跷板32 波浪形34 转换为-->
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='8'">
						<!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='8' or descendant::p:cTn[@presetID]/@presetID='32' or descendant::p:cTn[@presetID]/@presetID='34'">-->
						<演:陀螺旋_6B9B>
							<xsl:call-template name="effectdur"/>
							<!--luowentian-->
							<xsl:choose>
								<xsl:when test="contains(descendant::p:cTn[@presetID]/p:childTnLst/p:animRot/@by,'-')">
									<xsl:attribute name="是否顺时针方向_6B9C">false</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="是否顺时针方向_6B9C">true</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<xsl:variable name="angle">
								<xsl:choose>
									<xsl:when test="contains(descendant::p:cTn[@presetID]/p:childTnLst/p:animRot/@by,'-')">
										<xsl:value-of select="substring-after(descendant::p:cTn[@presetID]/p:childTnLst/p:animRot/@by,'-')"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="descendant::p:cTn[@presetID]/p:childTnLst/p:animRot/@by"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:variable>
							<!--<xsl:attribute name="自定义角度_6B9E"> 待添加 李娟 11.11.10</xsl:attribute>-->
							<xsl:attribute name="预定义角度_6B9D">
								<xsl:choose>
									<xsl:when test="$angle='5400000'">quarter-spin</xsl:when>
									<xsl:when test="$angle='10800000'">half-spin</xsl:when>
									<xsl:when test="$angle='43200000'">two-spins</xsl:when>
									<xsl:otherwise>full-spin</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:陀螺旋_6B9B>
					</xsl:when>
					<!--2011-01-03 罗文甜：脉冲26 不饱和25 闪现36 加粗显示15 下划线18 转换为-->
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='6'">
						<演:缩放_6B72>
							<xsl:call-template name="effectdur"/>
              <xsl:variable name="tox" select="descendant::p:cTn[@presetID]/p:childTnLst/p:animScale/p:by/@x"/>
              <xsl:variable name="toy" select="descendant::p:cTn[@presetID]/p:childTnLst/p:animScale/p:by/@y"/>
							<!--2011-4-13 luowentian-->
              <xsl:attribute name="方向_6B91">
              <xsl:choose>
                <xsl:when test="$tox='25000' or $toy='25000'or $tox='50000' or $toy='50000'">
                  
                    <xsl:choose>
                      <xsl:when test="$toy &gt;$tox">
                        <xsl:value-of select="'horizontal'"/>
                      </xsl:when>
                      <xsl:when test="$tox &gt;$toy">
                        <xsl:value-of select="'vertical'"/>
                      </xsl:when>
                      <xsl:when test="$toy = $tox">
                        <xsl:value-of select="'both'"/>
                      </xsl:when>
                    </xsl:choose>
                </xsl:when>
                <xsl:when test="$tox='150000' or $toy='150000'or $tox='400000' or $toy='400000'">
                    <xsl:choose>
                      <xsl:when test="$toy &gt;$tox">
                        <xsl:value-of select="'vertical'"/>
                      </xsl:when>
                      <xsl:when test="$tox &gt;$toy">
                        <xsl:value-of select="'horizontal'"/>
                      </xsl:when>
                      <xsl:when test="$toy = $tox">
                        <xsl:value-of select="'both'"/>
                      </xsl:when>
                    </xsl:choose>
                </xsl:when>
              </xsl:choose>
              </xsl:attribute>
                  
							<!--<xsl:attribute name="方向_6B91">
								<xsl:choose>
									<xsl:when test="$toy &gt;$tox)">
										<xsl:value-of select="'vertical'"/>
									</xsl:when>
									<xsl:when test="$tox &gt;$toy">
										<xsl:value-of select="'horizontal'"/>
									</xsl:when>
                  <xsl:when test="$toy &lt;$tox">
                    <xsl:value-of select="'vertical'"/>
                  </xsl:when>
                  <xsl:when test="$tox &lt;$toy">
                    <xsl:value-of select="'horizontal'"/>
                  </xsl:when>
                  
									<xsl:when test="$toy = $tox">
										<xsl:value-of select="'both'"/>
									</xsl:when>
								</xsl:choose>
							</xsl:attribute>-->
							<!--<xsl:attribute name="自定义角度_6B9E"> 待添加 李娟 11.11.10</xsl:attribute>-->
                 <xsl:attribute name="预定义尺寸_6B92">
								<xsl:choose>
									<xsl:when test="$tox='25000' or $toy='25000'">tiny</xsl:when>
									<xsl:when test="$tox='150000' or $toy='150000'">larger</xsl:when>
									<xsl:when test="$tox='400000' or $toy='400000'">huge</xsl:when>
									<xsl:otherwise>smaller</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:缩放_6B72>
					</xsl:when>
				</xsl:choose>
			</演:基本型_6C19>
		</xsl:if>
		<!-- 李娟 画笔颜色16  色21 补色2（22） 对比色23 对象颜色19 加深24   下划线18  不饱和25 变淡30  脉冲26 加粗闪烁 -->
		<!--李娟 画笔颜色 转为彩色波纹 对象颜色-垂直突出显示，脉冲-忽明忽暗-->
    <!--2013-1-11 修改加深变淡效果 以及对象颜色转为垂直突出显示 liqiuling start 修改if控制条件-->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='19' or descendant::p:cTn[@presetID]/@presetID='30' or descendant::p:cTn[@presetID]/@presetID='18' or descendant::p:cTn[@presetID]/@presetID='26' or descendant::p:cTn[@presetID]/@presetID='16' or descendant::p:cTn[@presetID]/@presetID='21' or descendant::p:cTn[@presetID]/@presetID='22' or descendant::p:cTn[@presetID]/@presetID='23' or descendant::p:cTn[@presetID]/@presetID='24' or descendant::p:cTn[@presetID]/@presetID='28'or descendant::p:cTn[@presetID]/@presetID='10'">
			<演:细微型_6C20>
				<xsl:choose>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='26'">
						<演:忽明忽暗_6BB2>
							<xsl:call-template name="effectdur"/>
						</演:忽明忽暗_6BB2>
					</xsl:when>
					
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='21'">
						<演:补色_6BAE>
							<xsl:call-template name="effectdur"/>
						</演:补色_6BAE>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='22'">
						<演:补色2_6BA8>
							<xsl:call-template name="effectdur"/>
						</演:补色2_6BA8>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='24'">
						<演:加深_6BAC>
							<xsl:call-template name="effectdur"/>
						</演:加深_6BAC>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='19'"><!-- 颜色 有待添加  李娟 11.12.01              -->
						<演:垂直突出显示_6BB0>
							<xsl:call-template name="effectdur"/>
							<xsl:attribute name="颜色_6B95">
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
										<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
										<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
									</xsl:when>
									<xsl:otherwise>#ffffff</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:垂直突出显示_6BB0>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='18'">
						<演:添加下划线_6BB4>
							<xsl:call-template name="effectdur"/>
						</演:添加下划线_6BB4>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='30'">
						<演:变淡_6BA7>
							<xsl:call-template name="effectdur"/>
						</演:变淡_6BA7>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='25'">
						<演:不饱和_6BAF>
							<xsl:call-template name="effectdur"/>
						</演:不饱和_6BAF>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='26'">
						<演:忽明忽暗_6BB2>
							<xsl:call-template name="effectdur"/>
						</演:忽明忽暗_6BB2>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='16'">
						<演:彩色波纹_6BA9>
							<xsl:call-template name="effectdur"/>
							<xsl:attribute name="颜色_6B95">
								<xsl:choose>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
										<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
									</xsl:when>
									<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
										<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
									</xsl:when>
									<xsl:otherwise>#ffffff</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</演:彩色波纹_6BA9>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='23'">
						<演:对比色_6BAA>
							<xsl:call-template name="effectdur"/>
						</演:对比色_6BAA>
					</xsl:when>
					<xsl:when test="descendant::p:cTn[@presetID]/@presetID='10'">
						<演:加粗闪烁_6BB3>
							<xsl:call-template name="effectdur"/>
						</演:加粗闪烁_6BB3>
					</xsl:when>
				</xsl:choose>
			</演:细微型_6C20>
		</xsl:if>
		<!--温和型 跷跷板32  彩色脉冲  彩色延伸28  -闪现36  各元素的子元素 颜色 有待添加 李娟 11.12.01 彩色脉冲27转为闪动-->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='32' or descendant::p:cTn[@presetID]/@presetID='28' or descendant::p:cTn[@presetID]/@presetID='27' or descendant::p:cTn[@presetID]/@presetID='36'">
		<演:温和型_6C21>
			<xsl:choose>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='32'">
					<演:跷跷板_6BB7>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="颜色_6B95">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
									<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
								</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
									<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
								</xsl:when>
								<xsl:otherwise>#ffffff</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:跷跷板_6BB7>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='28'">
					<演:彩色延伸_6BB5>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="颜色_6B95">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
									<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
								</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
									<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
								</xsl:when>
								<xsl:otherwise>#ffffff</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:彩色延伸_6BB5>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='27'">
					<演:闪动_6BB6>
						<xsl:call-template name="effectdur"/>
						<xsl:attribute name="颜色_6B95">
							<xsl:choose>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr">
									<xsl:value-of select="concat('#',descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:srgbClr/@val)"/>
								</xsl:when>
								<xsl:when test="descendant::p:cTn[@presetID]/p:childTnLst/p:animClr/p:to/a:schemeClr">
									<xsl:apply-templates select="descendant::p:cTn[@presetID]/p:subTnLst/p:animClr/p:to/a:schemeClr"/>
								</xsl:when>
								<xsl:otherwise>#ffffff</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</演:闪动_6BB6>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='36'">
					<演:闪现_6BB8>
						<xsl:call-template name="effectdur"/>
					</演:闪现_6BB8>
				</xsl:when>
			</xsl:choose>
		</演:温和型_6C21>
		</xsl:if>
		<!--华丽型  波浪形34 加粗显示15 的属性2秒         闪烁35                                李娟11.12.01-->
		<xsl:if test="descendant::p:cTn[@presetID]/@presetID='34' or descendant::p:cTn[@presetID]/@presetID='15' or descendant::p:cTn[@presetID]/@presetID='35'">
			<演:华丽型_6C22>
			<xsl:choose>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='34'">
					<演:波浪形_6BBC>
						<xsl:call-template name="effectdur"/>
					</演:波浪形_6BBC>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='35'">
					<演:闪烁_6BBD>
						<xsl:call-template name="effectdur"/>
					</演:闪烁_6BBD>
				</xsl:when>
				<xsl:when test="descendant::p:cTn[@presetID]/@presetID='15'">
					<演:加粗展示_6BBA>
						<xsl:attribute name="期间_6B98">2.0</xsl:attribute>
					</演:加粗展示_6BBA>
				</xsl:when>
			</xsl:choose>
			</演:华丽型_6C22>
		</xsl:if>
	</xsl:template>
	
	<!--修改退出效果，李娟 11.12.05-->
  <xsl:template name="exit">
	  <xsl:if test="descendant::p:cTn[@presetID]/@presetID='6' or descendant::p:cTn[@presetID]/@presetID='20' or descendant::p:cTn[@presetID]/@presetID='1' 
                or descendant::p:cTn[@presetID]/@presetID='2' or descendant::p:cTn[@presetID]/@presetID='9' or descendant::p:cTn[@presetID]/@presetID='14' or descendant::p:cTn[@presetID]/@presetID='12' or descendant::p:cTn[@presetID]/@presetID='5'
			  or descendant::p:cTn[@presetID]/@presetID='16' or descendant::p:cTn[@presetID]/@presetID='8' or descendant::p:cTn[@presetID]/@presetID='21' or descendant::p:cTn[@presetID]/@presetID='18' or descendant::p:cTn[@presetID]/@presetID='4'
			  or descendant::p:cTn[@presetID]/@presetID='22' or descendant::p:cTn[@presetID]/@presetID='13' or descendant::p:cTn[@presetID]/@presetID='3'">
	  <演:基本型_6C23>
    <xsl:choose>
      <!--2011-01-03 罗文甜 中心旋转43 玩具风车35 基本旋转19 转换为-->
		
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='6'">
		  <演:圆形扩展_6B56>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
              <xsl:attribute name="方向_6B48">out</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B48">in</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:圆形扩展_6B56>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='20'">
		  <演:扇形展开_6B61>
          <xsl:call-template name="effectdur"/>
        </演:扇形展开_6B61>
      </xsl:when>
      <!--2011-01-03 罗文甜 淡出10 转换为-->
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='1'">
		  <演:消失_6BC7/>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='9'">
		  <演:向外溶解_6BC5>
          <xsl:call-template name="effectdur"/>
        </演:向外溶解_6BC5>
      </xsl:when>
      <!--xsl:when test="descendant::p:cTn[@presetID]/@presetID='24'"-->
      <!--2011-01-03 罗文甜 收缩并旋转31 弹跳26 挥鞭式41 转换为-->
		<!-- 注销这部分内容 李娟 11.12.05-->
      <!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='31' or descendant::p:cTn[@presetID]/@presetID='26' or descendant::p:cTn[@presetID]/@presetID='41'">
		  <演:随机效果_6B55/>
      </xsl:when>-->
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='14'">
		  <演:随机线条_6B62>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='5'">
              <xsl:attribute name="方向_6B45">vertical</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B45">horizontal</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:随机线条_6B62>
      </xsl:when>
      <!--xsl:when test="descendant::p:cTn[@presetID]/@presetID='11'">
        <演:闪烁一次 uof:locID="p0115" uof:attrList="速度">
          <xsl:call-template name="effectdur"/>
        </演:闪烁一次>
      </xsl:when-->
      <!--李娟 11.12。05  缩放53 基本缩放23被注销掉 不是此部分内容  转换为-->
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='12'">
		  <演:切出_6BC4>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="方向_6BC2">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">to-bottom</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">to-left</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">to-right</xsl:when>
              <xsl:otherwise>to-top</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:切出_6BC4>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='5'">
		  <演:棋盘_6B4E>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='5'">
              <xsl:attribute name="方向_6B50">down</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B50">across</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:棋盘_6B4E>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='16'">
		  <演:劈裂_6B5E>
          <xsl:call-template name="effectdur"/>
          <!--2011-4-13 luowentian-->
          <xsl:attribute name="方向_6B5F">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='26'">horizontal-in</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='21'">horizontal-out</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='42'">vertical-in</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='37'">vertical-out</xsl:when>
              
            </xsl:choose>
          </xsl:attribute>
        </演:劈裂_6B5E>
      </xsl:when>
      <!--2010-10-25罗文甜：修改辐射状属性为轮辐-->
      <!--   不再基本型 效果 暂时注销 李娟 11.12.05 转换为-->
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='21'">
		  <演:轮子_6B4B>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="轮辐_6B4D">
            <xsl:value-of select="descendant::p:cTn[@presetID]/@presetSubtype"/>
          </xsl:attribute>
        </演:轮子_6B4B>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='8'">
		  <演:菱形_6B5D>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
              <xsl:attribute name="方向_6B48">out</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B48">in</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:菱形_6B5D>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='18'">
		  <演:阶梯状_6B49>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="方向_6B4A">
            <xsl:choose>
              <!--luowentian -->
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='12'">left-down</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='9'">left-up</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='6'">right-down</xsl:when>
              <xsl:otherwise>right-up</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:阶梯状_6B49>
      </xsl:when>
      <!--<xsl:when test="descendant::p:cTn[@presetID]/@presetID='7'">
		  <演:缓慢移出_6BC1>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="方向_6BC2">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">to-bottom</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">to-left</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">to-right</xsl:when>
              <xsl:otherwise>to top</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:缓慢移出_6BC1>
      </xsl:when>-->
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='4'">
		  <演:盒状_6B47>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='32'">
              <xsl:attribute name="方向_6B48">out</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B48">in</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:盒状_6B47>
      </xsl:when>
      
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='2'">
		  <演:飞出_6BBF>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="方向_6BC0">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">to-bottom</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">to-left</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">to-right</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='1'">to-top</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='12'">to-bottom-left</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='6'">to-bottom-right</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='9'">to-top-left</xsl:when>
              <xsl:otherwise>to-top-right</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:飞出_6BBF>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='22'">
		  <演:擦除_6B57>
          <xsl:call-template name="effectdur"/>
          <xsl:attribute name="方向_6B58">
            <xsl:choose>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='4'">from-bottom</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='8'">from-left</xsl:when>
              <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='2'">from-right</xsl:when>
              <xsl:otherwise>from-top</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </演:擦除_6B57>
      </xsl:when>
      <xsl:when test="descendant::p:cTn[@presetID]/@presetID='13'">
		  <演:十字形扩展_6B53>
          <xsl:call-template name="effectdur"/>
          <!--luowentian-->
          <xsl:choose>
            <xsl:when test ="descendant::p:cTn[@presetID]/@presetSubtype='32'">
              <xsl:attribute name="方向_6B48">out</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B48">in</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:十字形扩展_6B53>
      </xsl:when>
      <!--2011-01-03 罗文甜 收缩55 转换为-->
		<xsl:when test="descendant::p:cTn[@presetID]/@presetID='3'">
		  <演:百叶窗_6B43>
          <xsl:call-template name="effectdur"/>
         <!--luowentian-->
          <xsl:choose>
            <xsl:when test="descendant::p:cTn[@presetID]/@presetSubtype='5'">
              <xsl:attribute name="方向_6B45">vertical</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="方向_6B45">horizontal</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </演:百叶窗_6B43>
      </xsl:when>
    </xsl:choose>
	  </演:基本型_6C23>
	  </xsl:if>
	  <!--添加细微型 李娟 11.12。05
	   -->
	  <!--  李娟11.12.05 淡出10，缩放53  收缩55 旋转45-->
	  <xsl:if test="descendant::p:cTn[@presetID]/@presetID='10' or descendant::p:cTn[@presetID]/@presetID='53' or descendant::p:cTn[@presetID]/@presetID='55' or descendant::p:cTn[@presetID]/@presetID='45'">
		  <演:细微型_6C24>
			  <xsl:choose>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='10'">
					  <演:渐变_6B67>
						  <xsl:call-template name="effectdur"/> 
					  </演:渐变_6B67>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='53'">
					  <演:渐变式缩放_6B6A>
						  <xsl:call-template name="effectdur"/>
					  </演:渐变式缩放_6B6A>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='55'">
					  <演:收缩_6BC8>
						  <xsl:call-template name="effectdur"/>
					  </演:收缩_6BC8>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='45'">
					  <演:渐变式回旋_6B69>
						  <xsl:call-template name="effectdur"/>
					  </演:渐变式回旋_6B69>
				  </xsl:when>
			  </xsl:choose>
		  </演:细微型_6C24>
	  </xsl:if>

	  <!-- 温和型 李娟   基本缩放23  回旋49 上浮47 下浮42 下沉37 中心旋转3 收缩旋转31 11.12.05-->
	  
	  <xsl:if test="descendant::p:cTn[@presetID]/@presetID='23' or descendant::p:cTn[@presetID]/@presetID='49' or descendant::p:cTn[@presetID]/@presetID='47' or descendant::p:cTn[@presetID]/@presetID='42' or  descendant::p:cTn[@presetID]/@presetID='37' or descendant::p:cTn[@presetID]/@presetID='3' or descendant::p:cTn[@presetID]/@presetID='31'">
		  <演:温和型_6C25>
			  <xsl:choose>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='23'">
					  <演:缩放_6B72>
						  <xsl:call-template name="effectdur"/>
					  </演:缩放_6B72>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='49'">
					  <演:回旋_6B75>
						  <xsl:call-template name="effectdur"/>
					  </演:回旋_6B75>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='47'">
					  <演:上升_6B76>
						  <xsl:call-template name="effectdur"/>
					  </演:上升_6B76>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='42'">
					  <演:下降_6B78>
						  <xsl:call-template name="effectdur"/>
					  </演:下降_6B78>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='37'">
					  <演:下沉_6BCE>
						  <xsl:call-template name="effectdur"/>
					  </演:下沉_6BCE>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='3'">
					  <演:中心旋转_6B7B>
						  <xsl:call-template name="effectdur"/>
					  </演:中心旋转_6B7B>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='31'">
					  <演:翻转时由近及远_6BCB>
						  <xsl:call-template name="effectdur"/>
					  </演:翻转时由近及远_6BCB>
				  </xsl:when>
            </xsl:choose>
		  </演:温和型_6C25>
	  </xsl:if>
	  <!--李娟 添加 华丽型   11.12.05 -->
	  <!-- 李娟 11.12.05 字幕式28 向下曲线52 空翻56 浮动30 下拉38转为挥舞 螺旋飞出15 飞旋25 弹跳26   挥鞭式41  玩具风车35 基本旋转19转换为-->
	  <xsl:if test="descendant::p:cTn[@presetID]/@presetID='28' or descendant::p:cTn[@presetID]/@presetID='52' or descendant::p:cTn[@presetID]/@presetID='56' or descendant::p:cTn[@presetID]/@presetID='30' or  descendant::p:cTn[@presetID]/@presetID='38' or descendant::p:cTn[@presetID]/@presetID='15' or descendant::p:cTn[@presetID]/@presetID='25' or descendant::p:cTn[@presetID]/@presetID='26' or descendant::p:cTn[@presetID]/@presetID='41' or descendant::p:cTn[@presetID]/@presetID='35' or descendant::p:cTn[@presetID]/@presetID='19'">
		  <演:华丽型_6C26>
			  <xsl:choose>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='28'">
					  <演:字幕式_6B87>
						  <xsl:call-template name="effectdur"/>
					  </演:字幕式_6B87>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='52'">
					  <演:向下曲线_6BD0>
						  <xsl:call-template name="effectdur"/>
					  </演:向下曲线_6BD0>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='56'">
					  <演:空翻_6B81>
						  <xsl:call-template name="effectdur"/>
					  </演:空翻_6B81>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='30'">
					  <演:浮动_6B8A>
						  <xsl:call-template name="effectdur"/>
					  </演:浮动_6B8A>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='38'">
					  <演:挥舞_6B8C>
						  <xsl:call-template name="effectdur"/>
					  </演:挥舞_6B8C>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='15'">
					  <演:螺旋飞出_6BCF>
						  <xsl:call-template name="effectdur"/>
					  </演:螺旋飞出_6BCF>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='25'">
					  <演:飞旋_6B7E>
						  <xsl:call-template name="effectdur"/>
					  </演:飞旋_6B7E>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='26'">
					  <演:弹跳_6B7D>
						  <xsl:call-template name="effectdur"/>
					  </演:弹跳_6B7D>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='41'">
					  <演:挥鞭式_6B80>
						  <xsl:call-template name="effectdur"/>
					  </演:挥鞭式_6B80>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='35'">
					  <演:玩具风车_6B83>
						  <xsl:call-template name="effectdur"/>
					  </演:玩具风车_6B83>
				  </xsl:when>
				  <xsl:when test="descendant::p:cTn[@presetID]/@presetID='19'">
					  <演:旋转_6B84>
						  <xsl:call-template name="effectdur"/>
					  </演:旋转_6B84>
				  </xsl:when>
			  </xsl:choose>
		  </演:华丽型_6C26>
	  </xsl:if>
  </xsl:template>

  <!--xsl:template name="path">
    <xsl:param name="val"/>
    <xsl:param name="w"/>
    <xsl:param name="h"/>
    <xsl:param name="step"/>
    <xsl:choose>
      <xsl:when test="not(contains($val,' '))">
        <xsl:choose>
          <xsl:when test="string($val * $h)!='NaN'">
            
            <xsl:call-template name="computepathvalue">
              <xsl:with-param name="val" select="$val * $h * 1.34"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$val"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="val1" select="substring-before($val,' ')"/>
        <xsl:choose>
          <xsl:when test="string($val1 * $w)!='NaN' and string($val1 * $h)!='NaN'">
            <xsl:choose>
              <xsl:when test="$step='1'">
               
                <xsl:call-template name="computepathvalue">
                  <xsl:with-param name="val" select="$val1 * $w * 1.34"/>
                </xsl:call-template>
                <xsl:call-template name="path">
                  <xsl:with-param name="val" select="substring-after($val,' ')"/>
                  <xsl:with-param name="w" select="$w"/>
                  <xsl:with-param name="h" select="$h"/>
                  <xsl:with-param name="step" select="2"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$step='2'">
                
                <xsl:call-template name="computepathvalue">
                  <xsl:with-param name="val" select="$val1 * $h * 1.34"/>
                </xsl:call-template>
                <xsl:call-template name="path">
                  <xsl:with-param name="val" select="substring-after($val,' ')"/>
                  <xsl:with-param name="w" select="$w"/>
                  <xsl:with-param name="h" select="$h"/>
                  <xsl:with-param name="step" select="1"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat($val1,' ')"/>
            <xsl:call-template name="path">
              <xsl:with-param name="val" select="substring-after($val,' ')"/>
              <xsl:with-param name="w" select="$w"/>
              <xsl:with-param name="h" select="$h"/>
              <xsl:with-param name="step" select="1"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template-->

  <!--xsl:template name="computepathvalue">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='0'">
        <xsl:value-of select="concat($val,' ')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat(format-number($val,'#.##'),' ')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template-->
  
  <xsl:template name="effectdur">
    <xsl:attribute name="速度_6B44">
      <xsl:call-template name="dur">
        <xsl:with-param name="loc" select="速度_6B44"/>
        <xsl:with-param name="dur" select=".//*[@dur!='1']/@dur"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template name="dur">
    <xsl:param name="loc"/>
    <xsl:param name="dur"/>
    <xsl:choose>
      <xsl:when test="$loc='速度_6B44'">
        <xsl:choose>
          <xsl:when test="$dur='5000'">very-slow</xsl:when>
          <xsl:when test="$dur='3000'">slow</xsl:when>
          <xsl:when test="$dur='2000'">medium</xsl:when>
          <xsl:when test="$dur='1000'">fast</xsl:when>
          <xsl:when test="$dur='500'">very-fast</xsl:when>
          <xsl:otherwise>medium</xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$dur='5000'">very-slow</xsl:when>
          <xsl:when test="$dur='3000'">slow</xsl:when>
          <xsl:when test="$dur='2000'">medium</xsl:when>
          <xsl:when test="$dur='1000'">fast</xsl:when>
          <xsl:when test="$dur='500'">very-fast</xsl:when>
          <xsl:otherwise>medium</xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="target">
    <!--2010-11-23 罗文甜，修改bug：spid的精确定位问题，避免跟触发器动画id值冲突-->
    <xsl:variable name="spid" select="descendant::p:cBhvr/p:tgtEl/p:spTgt/@spid"/>
    <!--2010.04.30<xsl:if test="descendant::p:spTgt/p:txEl/p:pRg[@st and @end and @st!='0']">-->
    <xsl:if test="descendant::p:cBhvr/p:tgtEl/p:spTgt/p:txEl/p:pRg[@st]">
      <xsl:attribute name="段落引用_6C27">
        <!--<xsl:variable name="st" select="descendant::p:spTgt/p:txEl/p:pRg/@st"/>-->        
        <xsl:variable name="st" select="descendant::p:cBhvr/p:tgtEl/p:spTgt/p:txEl/p:pRg/@st+1"/>
        <!--2010-11-23 罗文甜：修改节点定位（增加母版动画的引用后）-->
        <!--xsl:for-each select="ancestor::p:sld//p:cNvPr[@id=$spid]/ancestor::p:sp/p:txBody/a:p[a:r/a:t]"-->
         <xsl:for-each select="ancestor::p:timing/preceding-sibling::p:cSld//p:cNvPr[@id=$spid]/ancestor::p:sp/p:txBody/a:p[a:r/a:t]">
          <xsl:if test="position()=$st">
            <xsl:value-of select="@id"/>
          </xsl:if>
        </xsl:for-each>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name="对象引用_6C28">
      <!--2010-11-23 罗文甜：修改节点定位（增加母版动画的引用后）-->
      <xsl:choose>
        <xsl:when test="ancestor::p:timing/preceding-sibling::p:cSld//p:cNvPr[@id=$spid]/parent::*/parent::*/@id = '' ">
          <xsl:value-of select="ancestor::p:timing//p:cSld//p:cNvPr[@id=$spid]/parent::*/parent::*/@id"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="ancestor::p:timing/preceding-sibling::p:cSld//p:cNvPr[@id=$spid]/parent::*/parent::*/@id"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
</xsl:stylesheet>
