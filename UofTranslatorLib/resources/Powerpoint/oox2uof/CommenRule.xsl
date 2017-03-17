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
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:规则="http://schemas.uof.org/cn/2009/rules">
  <xsl:import href="sldLayout.xsl"/>
  <xsl:import href="txStyles.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="CommenRule">
    <规则:演示文稿_B66D 演:liuyin="liuyin">
      <!--liuyin 20121109 解决多出命名空间问题 <演:纸张方向_6BE1 xmlns:演="http://schemas.uof.org/cn/2009/presentation">portrait</演:纸张方向_6BE1>-->

      <规则:页面设置集_B670>
        <规则:页面设置_B638>
          <xsl:if test="p:presentation/p:sldSz">
            <xsl:call-template name="sldSz"/>
          </xsl:if>
        </规则:页面设置_B638>
      </规则:页面设置集_B670>
      <!-- 罗文甜：配色方案集移到母版里-->
      <规则:页面版式集_B651>
        <xsl:apply-templates select="p:presentation/p:sldMaster/p:sldLayout"/>
      </规则:页面版式集_B651>
      <!--演:文本式样集 这段给取消了 李娟20111105-->
      <!--<演:文本式样集 uof:locID="p0131">
        <xsl:apply-templates select="p:presentation/p:sldMaster/p:txStyles">
          <xsl:with-param name="ID" select="p:presentation/p:sldMaster/@id"/>
        </xsl:apply-templates>
      </演:文本式样集>-->

      <!--最后视图-->
      <规则:最后视图_B639>
        <xsl:apply-templates select="/p:presentation/p:viewPr"/>
      </规则:最后视图_B639>
      <规则:显示比例_B63F>
        <xsl:apply-templates select="//p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx" mode="sldview"/>
      </规则:显示比例_B63F>

      <!--放映设置>
			<xsl:apply-templates select="p:presentation/p:custShowLst"/-->
      <规则:放映设置_B653>
        <xsl:apply-templates select="p:presentation/p:custShowLst/p:custShow" mode="list"/>
        <xsl:apply-templates select="/p:presentation/p:presentationPr"/>
      </规则:放映设置_B653>

      <!--2014-04-09, tangjiang, 添加网格线的转换 start -->
      <xsl:if test="/p:presentation/p:viewPr/p:gridSpacing/@cx">
        <规则:绘图网格与参考线_B602>
          <xsl:attribute name="是否显示参考线_C600">false</xsl:attribute>
          <xsl:attribute name="对象是否对齐绘图网格_C601">false</xsl:attribute>
          <xsl:attribute name="是否显示绘图网格_C602">false</xsl:attribute>
          <xsl:variable name="gridSpace" select="/p:presentation/p:viewPr/p:gridSpacing/@cx"/>
          <xsl:if test="$gridSpace">
            <xsl:attribute name="绘图网格间距_C603">
              <xsl:value-of select="$gridSpace div 12700"/>
            </xsl:attribute>
          </xsl:if>
        </规则:绘图网格与参考线_B602>
      </xsl:if>
      <!--end 2014-04-09, tangjiang, 添加网格线的转换 -->
      
      <!--2010-11-8 罗文甜：增加页眉页脚集-->
      <!--2012-12-20, liqiuling, 解决OOXML到UOF备注页页脚丢失  start -->

      <规则:页眉页脚集_B640>        
        <!-- 母板的页眉页脚  -->
        <xsl:for-each select="/p:presentation/p:sldMaster/p:hf">
          <xsl:call-template name="MasterSlidehf">
            <xsl:with-param name="hf" select="/p:presentation/p:sldMaster/p:hf"/>
          </xsl:call-template>
        </xsl:for-each>
        
        <!-- 幻灯片的页眉页脚  -->
        <xsl:for-each select="/p:presentation/p:sld">
          <xsl:call-template name="Slidehf">
            <xsl:with-param name="sld" select="."/>
          </xsl:call-template>
        </xsl:for-each>
        
        <!-- 讲义和备注的页眉页脚  -->
        <xsl:for-each select="/p:presentation/p:notesMaster/p:hf">
          <xsl:call-template name="NHhf">
            <xsl:with-param name="hf" select="/p:presentation/p:notesMaster/p:hf"/>
          </xsl:call-template>
        </xsl:for-each>
        <xsl:for-each select="/p:presentation/p:handoutMaster/p:hf">
          <xsl:call-template name="NHhf">
            <xsl:with-param name="hf" select="/p:presentation/p:handoutMaster/p:hf"/>
          </xsl:call-template>
        </xsl:for-each>
        <!--liqiuling  备注页页脚-->
        <xsl:for-each select="/p:presentation/p:notes">
          <xsl:call-template name="noteshf"/>
        </xsl:for-each>


      </规则:页眉页脚集_B640>
      <!--end-->
    </规则:演示文稿_B66D>

  </xsl:template>
  <!--xsl:template match="p:custShowLst">
		<演:放映设置 uof:locID="p0021">
			<xsl:apply-templates select="p:custShow" mode="list"/>
			<xsl:apply-templates select="document('ppt/presProps.xml')/p:presentationPr"/>
		</演:放映设置>
	</xsl:template-->

  <xsl:template match="p:custShow" mode="list">
    <规则:幻灯片序列_B654>
      <xsl:attribute name="标识符_B655">
        <xsl:value-of select="concat('customListId',@id)"/>
      </xsl:attribute>
      <xsl:attribute name="名称_B656">
        <xsl:value-of select="@name"/>
      </xsl:attribute>
      <xsl:attribute name="是否自定义_B657">
        <xsl:value-of select="'true'"/>
      </xsl:attribute>
      <!--2010.05.04<xsl:apply-templates select="p:sldLst/p:sld" mode="sldLst"/>-->
      <xsl:for-each select="p:sldLst/p:sld">
        <xsl:apply-templates select="." mode="sldLst"/>
      </xsl:for-each>
    </规则:幻灯片序列_B654>
  </xsl:template>

  <!--有专门的文件存放幻灯片放映序列-->
  <xsl:template match="p:sld" mode="sldLst">
    <xsl:variable name="sldShowLst-rId" select="@r:id"/>
    <!--错误xsl:value-of select="concat(substring-before(substring-after(document('_rels/presentation.xml.rels')/rel:Relationships/rel:Relationship[@Id=$sldShowLst-rId]/@Target,'/'),'.xml'),' '"/-->
    <!--改变位置xsl:value-of select="concat(substring-before(substring-after(document('_rels/presentation.xml.rels')/rel:Relationships/node()[@Id=$sldShowLst-rId]/@Target,'/'),'.xml'),' ')"/-->
    <!--<xsl:value-of select="concat(substring-before(substring-after(document('ppt/_rels/presentation.xml.rels')/rel:Relationships/node()[@Id=$sldShowLst-rId]/@Target,'/'),'.xml'),' ')"/>-->
    <!--2010.05.04 -->
    <xsl:variable name="slideId">
      <xsl:value-of select="substring-after(/p:presentation/rel:Relationships[1]/node()[@Id=$sldShowLst-rId]/@Target,'/')"/>
    </xsl:variable>
    <xsl:for-each select="/p:presentation/p:sld">
      <xsl:if test="@id=$slideId">
        <xsl:value-of select="concat(position()-1,' ')"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="p:presentationPr">
    <xsl:if test="p:showPr">
      <xsl:apply-templates select="p:showPr"/>
    </xsl:if>
    <xsl:if test="not(p:showPr)">
      <规则:绘图笔颜色_B661>#FF0000</规则:绘图笔颜色_B661>
    </xsl:if>
  </xsl:template>

  <xsl:template match="p:showPr">
    <xsl:if test ="p:sldRg">
      <规则:幻灯片序列_B654>
        <!-- 09.12.14 马有旭 修改 -->
        <xsl:call-template name="sldLst"/>
        <!--<xsl:value-of select="concat(substring-before(substring-after(document('ppt/_rels/presentation.xml.rels')/rel:Relationships/node()[@Id=document('ppt/presentation.xml')/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@st)]/@r:id]/@Target,'/'),'.xml'),' ')"/>
        <xsl:value-of select="substring-before(substring-after(document('ppt/_rels/presentation.xml.rels')/rel:Relationships/node()[@Id=document('ppt/presentation.xml')/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@end)]/@r:id]/@Target,'/'),'.xml')"/>-->
      </规则:幻灯片序列_B654>
      <!--<演:放映顺序 uof:locID="p0023" uof:attrList="名称 序列引用" 演:序列引用="sequentListId"/>-->
      <!--演：放映顺序是不是可以删除-->
    </xsl:if>
    <xsl:if test="p:custShow">
      <规则:放映顺序_B658>
        <!--<xsl:attribute name="演:序列引用">-->
        <xsl:value-of select="concat('customListId',p:custShow/@id)"/>
        <!--</xsl:attribute>-->
      </规则:放映顺序_B658>
    </xsl:if>
    <xsl:if test="p:present|p:kiosk">
      <规则:是否全屏放映_B659>true</规则:是否全屏放映_B659>
    </xsl:if>
    <xsl:if test="(@loop='true') or (@loop='1') or (@loop='on')">
      <规则:是否循环放映_B65A>true</规则:是否循环放映_B65A>
    </xsl:if>
    <!-- 09.10.19 马有旭 修改 -->
    <xsl:choose>
      <xsl:when test="not(@showAnimation) or (@showAnimation='true') or (@showAnimation='1') or (@showAnimation='on')">
        <规则:是否放映动画_B65E>true</规则:是否放映动画_B65E>
      </xsl:when>
      <xsl:otherwise>
        <规则:是否放映动画_B65E>false</规则:是否放映动画_B65E>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="p:kiosk">
      <规则:放映间隔_B65B>
        <xsl:if test="p:kiosk/@restart">
          <xsl:value-of select="p:kiosk/@restart"/>
        </xsl:if>
        <xsl:if test="not(p:kiosk/@restart)">300000</xsl:if>
      </规则:放映间隔_B65B>
    </xsl:if>
    <xsl:if test="(@useTimings='0') or (@useTimings='false') or (@useTimings='off')">
      <!--oxml、uof中单位都为毫秒-->
      <规则:是否手动方式_B65C>true</规则:是否手动方式_B65C>
    </xsl:if>
    <!--罗文甜：增加绘图笔颜色-->
    <xsl:if test="p:penClr">
      <规则:绘图笔颜色_B661>
        <!--2011-2-18罗文甜，修改绘图笔其他颜色-->
        <xsl:for-each select="p:penClr">
          <xsl:call-template name="colorChoice"/>
        </xsl:for-each>
        <!--xsl:if test="p:penClr/a:srgbClr/@val">
          <xsl:value-of select="concat('#',p:penClr/a:srgbClr/@val)"/>
        </xsl:if>       
        <xsl:if test="p:penClr/a:schemeClr/@val">
          <xsl:value-of select="concat('#',p:penClr/a:srgbClr/@val)"/>
        </xsl:if-->
      </规则:绘图笔颜色_B661>
    </xsl:if>
    <xsl:if test="not(p:penClr)">
      <规则:绘图笔颜色_B661>#FF0000</规则:绘图笔颜色_B661>
    </xsl:if>
  </xsl:template>
  <xsl:template name="sldSz">
    <!-- 09.11.08 马有旭 添加纸张类型判断 -->
    <!-- 10.02.04 马有旭 Bug ID:2944443 -->
    <xsl:attribute name="标识符_B671">IDymsz</xsl:attribute>
    <xsl:if test="/p:presentation/p:sldSz/@cy or /p:presentation/p:sldSz/@cx">
      <演:纸张_6BDD>
        <xsl:choose>
          <xsl:when test="p:presentation/p:sldSz/@cy and p:presentation/p:sldSz/@cx">
            <xsl:choose>
              <xsl:when test="p:presentation/p:sldSz/@type='A3'">
                <xsl:if test="p:presentation/p:sldSz/@cx='9601200'">
                  <!--<xsl:attribute name="uof:纸型">A3</xsl:attribute>-->
                  <xsl:attribute name="宽_C605">756</xsl:attribute>
                  <xsl:attribute name="长_C604">1008</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='9601200'">
                  <!--<xsl:attribute name="uof:纸型">A3</xsl:attribute>-->
                  <xsl:attribute name="长_C604">756</xsl:attribute>
                  <xsl:attribute name="宽_C605">1008</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='A4'">
                <xsl:if test="p:presentation/p:sldSz/@cx='9906000'">
                  <!--<xsl:attribute name="uof:纸型">A4</xsl:attribute>-->
                  <xsl:attribute name="宽_C605">779</xsl:attribute>
                  <xsl:attribute name="长_C604">540</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='9906000'">
                  <!--<xsl:attribute name="uof:纸型">A4</xsl:attribute>-->
                  <xsl:attribute name="长_C604">779</xsl:attribute>
                  <xsl:attribute name="宽_C605">540</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='B4ISO'">
                <xsl:if test="p:presentation/p:sldSz/@cx='8120063'">
                  <!--<xsl:attribute name="uof:纸型">B4</xsl:attribute>-->
                  <xsl:attribute name="宽_C605">639</xsl:attribute>
                  <xsl:attribute name="长_C604">852</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='8120063'">
                  <!--<xsl:attribute name="uof:纸型">B4</xsl:attribute>-->
                  <xsl:attribute name="长_C604">639</xsl:attribute>
                  <xsl:attribute name="宽_C605">852</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='B5ISO'">
                <xsl:if test="p:presentation/p:sldSz/@cx='5376863'">
                  <!--<xsl:attribute name="uof:纸型">B5</xsl:attribute>-->
                  <xsl:attribute name="宽_C605">423</xsl:attribute>
                  <xsl:attribute name="长_C604">564</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='5376863'">
                  <!--<xsl:attribute name="uof:纸型">B5</xsl:attribute>-->
                  <xsl:attribute name="长_C604">423</xsl:attribute>
                  <xsl:attribute name="宽_C605">564</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='letter'">
                <xsl:if test="p:presentation/p:sldSz/@cx='9144000'">
                  <!--<xsl:attribute name="uof:纸型">letter</xsl:attribute>-->
                  <xsl:attribute name="宽_C605">720</xsl:attribute>
                  <xsl:attribute name="长_C604">540</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='9144000'">

                  <!--<xsl:attribute name="uof:纸型">letter</xsl:attribute>-->
                  <xsl:attribute name="长_C604">720</xsl:attribute>
                  <xsl:attribute name="宽_C605">540</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='ledger'">
                <xsl:if test="p:presentation/p:sldSz/@cx='9134475'">
                  <xsl:attribute name="宽_C605">719</xsl:attribute>
                  <xsl:attribute name="长_C604">959</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='9134475'">
                  <xsl:attribute name="长_C604">719</xsl:attribute>
                  <xsl:attribute name="宽_C605">959</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='35mm'">
                <xsl:if test="p:presentation/p:sldSz/@cx='6858000'">
                  <xsl:attribute name="宽_C605">540</xsl:attribute>
                  <xsl:attribute name="长_C604">809</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='6858000'">
                  <xsl:attribute name="长_C604">540</xsl:attribute>
                  <xsl:attribute name="宽_C605">809</xsl:attribute>
                </xsl:if>
              </xsl:when>

              <xsl:when test="p:presentation/p:sldSz/@type='banner'">
                <xsl:if test="p:presentation/p:sldSz/@cx='914400'">
                  <xsl:attribute name="宽_C605">72</xsl:attribute>
                  <xsl:attribute name="长_C604">576</xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy='914400'">
                  <xsl:attribute name="长_C604">72</xsl:attribute>
                  <xsl:attribute name="宽_C605">576</xsl:attribute>
                </xsl:if>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="p:presentation/p:sldSz/@cx">
                  <xsl:attribute name="宽_C605">
                    <!--单位换算：1 pt=127000 EMUs-->
                    <xsl:value-of select="round(/p:presentation/p:sldSz/@cx div 12700)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="p:presentation/p:sldSz/@cy">
                  <xsl:attribute name="长_C604">
                    <!--单位换算：1 pt=127000 EMUs-->
                    <xsl:value-of select="round(/p:presentation/p:sldSz/@cy div 12700)"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>

        </xsl:choose>
      </演:纸张_6BDD>

    </xsl:if>
    <xsl:if test="/p:presentation/@firstSlideNum">
      <演:页码格式_6BDF>
        <xsl:attribute name="起始编号_6BE0">
          <xsl:value-of select="/p:presentation/@firstSlideNum"/>
        </xsl:attribute>
      </演:页码格式_6BDF>
    </xsl:if>
    <!-- liuyin 20121109 备注讲义大纲方向错误 <xsl:if test="/p:presentation/p:notesSz/@cy &lt;/p:presentation/p:notesSz/@cx">
    <xsl:if test="(/p:presentation/p:sldSz/@cy &gt;/p:presentation/p:sldSz/@cx) and (/p:presentation/p:notesSz/@cy &lt;/p:presentation/p:notesSz/@cx)">-->
    <xsl:choose>
      <xsl:when test="(/p:presentation/p:notesSz/@cy &lt;/p:presentation/p:notesSz/@cx)">
        <演:纸张方向_6BE1>
          <xsl:value-of select="'landscape'"/>
        </演:纸张方向_6BE1>
      </xsl:when>
      <xsl:when test="(/p:presentation/p:notesSz/@cy &gt;/p:presentation/p:notesSz/@cx)">
        <演:纸张方向_6BE1>
          <xsl:value-of select="'portrait'"/>
        </演:纸张方向_6BE1>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="p:viewPr">
    <!--<规则:最后视图_B639 >-->
    <!--UOF2.0 最后视图无sldThumbnailView 对应到note-page？lijuan-->
    <规则:类型_B63A>
      <xsl:choose>
        <xsl:when test="@lastView='sldMasterView'">slide-master</xsl:when>
        <xsl:when test="@lastView='notesView'">note-page</xsl:when>
        <xsl:when test="@lastView='handoutView'">handout-master</xsl:when>
        <xsl:when test="@lastView='notesMasterView'">note-master</xsl:when>
        <xsl:when test="@lastView='sldSorterView'">sort</xsl:when>
        <!--<xsl:when test="@lastView='sldThumbnailView'">note-page</xsl:when>-->
        <xsl:when test="@lastView='outlineView'">outline</xsl:when>
        <!--'sldView''outlineView''sldThumbnailView'-->
        <xsl:when test="not(@lastView)">normal</xsl:when>
        <xsl:otherwise>normal</xsl:otherwise>
      </xsl:choose>
    </规则:类型_B63A>
    <!--</规则:最后视图_B639>-->
    <!--<xsl:apply-templates select="p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx" mode="sldview"/>-->
  </xsl:template>
  <xsl:template match="a:sx" mode="sldview">
    <!--<xsl:value-of select="@n"/>-->
    <xsl:value-of select="round((@n div @d)*100)"/>

  </xsl:template>
  <!-- 09.12.14 马有旭 修改 -->
  <!--10.1.13 黎美秀修改 
  <xsl:template name="sldLst">
    <xsl:variable name="st">
      <xsl:choose>
        <xsl:when test="/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@st)]]">
          <xsl:value-of select="concat(substring-before(substring-after(/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@st)]/@r:id]/@Target,'/'),'.xml'),' ')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat(concat(substring-before(substring-after(/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[1]/@r:id]/@Target,'/'),'.xml'),' '),' ')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="end">
      <xsl:choose>
        <xsl:when test="/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@end)]/@r:id]">
          <xsl:value-of select="substring-before(substring-after(/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@end)]/@r:id]/@Target,'/'),'.xml')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat(concat(substring-before(substring-after(/p:presentation/Relationships[1]/rel:Relationships/node()[@Id=/p:presentation/p:sldIdLst/node()[1]/@r:id]/@Target,'/'),'.xml'),' '),' ')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="concat($st,' ',$end)"/>
  </xsl:template>-->
  <xsl:template name="sldLst">
    <xsl:variable name="st">
      <xsl:choose>
        <xsl:when test="/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@st)]]">
          <xsl:value-of select="concat(substring-before(substring-after(/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@st)]/@r:id]/@Target,'/'),'.xml'),' ')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat(concat(substring-before(substring-after(/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[1]/@r:id]/@Target,'/'),'.xml'),' '),' ')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="end">
      <xsl:choose>
        <xsl:when test="/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@end)]/@r:id]">
          <xsl:value-of select="substring-before(substring-after(/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[number(current()/p:sldRg/@end)]/@r:id]/@Target,'/'),'.xml')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat(concat(substring-before(substring-after(/p:presentation/rel:Relationships[1]/node()[@Id=/p:presentation/p:sldIdLst/node()[1]/@r:id]/@Target,'/'),'.xml'),' '),' ')"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="concat($st,' ',$end)"/>
  </xsl:template>
  <!--2010-11-8 罗文甜：增加幻灯片页眉页脚集-->
  <xsl:template name="MasterSlidehf">
    <xsl:param name ="hf"/>
    <!--xsl:for-each select="/p:presentation/p:sld"-->
    <规则:幻灯片_B641>
      <xsl:attribute name="类型_B645">presentation</xsl:attribute>
      <!--<xsl:attribute name="标识符_B646">mrYMYJ</xsl:attribute>-->
      <xsl:if test="//p:hf">
        <xsl:attribute name="标识符_B646">
          <xsl:value-of select="generate-id(//p:hf)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="$hf/@dt='0' or not($hf)">
          <xsl:attribute name="是否显示日期和时间_B647">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示日期和时间_B647">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="$hf/@ftr='0' or not($hf)">
          <xsl:attribute name="是否显示页脚_B648">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页脚_B648">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <!--2011-01-04罗文甜：增加判断-->
      <xsl:choose>
        <xsl:when test="$hf/@sldNum='0' or not($hf)">
          <xsl:attribute name="是否显示幻灯片编号_B64A">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示幻灯片编号_B64A">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="/p:presentation/@showSpecialPlsOnTitleSld='0' or not(/p:presentation/@showSpecialPlsOnTitleSld)">
          <xsl:attribute name="标题幻灯片中是否显示_B64B">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="标题幻灯片中是否显示_B64B">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="$hf/parent::*/descendant::p:ph[@type='dt']">
        <xsl:choose>
          <xsl:when test="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']/p:txBody/a:p/a:fld[contains(@type,'date')]">
            <xsl:attribute name="是否自动更新日期和时间_B649">true</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="是否自动更新日期和时间_B649">false</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:variable name="dt">
          <xsl:for-each select="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']//a:t">
            <xsl:value-of select="text()"/>
          </xsl:for-each>
        </xsl:variable>

        <!--2014-03-26, tangjiang, 修复日期时间的转换 start -->
        <xsl:variable name="dateFormate">
          <xsl:for-each select="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']//a:fld">
            <xsl:value-of select="@type"/>
          </xsl:for-each>
        </xsl:variable>
        <xsl:if test="$hf and not($hf/@dt)">
          <规则:日期和时间格式_B642>
            <xsl:choose>
              <!-- yyyy年M月d日星期Wh时mm分s秒  -->
              <xsl:when test="$dateFormate='datetime9'">
                <xsl:value-of select="'yyyy年M月d日星期Wh时mm分s秒'"/>
              </xsl:when>
              <!-- yyyy年M月d日星期W -->
              <xsl:when test="$dateFormate='datetime3'">
                <xsl:value-of select="'yyyy年M月d日星期W'"/>
              </xsl:when>
              <!-- yyyy年M月d日H时mm分 -->
              <xsl:when test="$dateFormate='datetime8'">
                <xsl:value-of select="'yyyy年M月d日H时mm分'"/>
              </xsl:when>
              <!-- 上午/下午h时mm分ss秒 -->
              <xsl:when test="$dateFormate='datetime13'">
                <xsl:value-of select="'上午/下午h时mm分ss秒'"/>
              </xsl:when>
              <!-- 上午/下午h时m分 -->
              <xsl:when test="$dateFormate='datetime12'">
                <xsl:value-of select="'上午/下午h时m分'"/>
              </xsl:when>
              <!-- yyyy年M月d日 -->
              <xsl:when test="$dateFormate='datetime2'">
                <xsl:value-of select="'yyyy年M月d日'"/>
              </xsl:when>
              <!-- yyyy年M月 -->
              <xsl:when test="$dateFormate='datetime6'">
                <xsl:value-of select="'yyyy年M月'"/>
              </xsl:when>
              <!-- yyyy/M/d -->
              <xsl:otherwise>
                <xsl:value-of select="'yyyy/M/d'"/>
              </xsl:otherwise>
            </xsl:choose>
          </规则:日期和时间格式_B642>
          <规则:日期和时间字符串_B643>
            <xsl:value-of select="$dt"/>
          </规则:日期和时间字符串_B643>
        </xsl:if>
        <!--end 2014-03-26, tangjiang, 修复日期时间的转换 -->

      </xsl:if>

      <!--2011-2-22罗文甜，增加判断页脚是否存在-->
      <xsl:if test="$hf and not($hf/@ftr)  and //p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']//a:t ">
        <规则:页脚_B644>
          <xsl:value-of select="//p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']//a:t"/>
        </规则:页脚_B644>
      </xsl:if>

    </规则:幻灯片_B641>
    <!--/xsl:for-each-->
  </xsl:template>
  <!--2010-11-8罗文甜：增加讲义和备注页眉页脚-->
  <xsl:template name="NHhf">
    <xsl:param name ="hf"/>
    <!--xsl:if test="/p:presentation/p:notesMaster/p:hf"-->
    <规则:讲义和备注_B64C>
      <xsl:attribute name="标识符_B646">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <xsl:attribute name="是否显示日期和时间_B647">
        <xsl:choose>
          <xsl:when test="$hf/@dt='0' or not($hf)">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:choose>
        <xsl:when test="$hf/@ftr='0' or not($hf)">
          <xsl:attribute name="是否显示页脚_B648">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页脚_B648">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="$hf/@hdr='0' or not($hf)">
          <xsl:attribute name="是否显示页眉_B64F">false</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页眉_B64F">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:attribute name="是否显示页码_B650">
        <xsl:choose>
          <xsl:when test="$hf/@sldNum='0' or not($hf)">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:if test="not($hf/@dt) or $hf/@dt!='0'">
        <xsl:choose>
          <xsl:when test="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']/p:txBody/a:p/a:fld[contains(@type,'date')]">
            <xsl:attribute name="是否自动更新日期和时间_B649">true</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="是否自动更新日期和时间_B649">false</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>

      <xsl:if test="$hf and not($hf/@hdr) and $hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='hdr']//a:t">
        <规则:页眉_B64D>
          <xsl:value-of select="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='hdr']//a:t/text()"/>
        </规则:页眉_B64D>
      </xsl:if>
      <xsl:if test="$hf and not($hf/@ftr) and $hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']//a:t">
        <规则:页脚_B644>
          <!--<xsl:value-of select="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']//a:t/text()"/>-->
        </规则:页脚_B644>
      </xsl:if>
      <xsl:variable name="dt">
        <xsl:for-each select="$hf/parent::*/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']//a:t">
          <xsl:value-of select="text()"/>
        </xsl:for-each>
      </xsl:variable>
      <xsl:if test="$hf and not($hf/@dt)">
        <规则:日期和时间字符串_B643>
          <xsl:value-of select="$dt"/>
        </规则:日期和时间字符串_B643>
      </xsl:if>
    </规则:讲义和备注_B64C>
    <!--/xsl:if-->
  </xsl:template>
  <!--2012-12-20, liqiuling, 解决OOXML到UOF备注页页脚丢失  2013-4-1 liqiuling start -->
  <xsl:template name="noteshf">
    <规则:讲义和备注_B64C>
      <xsl:attribute name="标识符_B646">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <xsl:attribute name="是否显示日期和时间_B647">
        <xsl:choose>
          <xsl:when test=".//p:sp[.//p:ph/@type ='dt']">true</xsl:when>
          <xsl:otherwise>false</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:choose>
        <xsl:when test=".//p:sp[descendant::p:ph/@type='ftr']">
          <xsl:attribute name="是否显示页脚_B648">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页脚_B648">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test=".//p:sp[descendant::p:ph/@type='hdr']">
          <xsl:attribute name="是否显示页眉_B64F">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页眉_B64F">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:attribute name="是否显示页码_B650">
        <xsl:choose>
          <xsl:when test=".//p:sp[descendant::p:ph/@type='sldNum']/p:txBody//a:t">true</xsl:when>
          <xsl:otherwise>false</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:if test=".//p:sp[descendant::p:ph/@type='hdr']//a:t">
        <规则:页眉_B64D>
          <xsl:value-of select=".//p:sp[.//p:ph/@type='hdr']/p:txBody//a:t/text()"/>
        </规则:页眉_B64D>
      </xsl:if>
      <xsl:if test=".//p:sp[descendant::p:ph/@type='ftr']//a:t">
        <规则:页脚_B644>
          <xsl:value-of select=".//p:sp[descendant::p:ph/@type='ftr']/p:txBody//a:t/text()"/>
        </规则:页脚_B644>
      </xsl:if>
      <xsl:if test=".//p:sp[descendant::p:ph/@type='dt']//a:t">
        <规则:日期和时间字符串_B643>
          <xsl:value-of select=".//p:sp[descendant::p:ph/@type='dt']/p:txBody//a:t/text()"/>
        </规则:日期和时间字符串_B643>
      </xsl:if>

    </规则:讲义和备注_B64C>

  </xsl:template>
  <!--2012-12-20, liqiuling, 解决OOXML到UOF备注页页脚丢失  2013-4-1 liqiuling end -->

  <!--2014-06-02, tangjiang, 添加幻灯片页眉页脚 start-->
  <xsl:template name="Slidehf">
    <xsl:param name ="sld"/>
    <规则:幻灯片_B641>
      <xsl:attribute name="类型_B645">slide</xsl:attribute>
      <xsl:attribute name="标识符_B646">
        <xsl:value-of select="generate-id($sld)"/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="$sld/p:cSld/p:spTree/p:sp[./p:nvSpPr/p:nvPr/p:ph/@type='dt']">
          <xsl:attribute name="是否显示日期和时间_B647">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示日期和时间_B647">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="$sld/p:cSld/p:spTree/p:sp[./p:nvSpPr/p:nvPr/p:ph/@type='ftr']">
          <xsl:attribute name="是否显示页脚_B648">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示页脚_B648">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <!--2011-01-04罗文甜：增加判断-->
      <xsl:choose>
        <xsl:when test="$sld/p:cSld/p:spTree/p:sp[./p:nvSpPr/p:nvPr/p:ph/@type='sldNum']">
          <xsl:attribute name="是否显示幻灯片编号_B64A">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否显示幻灯片编号_B64A">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="/p:presentation/@showSpecialPlsOnTitleSld='0' or not(/p:presentation/@showSpecialPlsOnTitleSld)">
          <xsl:attribute name="标题幻灯片中是否显示_B64B">true</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="标题幻灯片中是否显示_B64B">false</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

    </规则:幻灯片_B641>
    <!--/xsl:for-each-->
  </xsl:template>
  <!--2014-06-02, tangjiang, 添加幻灯片页眉页脚 end-->
</xsl:stylesheet>