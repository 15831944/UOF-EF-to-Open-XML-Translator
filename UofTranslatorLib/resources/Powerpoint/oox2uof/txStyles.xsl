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
                
                xmlns:uof="http://schemas.uof.org/cn/2003/uof"
                xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow"
                xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc"
                xmlns:图="http://schemas.uof.org/cn/2003/graph"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="PPr-commen.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template match="p:txStyles">
    <!--<xsl:param name="ID" select="//p:sldMaster/@id"/>-->
	  <xsl:param name="ID"/>
	  <式样:文本式样集_9913>
		  <式样:文本式样_9914>
          <xsl:attribute name="标识符_9909">
        <xsl:value-of select="concat('txStyle-',substring-before($ID,'.xml'))"/>
      </xsl:attribute>
      <!--
      2010.4.15 黎美秀修改
      母版中文本式样的转换
      -->
			
      <xsl:apply-templates select="p:titleStyle/*|p:bodyStyle/*|p:otherStyle/*|ancestor::p:sldMaster/p:cSld//a:lstStyle/*"/>
      <!--
      2010.4.15LMX版式中及幻灯片中文本式样的转换
      -->
      <xsl:call-template name="textStyles"/>
    </式样:文本式样_9914>
	  </式样:文本式样集_9913>
  </xsl:template>
  
  <xsl:template name="textStyles">
	 
    <xsl:apply-templates select="//p:defaultTextStyle"/>
    <xsl:apply-templates select="//p:sld//p:sp|//p:sldMaster/p:sldLayout//p:sp" mode="notM-style">
      <xsl:with-param name="sz" select="p:nvSpPr/p:nvPr/p:ph/@sz"/>
    </xsl:apply-templates>
  </xsl:template>
  <xsl:template match="p:sp" mode="notM-style">
    <!--
      2010.4.15 黎美秀添加 <xsl:param name ="sz"/>
      -->
	  
    <xsl:param name="sz"/>
	  
    <xsl:variable name="hID" select="@hID"/>    
    <xsl:apply-templates select=".//a:lstStyle" mode="notM-style">
      <xsl:with-param name="hID" select="$hID"/>
      <!--变量sz用于判断两栏版式-->
      <xsl:with-param name="sz" select="$sz"/>
      <xsl:with-param name="pht" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
    </xsl:apply-templates>
    
  </xsl:template>
  <xsl:template match="p:defaultTextStyle">
	  
	  <xsl:param name="ID"/>
	 
	  <xsl:for-each select="*">
		  <xsl:comment>333</xsl:comment>
      <!--2010-11-10 罗文甜 后继式样引用-->
		<式样:段落式样_9905>
      <!--<演:段落式样 uof:locID="p0135" uof:attrList="标识符 名称 类型 别名 基式样引用 后继式样引用">-->
        <xsl:attribute name="标识符_4100">
          <xsl:value-of select="@id"/>
        </xsl:attribute>
        <xsl:attribute name="名称_4101">shape</xsl:attribute>
        <xsl:attribute name="类型_4102">default</xsl:attribute>
        <xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">
          <字:大纲级别_417C>
            <xsl:choose>
              <!--2011-3-6罗文甜 修改大纲级别加1-->
              <xsl:when test="name(.)='a:lvl1pPr'">1</xsl:when>
              <xsl:when test="name(.)='a:lvl2pPr'">2</xsl:when>
              <xsl:when test="name(.)='a:lvl3pPr'">3</xsl:when>
              <xsl:when test="name(.)='a:lvl4pPr'">4</xsl:when>
              <xsl:when test="name(.)='a:lvl5pPr'">5</xsl:when>
              <xsl:when test="name(.)='a:lvl6pPr'">6</xsl:when>
              <xsl:when test="name(.)='a:lvl7pPr'">7</xsl:when>
              <xsl:when test="name(.)='a:lvl8pPr'">8</xsl:when>
              <xsl:when test="name(.)='a:lvl9pPr'">9</xsl:when>
            </xsl:choose>
          </字:大纲级别_417C>
        </xsl:if>
        <!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
		<xsl:apply-templates select="a:defRPr" mode="pPr-child"/>
        <xsl:call-template name="PPr-commen"/>
       
      <!--</演:段落式样>-->
		</式样:段落式样_9905>
    </xsl:for-each>
  </xsl:template>
  <!--luowentian 母版中的式样-->
  <xsl:template match="p:titleStyle/*|p:bodyStyle/*|p:otherStyle/*|a:lstStyle/*">
	 
	  <式样:段落式样_9905>
    <!--<演:段落式样 uof:locID="p0135" uof:attrList="标识符 名称 类型 别名 基式样引用 后继式样引用">-->
      <xsl:attribute name="标识符_4100">
        <xsl:value-of select="@id"/>
      </xsl:attribute>
      <xsl:attribute name="名称_4101">
        <xsl:choose>
          <!--2011-3-8罗文甜 1.1中母板中的标题的名称都是slide-->
          <xsl:when test="name(..)='p:titleStyle'">slide</xsl:when>
          <xsl:when test="name(..)='p:bodyStyle'">slide</xsl:when>
          <xsl:when test="name(..)='p:otherStyle'">shape</xsl:when>
          <!--2010.4.15 黎美秀添加 a:lstStyle的情况 -->
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='body'">
                <xsl:value-of select="'slide'"/>
              </xsl:when>
              <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='title'">
                <xsl:value-of select="'slide'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'shape'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--母版中文本的式样，考虑基式样引用-->
      <xsl:if test="(name(..)='a:lstStyle') and (ancestor::p:sldMaster) and not(ancestor::p:sldLayout)">
        <xsl:variable name="pPr-lvl" select="name(.)"/>
        <!--2010.4.15 黎美秀修改 需要根据占位符类型判断引用-->   
        <xsl:variable name="styleInMaster">
          <xsl:choose>
            <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='title'">
              <xsl:choose>
                <xsl:when test="$pPr-lvl='a:lvl1pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl2pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl3pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl4pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl4pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl5pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl5pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl6pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl6pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl7pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl7pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl8pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl8pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl9pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl9pPr/@id"/>
                </xsl:when>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='body'">
              <xsl:choose>
                <xsl:when test="$pPr-lvl='a:lvl1pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl2pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl3pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl4pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl5pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl6pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl7pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl8pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl9pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/@id"/>
                </xsl:when>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$pPr-lvl='a:lvl1pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl2pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl2pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl3pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl4pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl5pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl5pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl6pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl6pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl7pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl7pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl8pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl8pPr/@id"/>
                </xsl:when>
                <xsl:when test="$pPr-lvl='a:lvl9pPr'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl9pPr/@id"/>
                </xsl:when>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!--        
        2010.4.15 黎美秀修改
        <xsl:if test ="$styleInMaster!=''">
            <xsl:attribute name="字:基式样引用">			
				  </xsl:attribute>
        -->
        <xsl:if test="$styleInMaster!=''">
          <xsl:attribute name="基式样引用_4104">
            <xsl:value-of select="$styleInMaster"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
      <xsl:if test="name(.)='a:defPPr'">
        <xsl:attribute name="类型_4102">default</xsl:attribute>
      </xsl:if>
      <xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">
        <xsl:attribute name="类型_4102">custom</xsl:attribute>
      </xsl:if>
      <!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
      <xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">
        <xsl:choose>
          <!--20110122罗文甜，修改大纲级别bug，页眉页脚编号p:otherStyle-->
          <xsl:when test="name(..)='p:titleStyle' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='title' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle'">
          
			  <字:大纲级别_417C>
              <xsl:choose>
                <xsl:when test="name(.)='a:lvl1pPr'">0</xsl:when>
                <xsl:when test="name(.)='a:lvl2pPr'">1</xsl:when>
                <xsl:when test="name(.)='a:lvl3pPr'">2</xsl:when>
                <xsl:when test="name(.)='a:lvl4pPr'">3</xsl:when>
                <xsl:when test="name(.)='a:lvl5pPr'">4</xsl:when>
                <xsl:when test="name(.)='a:lvl6pPr'">5</xsl:when>
                <xsl:when test="name(.)='a:lvl7pPr'">6</xsl:when>
                <xsl:when test="name(.)='a:lvl8pPr'">7</xsl:when>
                <xsl:when test="name(.)='a:lvl9pPr'">8</xsl:when>
              </xsl:choose>
            </字:大纲级别_417C>
          </xsl:when>     
          <xsl:otherwise>
			  <字:大纲级别_417C>
              <xsl:choose>
                <xsl:when test="name(.)='a:lvl1pPr'">1</xsl:when>
                <xsl:when test="name(.)='a:lvl2pPr'">2</xsl:when>
                <xsl:when test="name(.)='a:lvl3pPr'">3</xsl:when>
                <xsl:when test="name(.)='a:lvl4pPr'">4</xsl:when>
                <xsl:when test="name(.)='a:lvl5pPr'">5</xsl:when>
                <xsl:when test="name(.)='a:lvl6pPr'">6</xsl:when>
                <xsl:when test="name(.)='a:lvl7pPr'">7</xsl:when>
                <xsl:when test="name(.)='a:lvl8pPr'">8</xsl:when>
                <xsl:when test="name(.)='a:lvl9pPr'">9</xsl:when>
              </xsl:choose>
            </字:大纲级别_417C>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
     
      <xsl:apply-templates select="a:defRPr" mode="pPr-child"/>
		  <xsl:call-template name="PPr-commen"/>
	  </式样:段落式样_9905>
    <!--</演:段落式样>-->
  </xsl:template>
  <xsl:template match="a:lstStyle" mode="notM-style">
    <!--    
    2010.4.15 黎美秀增加变量sz 处理两栏内容的特殊情况
     <xsl:param name ="sz"/>
    -->
    <xsl:param name="pht"/>
    <xsl:param name="hID"/>    
    <xsl:param name="sz"/>
    <xsl:variable name="name">
      <xsl:choose>
        <!--2010.04.10 马有旭 修改<xsl:when test="$pht='title'">title</xsl:when>-->
        <!--2010.4.15黎美秀 修改<xsl:when test="$pht='title' or $pht='ctrTitle'">title</xsl:when>-->
        <xsl:when test="$pht='title' or $pht='ctrTitle' or $pht='subTitle' ">title</xsl:when>
        <xsl:when test="$pht='body' or $hID='obj' or $hID=''">slide</xsl:when>
        <xsl:otherwise>shape</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
     <xsl:for-each select="*">
      <xsl:variable name="style" select="name()"/>
		 <式样:段落式样_9905>
      <!--<演:段落式样 uof:locID="p0135" uof:attrList="标识符 名称 类型 别名 基式样引用 后继式样引用">-->
  
        <xsl:attribute name="标识符_4100">
          <xsl:value-of select="@id"/>
        </xsl:attribute>
        <xsl:attribute name="名称_4101">
          <!--2010.04.10 马有旭 修改<xsl:value-of select ="'shape'"/>-->
          <!--xsl:value-of select="$type"/-->

          <!--2014-04-14, tangjiang, 修复标题多出项目符号 start  -->
          <!--<xsl:value-of select="$name"/>-->
          <xsl:value-of select ="'shape'"/>
          <!--end 2014-04-14, tangjiang, 修复标题多出项目符号  -->

        </xsl:attribute>
        <xsl:attribute name="类型_4102">custom</xsl:attribute>

       <!--2014-03-27, tangjiang, 添加没有对应的则默认第一个段落的样式 start -->
       <!--找到被继承图形中大纲级别相同的段落所应用的式样,若没有对应的则默认第一个段落的样式-->
       <xsl:variable name="styleID">
         <xsl:if test="$hID!=''">
           <xsl:variable name="lvlValue" select="ancestor::p:sldMaster//p:sp[@id = $hID]/p:txBody/a:p/a:pPr/@lvl"/>
           <xsl:choose>
             <xsl:when test="$lvlValue!=''">
               <xsl:value-of select="ancestor::p:sldMaster//p:sp[@id = $hID]/p:txBody/a:p[string(a:pPr/@lvl + 1)=substring-before(substring-after($style,'lvl'),'pPr')]/@styleID"/>
             </xsl:when>
             <xsl:otherwise>
               <xsl:value-of select="ancestor::p:sldMaster//p:sp[@id = $hID]/p:txBody/a:p/@styleID"/>
             </xsl:otherwise>
           </xsl:choose>
         </xsl:if>
       </xsl:variable>
       <!-- end 2014-03-27, tangjiang, 添加没有对应的则默认第一个段落的样式 -->
       
        <xsl:if test="$styleID!=''">
          <xsl:attribute name="基式样引用_4104">
            <xsl:value-of select="$styleID"/>
          </xsl:attribute>
        </xsl:if>
        <!--被继承图形中大纲级别相同的段落不存在 或 不具有被继承图形 这时要从母版的textStyles中去找-->
        <xsl:if test="$styleID=''">
          <xsl:variable name="styleID2">
            <xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph">
              <xsl:choose>
                <xsl:when test="$pht='title'">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/*[name()=$style]/@id"/>
                </xsl:when>
                <xsl:when test="$pht='body' or $pht='obj' or not($pht)">
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/*[name()=$style]/@id"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/*[name()=$style]/@id"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </xsl:variable>
          <xsl:if test="$styleID2!=''">
            <xsl:attribute name="基式样引用_4104">
              <xsl:value-of select="$styleID2"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
        <!--xsl:if test="name(.)='a:defPPr'">
          <xsl:attribute name="字:类型">default</xsl:attribute>
        </xsl:if>
        <xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">
          <xsl:attribute name="字:类型">custom</xsl:attribute>
        </xsl:if-->
        <xsl:if test="name(.)!='a:defPPr' and name(.)!='a:extLst'">

          <xsl:choose>
           <!--20110222罗文甜，增加判断（页眉页脚，时间）-->
           <!--2011-2-22罗：当占位符是title或者centertitle时，uof大纲级别是从0开始-->
            <xsl:when test="$pht='ctrTitle' or $pht='title'">
            <!--xsl:when test="$pht='ctrTitle' or $pht='title' or $pht='dt' or $pht='ftr' or $pht='sldNum' or $pht='hdr'"-->
              <字:大纲级别_417C>
                <xsl:choose>
                  <xsl:when test="name(.)='a:lvl1pPr'">0</xsl:when>
                  <xsl:when test="name(.)='a:lvl2pPr'">1</xsl:when>
                  <xsl:when test="name(.)='a:lvl3pPr'">2</xsl:when>
                  <xsl:when test="name(.)='a:lvl4pPr'">3</xsl:when>
                  <xsl:when test="name(.)='a:lvl5pPr'">4</xsl:when>
                  <xsl:when test="name(.)='a:lvl6pPr'">5</xsl:when>
                  <xsl:when test="name(.)='a:lvl7pPr'">6</xsl:when>
                  <xsl:when test="name(.)='a:lvl8pPr'">7</xsl:when>
                  <xsl:when test="name(.)='a:lvl9pPr'">8</xsl:when>
                </xsl:choose>
              </字:大纲级别_417C>
            </xsl:when>           
            <xsl:otherwise>
				<字:大纲级别_417C>
                <xsl:choose>
                  <xsl:when test="name(.)='a:lvl1pPr'">1</xsl:when>
                  <xsl:when test="name(.)='a:lvl2pPr'">2</xsl:when>
                  <xsl:when test="name(.)='a:lvl3pPr'">3</xsl:when>
                  <xsl:when test="name(.)='a:lvl4pPr'">4</xsl:when>
                  <xsl:when test="name(.)='a:lvl5pPr'">5</xsl:when>
                  <xsl:when test="name(.)='a:lvl6pPr'">6</xsl:when>
                  <xsl:when test="name(.)='a:lvl7pPr'">7</xsl:when>
                  <xsl:when test="name(.)='a:lvl8pPr'">8</xsl:when>
                  <xsl:when test="name(.)='a:lvl9pPr'">9</xsl:when>
                </xsl:choose>                
              </字:大纲级别_417C>
              </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
        <xsl:call-template name="PPr-commen"/>
        <xsl:apply-templates select="a:defRPr" mode="pPr-child"/>
      <!--</演:段落式样>-->
		 </式样:段落式样_9905>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
