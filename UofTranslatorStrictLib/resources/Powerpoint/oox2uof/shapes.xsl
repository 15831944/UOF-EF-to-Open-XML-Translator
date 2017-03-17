<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:fo="http://www.w3.org/1999/XSL/Format" 
                xmlns:dc="http://purl.org/dc/elements/1.1/" 
                xmlns:dcterms="http://purl.org/dc/terms/" 
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                
                xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" 
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
                xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
                xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                
                xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph">
  <xsl:import href="p.xsl"/>
  <xsl:import href="fill.xsl"/>
  <xsl:import href="table.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  
  <!-- 10.02.03 马有旭 spTree部分重新调整 -->
	<!--11.11.10 李娟-->
  <xsl:template match="p:spTree" mode="sp">
    <xsl:variable name="type" select="name(../..)"/>
    <!-- 10.黎美秀03.9  增加文字表部分处理-->    
   <!--<xsl:for-each select="p:cxnSp|p:grpSp|p:pic|p:sp|p:graphicFrame">-->
      <!--liuyangyang 2014-12-22 添加对公式的处理-->
   <xsl:for-each select="p:cxnSp|p:grpSp|p:pic|p:sp|p:graphicFrame|mc:AlternateContent">
     <xsl:if test="name(.) = 'mc:AlternateContent'">
       <图:图形_8062>
         <xsl:attribute name="标识符_804B">
           <xsl:value-of select="@id"/>
         </xsl:attribute>
         <图:预定义图形_8018>
           <图:类别_8019>11</图:类别_8019>
           <图:名称_801A>Rectangle</图:名称_801A>
           <图:生成软件_801B>Yozo Office</图:生成软件_801B>
           <图:属性_801D>
             <图:旋转角度_804D>0.0</图:旋转角度_804D>
             <图:是否打印对象_804E>true</图:是否打印对象_804E>
             <图:缩放是否锁定纵横比_8055>true</图:缩放是否锁定纵横比_8055>
             <xsl:if test="./mc:Choice/p:sp/p:spPr/a:xfrm/a:ext">
               <图:大小_8060>
                 <xsl:attribute name="宽_C605">
                   <xsl:value-of select="./mc:Choice/p:sp/p:spPr/a:xfrm/a:ext/@cx div 12700"/>
                 </xsl:attribute>
                 <xsl:attribute name="长_C604">
                   <xsl:value-of select="./mc:Choice/p:sp/p:spPr/a:xfrm/a:ext/@cy div 12700"/>
                 </xsl:attribute>
               </图:大小_8060>
             </xsl:if>
             <图:图片属性_801E>
               <图:颜色模式_801F>auto</图:颜色模式_801F>
               <图:亮度_8020>50.0</图:亮度_8020>
               <图:对比度_8021>50.0</图:对比度_8021>
             </图:图片属性_801E>
           </图:属性_801D>
         </图:预定义图形_8018>
         <图:图片数据引用_8037>
           <xsl:value-of select="concat(substring-before(../../../@id,'.xml'),.//@r:embed)"/>
         </图:图片数据引用_8037>
         <图:其他对象引用_8038>
           <xsl:value-of select="@id"/>
         </图:其他对象引用_8038>
       </图:图形_8062>
       <!--<xsl:variable name="id" select="@id"/>
        <xsl:for-each select="./mc:Fallback/p:sp">
          <xsl:apply-templates select=".">
            <xsl:with-param name="type" select="$type"/>
            <xsl:with-param name="id" select="$id"/>
          </xsl:apply-templates>
        </xsl:for-each>-->
     </xsl:if>
     <!--end -->
      <xsl:apply-templates select=".">
		<xsl:with-param name="picFrom" select="'slides'"/>
        <xsl:with-param name="type" select="$type"/>
      </xsl:apply-templates>
    </xsl:for-each>
    <!--guoyongbin 2013-12-30 修改图形批注引用-->
    <xsl:if test="//p:cmAuthorLst and ../../@commentID">
      <xsl:variable name="commentID" select="../../@commentID"/>
      <xsl:variable name="cmSlideID" select="../../@id"/>
      <图:图形_8062>
        <xsl:attribute name="标识符_804B">
          <xsl:value-of select="substring-before($commentID,'.xml')"/>
        </xsl:attribute>
        <图:预定义图形_8018>
          <图:类别_8019>11</图:类别_8019>
          <图:名称_801A>Rectangle</图:名称_801A>
          <图:生成软件_801B>Yozo Office</图:生成软件_801B>
          <图:属性_801D>
            <图:旋转角度_804D>0.0</图:旋转角度_804D>
            <图:是否打印对象_804E>true</图:是否打印对象_804E>
            <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
            <图:线_8057>
              <图:线粗细_805C>0.75</图:线粗细_805C>
              <图:线类型_8059 线型_805A="single" 虚实_805B="solid"/>
            </图:线_8057>
            <图:大小_8060 长_C604="35.699997" 宽_C605="173.5688"/>
          </图:属性_801D>
        </图:预定义图形_8018>
        <图:文本_803C 是否为文本框_8046="true" 是否自动换行_8047="true" 是否大小适应文字_8048="true" 是否文字随对象旋转_8049="true" 是否文字随边框自动缩放_804A="true" 是否文字绕排外部对象_8068="true">
          <图:边距_803D 左_C608="7.2" 上_C609="3.6" 右_C60A="7.2" 下_C60B="3.6"/>
          <图:对齐_803E 水平对齐_421D="left" 文字对齐_421E="top"/>
          <图:文字排列方向_8042>t2b-l2r-0e-0w</图:文字排列方向_8042>
          <图:内容_8043>
            <字:段落_416B 标识符_4169="Obj05208p0">

              <!--<字:句_419D>
							  <字:句属性_4158>
								  <字:字体_4128 颜色_412F="#000000"/>
							  </字:句属性_4158>
							  <字:文本串_415B>55</字:文本串_415B>
						  </字:句_419D>
						  <字:句_419D>
							  <字:区域开始_4165 标识符_4100="cmt_0" 类型_413B="annotation"/>
							  <字:区域结束_4167 标识符引用_4168="cmt_0"/>
						  </字:句_419D>-->
              <xsl:for-each select="//p:cmLst[@id=$commentID]/p:cm">
							  <字:句_419D>
								  <字:区域开始_4165>
									  <xsl:variable name="idx">
										  <xsl:value-of select ="./@idx"/>
									  </xsl:variable>
									  <xsl:attribute name="标识符_4100">
                      <xsl:value-of select="concat('cmt_',substring-before($cmSlideID,'.xml'),'_',$idx)"/>

									  </xsl:attribute>
									  <xsl:attribute name="类型_413B">annotation</xsl:attribute>
								  </字:区域开始_4165>
               
                  </字:句_419D>
                  <字:句_419D>
                  <字:区域结束_4167>
									  <xsl:attribute name="标识符引用_4168">

										  <xsl:variable name="idx">
											  <xsl:value-of select ="./@idx"/>
										  </xsl:variable>
                      <xsl:value-of select="concat('cmt_',substring-before($cmSlideID,'.xml'),'_',$idx)"/>
									  </xsl:attribute>
								  </字:区域结束_4167>
               
							  </字:句_419D>
						  </xsl:for-each>
            </字:段落_416B>
          </图:内容_8043>
        </图:文本_803C>
      </图:图形_8062>
    </xsl:if>
    <!--enf -->
	  <!--添加对批注的图形引用 李娟-->
    <!--<xsl:if test="//p:cmAuthorLst and //p:cmLst">
      <图:图形_8062>
        <xsl:attribute name="标识符_804B">
          <xsl:value-of select="substring-before(//p:cmLst/@id,'.xml')"/>
        </xsl:attribute>
        <图:预定义图形_8018>
          <图:类别_8019>11</图:类别_8019>
          <图:名称_801A>Rectanglewtf123</图:名称_801A>
          <图:生成软件_801B>Yozo Office</图:生成软件_801B>
          <图:属性_801D>
            <图:旋转角度_804D>0.0</图:旋转角度_804D>
            <图:是否打印对象_804E>true</图:是否打印对象_804E>
            <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
            <图:线_8057>
              <图:线粗细_805C>0.75</图:线粗细_805C>
              <图:线类型_8059 线型_805A="single" 虚实_805B="solid"/>
            </图:线_8057>
            <图:大小_8060 长_C604="35.699997" 宽_C605="173.5688"/>
          </图:属性_801D>
        </图:预定义图形_8018>
        <图:文本_803C 是否为文本框_8046="true" 是否自动换行_8047="true" 是否大小适应文字_8048="true" 是否文字随对象旋转_8049="true" 是否文字随边框自动缩放_804A="true" 是否文字绕排外部对象_8068="true">
          <图:边距_803D 左_C608="7.2" 上_C609="3.6" 右_C60A="7.2" 下_C60B="3.6"/>
          <图:对齐_803E 水平对齐_421D="left" 文字对齐_421E="top"/>
          <图:文字排列方向_8042>t2b-l2r-0e-0w</图:文字排列方向_8042>
          <图:内容_8043>
            <字:段落_416B 标识符_4169="Obj05208p0">

              <字:句_419D>
							  <字:句属性_4158>
								  <字:字体_4128 颜色_412F="#000000"/>
							  </字:句属性_4158>
							  <字:文本串_415B>55</字:文本串_415B>
						  </字:句_419D>
						  <字:句_419D>
							  <字:区域开始_4165 标识符_4100="cmt_0" 类型_413B="annotation"/>
							  <字:区域结束_4167 标识符引用_4168="cmt_0"/>
						  </字:句_419D>
              <xsl:for-each select="//p:cmLst/p:cm">
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
						  </xsl:for-each>
            </字:段落_416B>
          </图:内容_8043>
        </图:文本_803C>
      </图:图形_8062>
    </xsl:if>-->
  </xsl:template>
  <!--liuyagyang 2014-12-22 添加公式转换-->
  <!--<xsl:template match="mc:AlternateContent">
   <xsl:param name="picFrom"/>
    <xsl:param name="type"/>
    
    <xsl:for-each select="./mc:Fallback/p:sp">
      <xsl:apply-templates select=".">
        <xsl:with-param name="picFrom" select="$picFrom"/>
        <xsl:with-param name="type" select="$type"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>-->
  <!--end-->

  <!--2013-1-11 新增tempate 解决文字表和smartart调用冲突的问题 均用name调用相应模板 增加图表  liqiuling 2013-03-26  start-->
  <xsl:template match="p:graphicFrame">
    <xsl:param name="type"/>
    <xsl:variable name="graphicdataType">
      <xsl:value-of select="descendant::a:graphicData/@uri"/>
    </xsl:variable>
    <xsl:if test="contains($graphicdataType,'diagram')">
      <xsl:call-template name="smartart">
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test="contains($graphicdataType,'table')">
      <xsl:call-template name="table">
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test="contains($graphicdataType,'chart')">
      <xsl:value-of select="$type"/>
      <xsl:call-template name="chart">
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test="contains($graphicdataType,'oleObject')">
      <xsl:call-template name="oleObject">
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  <!--2013-1-11 新增tempate 解决文字表和smartart调用冲突的问题 均用name调用相应模板 增加图表  liqiuling 2013-03-26  end-->


  <!--2014-03-17, tangjiang, 新增OLE对象的转换,只将图片数据拷贝过来 start -->
  <xsl:template name="oleObject">
    <xsl:param name="type"/>
    <xsl:variable name="refslide">
      <xsl:value-of select ="ancestor::p:sld/@id"/>
    </xsl:variable>
    <xsl:variable name="idofralationship">
      <xsl:value-of select="concat($refslide,'.rels')"/>
    </xsl:variable>
    <xsl:variable name="imageID">
      <xsl:value-of select="./a:graphic/a:graphicData/mc:AlternateContent/mc:Fallback/p:oleObj/p:pic/p:blipFill/a:blip/@r:embed"/>
    </xsl:variable>

    <图:图形_8062>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="translate(@id,':','')"/>
      </xsl:attribute>
      <!--liuyangyang添加ole对象标记以及额外内容供互操作转换-->
      <!--liuyangyang 部分ole对象存在问题，暂不转换-->
      <!--<xsl:if test="./a:graphic/a:graphicData/mc:AlternateContent/mc:Choice/p:oleObj/@name != 'Visio'">-->
      <ole>
        <xsl:copy-of select="."/>
        <xsl:variable name="rid" select="./a:graphic/a:graphicData/mc:AlternateContent/mc:Fallback/p:oleObj/@r:id"/>
        <xsl:variable name ="id" select="concat(../../../@id,'.rels')"/>
        <olerel>
        <xsl:copy-of select="ancestor::p:presentation/rel:Relationships[@id=$id]/rel:Relationship[@Id=$rid]"/>
          <xsl:copy-of select="ancestor::p:presentation/rel:Relationships[@id=$id]/rel:Relationship[@Id=$imageID]"/>
        </olerel>
      </ole>
      <!--</xsl:if>-->
      <!--end-->
      <图:预定义图形_8018>
        <图:类别_8019>11</图:类别_8019>
        <图:名称_801A>Rectangle</图:名称_801A>
        <图:生成软件_801B>Yozo Office</图:生成软件_801B>
        <图:属性_801D>
          <图:旋转角度_804D>0.0</图:旋转角度_804D>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
          <图:缩放是否锁定纵横比_8055>true</图:缩放是否锁定纵横比_8055>
          <xsl:if test="a:xfrm/a:ext">
            <图:大小_8060>
              <xsl:attribute name="宽_C605">
                <xsl:value-of select="a:xfrm/a:ext/@cx div 12700"/>
              </xsl:attribute>
              <xsl:attribute name="长_C604">
                <xsl:value-of select="a:xfrm/a:ext/@cy div 12700"/>
              </xsl:attribute>
            </图:大小_8060>
          </xsl:if>
          <图:图片属性_801E>
            <图:颜色模式_801F>auto</图:颜色模式_801F>
            <图:亮度_8020>50.0</图:亮度_8020>
            <图:对比度_8021>50.0</图:对比度_8021>
          </图:图片属性_801E>
        </图:属性_801D>
      </图:预定义图形_8018>

      <xsl:for-each select="//rel:Relationships[@id=$idofralationship]/rel:Relationship[@Id=$imageID]">
      <!--<xsl:for-each select="//rel:Relationships[@id=$idofralationship]/rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']">-->
        <xsl:element name ="图:图片数据引用_8037">
          <xsl:value-of select="concat(substring-before($refslide,'.xml'),./@Id)"/>
        </xsl:element>
      </xsl:for-each>

    </图:图形_8062>
  </xsl:template>
  <!-- end 2014-03-17, tangjiang, 新增OLE对象的转换,只将图片数据拷贝过来 -->
  
  <!--2012-12-20, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
  <xsl:template  name="smartart">
    <xsl:param name="type"/>
    <xsl:variable name="refslide" select ="ancestor::p:sld/@id"/>
    <xsl:variable name="relationshipId" select="concat($refslide,'.rels')"/>
    <xsl:variable name="drawingId" select="concat('rId', substring-after(descendant::a:graphicData/dgm:relIds/@r:cs,'rId') + 1)"/>
    <xsl:variable name="drawingFileId" select="substring-after(ancestor::p:presentation//rel:Relationships[@id=$relationshipId]//rel:Relationship[@Id=$drawingId]/@Target, '../diagrams/')"/>
    <xsl:for-each select="//dsp:drawing[@refBy=$refslide and @id=$drawingFileId]//dsp:sp">
        <xsl:apply-templates select=".">
          <xsl:with-param name ="type" select="$type"/>
        </xsl:apply-templates>
    </xsl:for-each>
    <!--liuyangyang 2014-12-11 注释掉对smartart组合图形整体锚点转换-->
       <!-- <图:图形_8062>
          <xsl:attribute name="标识符_804B">
            <xsl:value-of select="translate(@id,':','')"/>
          </xsl:attribute>
          <xsl:attribute name="组合列表_8064">
            <xsl:for-each select="//dsp:drawing[@refBy=$refslide and @id=$drawingFileId]//dsp:sp">
              <xsl:value-of select="concat(' ',translate(@id,':',''))"/>
            </xsl:for-each>
          </xsl:attribute>
        </图:图形_8062>-->
    <!--end-->
  </xsl:template>
  <!--end -->
  
  <!--增加图表  liqiuling 2013-03-26 start-->
  <xsl:template  name="chart">
    <xsl:param name="type"/>
    <xsl:variable name="refslide">
      <xsl:value-of select ="ancestor::p:sld/@id"/>
    </xsl:variable>
    <xsl:variable name="idofralationship">
      <xsl:value-of select="concat($refslide,'.rels')"/>
    </xsl:variable>
      <图:图形_8062>
        <xsl:attribute name="标识符_804B">
          <xsl:value-of select="translate(@id,':','')"/>
        </xsl:attribute>
        <图:预定义图形_8018>
          <图:类别_8019>11</图:类别_8019>
          <图:名称_801A>Rectangle</图:名称_801A>
          <图:生成软件_801B>Yozo Office</图:生成软件_801B>
          <图:属性_801D>
            <图:旋转角度_804D>0.0</图:旋转角度_804D>
            <图:是否打印对象_804E>true</图:是否打印对象_804E>
            <图:缩放是否锁定纵横比_8055>true</图:缩放是否锁定纵横比_8055>
            <xsl:choose>
              <xsl:when test ="a:xfrm/a:ext">
                <图:大小_8060>
                  <xsl:attribute name="宽_C605">
                    <xsl:value-of select="a:xfrm/a:ext/@cx div 12700"/>

                  </xsl:attribute>

                  <xsl:attribute name="长_C604">
                    <xsl:value-of select="a:xfrm/a:ext/@cy div 12700"/>

                  </xsl:attribute>
                </图:大小_8060>
              </xsl:when>
              <xsl:when test ="p:xfrm/a:ext">
                <图:大小_8060>
                  <xsl:attribute name="宽_C605">
                    <xsl:value-of select="p:xfrm/a:ext/@cx div 12700"/>

                  </xsl:attribute>

                  <xsl:attribute name="长_C604">
                    <xsl:value-of select="p:xfrm/a:ext/@cy div 12700"/>

                  </xsl:attribute>
                </图:大小_8060>
              </xsl:when>
            </xsl:choose>
            <图:图片属性_801E>
              <图:颜色模式_801F>auto</图:颜色模式_801F>
              <图:亮度_8020>50.0</图:亮度_8020>
              <图:对比度_8021>50.0</图:对比度_8021>
            </图:图片属性_801E>
          </图:属性_801D>
        </图:预定义图形_8018>
        
        <!--2013-11-20, tangjiang, 修复Strict OOXML到UOF图表的转换 start -->
        <xsl:for-each select="//rel:Relationships[@id=$idofralationship]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/chart']">
          <xsl:element name ="图:图片数据引用_8037">
            <xsl:value-of select="substring-before(substring-after(@Target,'../charts/'),'.xml')"/>
          </xsl:element>
        </xsl:for-each>
        <!-- end 2013-11-20, tangjiang -->
        
      </图:图形_8062>
  </xsl:template>
  <!--增加图表  liqiuling 2013-03-26 end-->
  

  <!--2012-12-20, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
  <xsl:template match="p:sp|p:cxnSp|dsp:sp">
    <xsl:param name="type"/>
    <xsl:param name="id"/>
    <!--2010-11-10罗文甜：增加OLE数据-->
	  <!--层次 和组合列表没有添加 李娟 11.11.10-->
	   <图:图形_8062>
         
           <xsl:attribute name="标识符_804B">
             <!--liuyangyang 2014-12-22 修改公式转换id-->
             <xsl:choose>
               <xsl:when test="@id">
                 <xsl:value-of select="translate(@id,':','')"/>
               </xsl:when>
               <xsl:otherwise>
                 <xsl:value-of select="translate($id,':','')"/>
               </xsl:otherwise>
             </xsl:choose>
             <!--<xsl:value-of select="p:nvSpPr/p:cNvPr/@id"/>-->
            <!--<xsl:value-of select="translate(@id,':','')"/>-->
             <!--end-->
           </xsl:attribute>
     
			  <xsl:apply-templates select="p:spPr|dsp:spPr"/>
			  <xsl:if test="$type!='p:notesMaster' or p:nvSpPr/p:nvPr/p:ph/@type!='sldImg'">
				  <xsl:apply-templates select="p:txBody|dsp:txBody"/>
			  </xsl:if>
 
			  <!--2011-5-26罗文甜，修改翻转bug-->
			  <xsl:choose>
				  <xsl:when test="p:spPr/a:xfrm/@flipV and not(p:spPr/a:xfrm/@flipH)">
					  <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="y"/>-->
					  <图:翻转_803A>y</图:翻转_803A>
				  </xsl:when>
				  <xsl:when test="p:spPr/a:xfrm/@flipH and not(p:spPr/a:xfrm/@flipV)">
					  <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="x"/>-->
					  <图:翻转_803A>x</图:翻转_803A>
				  </xsl:when>
				  <xsl:when test="p:spPr/a:xfrm/@flipH and p:spPr/a:xfrm/@flipV">
					  <图:翻转_803A>xy</图:翻转_803A>
					  <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="xy"/>-->
				  </xsl:when>
          <!--liuyangyang，2014-12-08，添加leftCircularArrow翻转为CircularArrow-->
          <xsl:when test=".//a:prstGeom/@prst ='leftCircularArrow'">
            <图:翻转_803A>x</图:翻转_803A>
          </xsl:when>
          <!--end-->
				  <xsl:otherwise>
				  </xsl:otherwise>
			  </xsl:choose>
			  <!--10.02.02马有旭<xsl:if test ="ancestor::p:grpSp">-->
			  <xsl:if test="parent::p:grpSp">
				  <xsl:call-template name="groupLocation">
				  </xsl:call-template>
			  </xsl:if>
       <!--2012-12-18, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
       <xsl:if test="ancestor::dsp:drawing">
         <xsl:variable name="refdsp">
           <xsl:value-of select ="ancestor::dsp:drawing/@refBy"/>
         </xsl:variable>
         <xsl:call-template name="dspgroupLocation">
           <xsl:with-param name="framexfrm" select="//p:sld[@id=$refdsp]//p:graphicFrame/p:xfrm/a:off"/>
         </xsl:call-template>
       </xsl:if>
       <!--end-->
      <!--liuyangyang 2013-12-02 添加控制点转换-->
       <xsl:choose>
         <!--椭圆形标注，圆角矩形标注，矩形标注，云形标注-->
         <xsl:when test ="(.//a:prstGeom/@prst='wedgeEllipseCallout' or .//a:prstGeom/@prst='wedgeRoundRectCallout' or .//a:prstGeom/@prst='wedgeRectCallout' or .//a:prstGeom/@prst='cloudCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x * 21600) div 100000 + 10800"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y * 21600) div 100000 + 10800"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--线形标注1，线形标注1（带强调线），线形标注1（无边框线），线形标注1（带边框和强调线）-->
         <xsl:when test ="(.//a:prstGeom/@prst='borderCallout1' or .//a:prstGeom/@prst='accentCallout1' or .//a:prstGeom/@prst='callout1' or .//a:prstGeom/@prst='accentBorderCallout1') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y2 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--线形标注2，线形标注2（带强调线），线形标注2（无边框线），线形标注2（带边框和强调线）-->
         <xsl:when test="(.//a:prstGeom/@prst='borderCallout2' or .//a:prstGeom/@prst='accentCallout2' or .//a:prstGeom/@prst='callout2' or .//a:prstGeom/@prst='accentBorderCallout2') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj5">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj5']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj6">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj6']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp5">
             <xsl:value-of select="normalize-space(substring-after($adj5,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp6">
             <xsl:value-of select="normalize-space(substring-after($adj6,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="x3">
             <xsl:value-of select="number($tmp6)"/>
           </xsl:variable>
           <xsl:variable name="y3">
             <xsl:value-of select="number($tmp5)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x3 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y3 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y2 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--线形标注3，线形标注3（带强调线），线形标注3（无边框线），线形标注3（带边框和强调线）-->
         <xsl:when test="(.//a:prstGeom/@prst='borderCallout3' or .//a:prstGeom/@prst='accentCallout3' or .//a:prstGeom/@prst='callout3' or .//a:prstGeom/@prst='accentBorderCallout3') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj5">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj5']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj6">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj6']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj7">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj7']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj8">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj8']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp5">
             <xsl:value-of select="normalize-space(substring-after($adj5,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp6">
             <xsl:value-of select="normalize-space(substring-after($adj6,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp7">
             <xsl:value-of select="normalize-space(substring-after($adj7,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp8">
             <xsl:value-of select="normalize-space(substring-after($adj8,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="x3">
             <xsl:value-of select="number($tmp6)"/>
           </xsl:variable>
           <xsl:variable name="y3">
             <xsl:value-of select="number($tmp5)"/>
           </xsl:variable>
           <xsl:variable name="x4">
             <xsl:value-of select="number($tmp8)"/>
           </xsl:variable>
           <xsl:variable name="y4">
             <xsl:value-of select="number($tmp7)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x4 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y4 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x3 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y3 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y2 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--十字星，八角星，十六角星，二十四角星，三十二角星-->
         <!--五角星OOX有控制点，UOF无控制点-->
         <!--六角星，七角星转成八角星，十角星，十二角星转成十六角星-->
         <xsl:when test ="(.//a:prstGeom/@prst='star4' or .//a:prstGeom/@prst='star8' or .//a:prstGeom/@prst='star16' or .//a:prstGeom/@prst='star24' or .//a:prstGeom/@prst='star32' or .//a:prstGeom/@prst='star6' or .//a:prstGeom/@prst='star7' or .//a:prstGeom/@prst='star10' or .//a:prstGeom/@prst='star12') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上凸带形-->
         <xsl:when test ="(.//a:prstGeom/@prst='ribbon2') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="21600 - ($y * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--前凸带形-->
         <xsl:when test ="(.//a:prstGeom/@prst='ribbon') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上凸弯带形-->
         <xsl:when test ="(.//a:prstGeom/@prst='ellipseRibbon2') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:choose>
               <xsl:when test="number($tmp2) = 100000">
                 <xsl:value-of select="number($tmp2) - 25000"/>
               </xsl:when>
               <xsl:when test="number($tmp2) = 25000">
                 <xsl:value-of select="number($tmp2) + 4"/>
               </xsl:when>
               <xsl:otherwise>
                 <xsl:value-of select ="number($tmp2)"/>
               </xsl:otherwise>
             </xsl:choose>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="21600 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600) div 100000 + 675"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--前凸弯带形-->
         <xsl:when test ="(.//a:prstGeom/@prst='ellipseRibbon') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select ="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x2 * 21600) div 100000 - 675"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--竖卷形，横卷形-->
         <xsl:when test ="(.//a:prstGeom/@prst='verticalScroll' or .//a:prstGeom/@prst='horizontalScroll') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--波形，双波形-->
         <xsl:when test ="(.//a:prstGeom/@prst='wave' or .//a:prstGeom/@prst='doubleWave') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 + ($y * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--右箭头，虚尾箭头，燕尾形箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='rightArrow' or .//a:prstGeom/@prst='stripedRightArrow' or .//a:prstGeom/@prst='notchedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x * 21600 * $height) div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y * 10800) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左箭头，左右箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='leftArrow' or .//a:prstGeom/@prst='leftRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x * 21600 * $height) div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y * 10800) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='upArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($y * 21600 * $width) div $height div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上下箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='upDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="($y * 21600 * $width) div $height div 100000 "/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--下箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='downArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($y * 21600 * $width) div $height div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--十字箭头，丁字箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='quadArrow' or .//a:prstGeom/@prst='leftRightUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x1 * 10800) div 50000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 10800) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 10800 div 50000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--圆角右箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='bentArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x * $height * 21600) div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="6079 - ($y * $height * 10800) div $width div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--直角双向箭头，直角上箭头-->
         <xsl:when test ="(.//a:prstGeom/@prst='leftUpArrow' or .//a:prstGeom/@prst='bentUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x1 * $height * 21600) div $width div 50000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="21600 - ($x1 * $height * 21600) div $width div 100000 + ($y1 * $height * 21600) div $width div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 10800 div 50000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左弧形箭头，显示效果有差异-->
         <xsl:when test ="(.//a:prstGeom/@prst='curvedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $width *  21600 div $height div 100000"/>-->
               <xsl:value-of select="21600 - $x1 * $width *  21600 div $height div 100000 + 250"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $width * 21600 div $height div 200000"/>-->
               <xsl:value-of select="21600 - ($x1 - $y1) * $width * 21600 div $height div 200000 + 250"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - $x2 * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--右弧形箭头，显示效果有差异-->
         <xsl:when test ="(.//a:prstGeom/@prst='curvedLeftArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $width *  21600 div $height div 100000"/>-->
               <xsl:value-of select="21600 - $x1 * $width *  21600 div $height div 100000 + 250"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $width * 21600 div $height div 200000"/>-->
               <xsl:value-of select="21600 - ($x1 - $y1) * $width * 21600 div $height div 200000 + 250"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--下弧形箭头，显示效果有差异-->
         <xsl:when test ="(.//a:prstGeom/@prst='curvedUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $height *  21600 div $width div 100000"/>-->
               <xsl:value-of select="21600 - $x1 * $height *  21600 div $width div 100000 + 10"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $height * 21600 div $width div 200000"/>-->
               <xsl:value-of select="21600 - ($x1 - $y1) * $height * 21600 div $width div 200000 + 10"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上弧形箭头，显示效果有差异-->
         <xsl:when test ="(.//a:prstGeom/@prst='curvedDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $height *  21600 div $width div 100000"/>-->
               <xsl:value-of select="21600 - $x1 * $height *  21600 div $width div 100000 + 10"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $height * 21600 div $width div 200000"/>-->
               <xsl:value-of select="21600 - ($x1 - $y1) * $height * 21600 div $width div 200000 + 10"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - $x2 * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--五边形，燕尾形-->
         <xsl:when test ="(.//a:prstGeom/@prst='homePlate' or .//a:prstGeom/@prst='chevron') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - $x * 21600 * $height div $width div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--右箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='rightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x2 * 21600 * $height) div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='leftArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600 * $height) div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--上箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='upArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x2 * 21600 * $width) div $height div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--下箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='downArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="($x1 * 21600) div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - ($x2 * 21600 * $width) div $height div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左右箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='leftRightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 21600 * $height div $width div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--十字箭头标注-->
         <xsl:when test="(.//a:prstGeom/@prst='quadArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <xsl:variable name="y1">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="y2">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y1 * 21600 * $height) div $width div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x2 * 21600 div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="10800 - ($y2 * 21600 * $height) div $width div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--等腰三角形，八边形，十字型，缺角矩形，立方体，棱台，太阳形，新月形，双括号，双大括号-->
         <xsl:when test="(.//a:prstGeom/@prst='triangle' or .//a:prstGeom/@prst='octagon' or .//a:prstGeom/@prst='plus' or .//a:prstGeom/@prst='plaque' or .//a:prstGeom/@prst='cube' or .//a:prstGeom/@prst='bevel' or .//a:prstGeom/@prst='sun' or .//a:prstGeom/@prst='moon' or .//a:prstGeom/@prst='bracketPair' or .//a:prstGeom/@prst='bracePair') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--平行四边形，梯形，六边形，同心圆，禁止符-->
         <xsl:when test="(.//a:prstGeom/@prst='parallelogram' or .//a:prstGeom/@prst='trapezoid' or .//a:prstGeom/@prst='hexagon' or .//a:prstGeom/@prst='donut' or .//a:prstGeom/@prst='noSmoking') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x * 21600 * $height div $width div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--圆柱形-->
         <xsl:when test="(.//a:prstGeom/@prst='can') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$y * 21600 div 200000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--折角形-->
         <xsl:when test="(.//a:prstGeom/@prst='foldedCorner') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="21600 - $x * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--笑脸-->
         <xsl:when test="(.//a:prstGeom/@prst='smileyFace') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="16200 + $y * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左中括号，右中括号-->
         <xsl:when test="(.//a:prstGeom/@prst='leftBracket' or .//a:prstGeom/@prst='rightBracket') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$y * 21600 * $width div $height div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--左大括号，右大括号-->
         <xsl:when test="(.//a:prstGeom/@prst='leftBrace' or .//a:prstGeom/@prst='rightBrace') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="y">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="width">
             <xsl:value-of select=".//a:xfrm/a:ext/@cx"/>
           </xsl:variable>
           <xsl:variable name="height">
             <xsl:value-of select=".//a:xfrm/a:ext/@cy"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x * 21600 * $width div $height div 100000"/>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="$y * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--肘形连接符，肘形箭头连接符，肘形双箭头连接符，曲线连接符，曲线箭头连接符，曲线双箭头连接符-->
         <xsl:when test="(.//a:prstGeom/@prst='bentConnector3' or .//a:prstGeom/@prst='curvedConnector3' or .//a:prstGeom/@prst='bentConnector4' or .//a:prstGeom/@prst='bentConnector2') and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x">
             <xsl:value-of select="number($tmp)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="$x * 21600 div 100000"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--2014-12-04 liuyangyang 添加新图形控制点转换-->
         <xsl:when test=".//a:prstGeom/@prst='circularArrow' and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x3">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x4">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:choose>
                 <xsl:when test="$x4 &lt; 10800000">
                   <xsl:value-of select="$x4 div 10800000 * 11796480"/>
                 </xsl:when>
                 <xsl:otherwise>
                   <xsl:value-of select="($x4 - 21600000) div 10800000 * 11796480"/>
                 </xsl:otherwise>
               </xsl:choose>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:choose>
                 <xsl:when test="$x3 &lt; 10800000">
                   <xsl:value-of select="$x3 div 10800000 * 11796480"/>
                 </xsl:when>
                 <xsl:otherwise>
                   <xsl:value-of select="($x3 - 21600000) div 10800000 * 11796480"/>
                 </xsl:otherwise>
               </xsl:choose>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="(1 - $x1 div 50000) * 10800"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <xsl:when test=".//a:prstGeom/@prst='leftCircularArrow' and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x3">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="adj4">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp4">
             <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x4">
             <xsl:value-of select="number($tmp4)"/>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:choose>
                 <xsl:when test="$x4 &lt; 10800000">
                   <xsl:value-of select="(10800000 - $x4) div 10800000 * 11796480"/>
                 </xsl:when>
                 <xsl:otherwise>
                   <xsl:value-of select="(10800000 - $x4) div 10800000 * 11796480"/>
                 </xsl:otherwise>
               </xsl:choose>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:choose>
                 <xsl:when test="$x3 &lt; 10800000">
                   <xsl:value-of select="(10800000 - $x3) div 10800000 * 11796480"/>
                 </xsl:when>
                 <xsl:otherwise>
                   <xsl:value-of select="(10800000 - $x3) div 10800000 * 11796480"/>
                 </xsl:otherwise>
               </xsl:choose>
             </xsl:attribute>
           </图:控制点_8039>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:value-of select="(1 - $x1 div 50000) * 10800"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <xsl:when test=".//a:prstGeom/@prst='blockArc' and ./*/a:prstGeom/a:avLst/a:gd">
           <xsl:variable name="adj1">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp1">
             <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x1">
             <xsl:value-of select="number($tmp1)"/>
           </xsl:variable>
           <xsl:variable name="adj3">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp3">
             <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x3">
             <xsl:value-of select="number($tmp3)"/>
           </xsl:variable>
           <xsl:variable name="adj2">
             <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
           </xsl:variable>
           <xsl:variable name="tmp2">
             <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
           </xsl:variable>
           <xsl:variable name="x2">
             <xsl:value-of select="number($tmp2)"/>
           </xsl:variable>
           <xsl:variable name="angle">
             <xsl:choose>
               <xsl:when test="$x1 &lt; $x2">
                 <xsl:value-of select="($x1 - $x2) div 21600000 + 1"/>
               </xsl:when>
               <xsl:otherwise>
                 <xsl:value-of select="($x1 - $x2) div 21600000"/>
               </xsl:otherwise>
             </xsl:choose>
           </xsl:variable>
           <图:控制点_8039>
             <xsl:attribute name="x_C606">
               <xsl:choose>
                 <xsl:when test="$angle &lt; 0.5">
                   <xsl:value-of select="(0.5 + $angle) * 11796480"/>
                 </xsl:when>
                 <xsl:otherwise>
                   <xsl:value-of select="( $angle - 1.5) * 11796480"/>
                 </xsl:otherwise>
               </xsl:choose>
             </xsl:attribute>
             <xsl:attribute name="y_C607">
               <xsl:value-of select="(1 - $x3 div 50000) * 10800"/>
             </xsl:attribute>
           </图:控制点_8039>
         </xsl:when>
         <!--end-->
       </xsl:choose>
       <!--end-->
			  <!--2011-5-30 罗文甜： 增加控制点,按照标准转换>
      <xsl:if test="/a:prstGeom/a:avLst">        
        <xsl:variable name="x" select="p:spPr/a:xfrm/a:off/@x div 12700"/>
        <xsl:variable name="y" select="p:spPr/a:xfrm/a:off/@y div 12700"/>
        <xsl:variable name="conx" select="p:spPr/a:xfrm/a:ext/@cx div 12700"/>
        <xsl:variable name="cony" select="p:spPr/a:xfrm/a:ext/@cy div 12700"/>
        <图:控制点 uof:locID="g0003" uof:attrList="x坐标 y坐标">
          <xsl:attribute name="图:x坐标">
            <xsl:value-of select="$x div ( $x + $conx )"/>
          </xsl:attribute>
          <xsl:attribute name="图:y坐标">
            <xsl:value-of select="$y div ( $y + $cony )"/>
          </xsl:attribute>
        </图:控制点>
      </xsl:if-->
		  </图:图形_8062>
  </xsl:template>
  <!--end -->
  
  <xsl:template match="p:pic">
    <xsl:param name="type"/>
	  <xsl:param name="picFrom" />
	  <图:图形_8062>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="translate(@id,':','')"/>
      </xsl:attribute>
		  <!-- 李娟 11.12.09 这部分有待研究····@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
       
		  <!--<xsl:if  test="p:nvPicPr">-->
			
		  <!--</xsl:if>-->
      <xsl:apply-templates select="p:spPr"/>
	  <xsl:if test="./p:blipFill">
		  <xsl:element name ="图:图片数据引用_8037">
			  <xsl:call-template name="findPicName"/>
			  <!--<xsl:variable name="rId">
				  <xsl:if test="./p:blipFill/a:blip">
					  <xsl:value-of select="./p:blipFill/a:blip/@r:embed|../../p:bg/p:bgPr/a:blipFill/a:blip/@r:embed"/>
				  </xsl:if>
			  </xsl:variable>-->
			  <!--<xsl:variable name="slideName">
				  <xsl:value-of select="concat(ancestor::p:sld/@id,'.rels')"/>
			  </xsl:variable>-->
			  <!--<xsl:for-each select="ancestor::p:presentation">
				  <xsl:if test=".//rel:Relationships[@id=$slideName]">
					 <xsl:variable name="target">
							  <xsl:value-of select=".//rel:Relationships[@id=$slideName]/rel:Relationship[@Id=$rId]/@Target"/>
						  </xsl:variable>
						  <xsl:variable name="imageName">
							  <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
						  </xsl:variable>
						  <xsl:value-of select="$imageName"/>
					</xsl:if>
			  </xsl:for-each>-->
			 </xsl:element>
	  </xsl:if>

      <!--2014-06-04, tangjiang, 添加图片翻转转换 start -->
      <xsl:choose>
        <xsl:when test="p:spPr/a:xfrm/@flipV and not(p:spPr/a:xfrm/@flipH)">
          <图:翻转_803A>y</图:翻转_803A>
        </xsl:when>
        <xsl:when test="p:spPr/a:xfrm/@flipH and not(p:spPr/a:xfrm/@flipV)">
          <图:翻转_803A>x</图:翻转_803A>
        </xsl:when>
        <xsl:when test="p:spPr/a:xfrm/@flipH and p:spPr/a:xfrm/@flipV">
          <图:翻转_803A>xy</图:翻转_803A>
        </xsl:when>
        <xsl:otherwise>
        </xsl:otherwise>
      </xsl:choose>
      <!--2014-06-04, tangjiang, 添加图片翻转转换 end -->

      <!--增加判断，插入图片可以互操作 liqiuling  2013-03-21 start-->
      <xsl:if test=".//p14:media/@r:embed">
        <图:其他对象引用_8038>
          <xsl:variable name="findPictureName">
            <xsl:value-of select=".//p14:media/@r:embed"/>
          </xsl:variable>

          <xsl:call-template name="findRelatedRelationships">
            <xsl:with-param name="id">
              <xsl:value-of select="$findPictureName"/>
            </xsl:with-param>
          </xsl:call-template>
        </图:其他对象引用_8038>
      </xsl:if>
      <!--增加判断，插入图片可以互操作 liqiuling  2013-03-21 end-->
      <xsl:if test="parent::p:grpSp">
        <xsl:call-template name="groupLocation">
        </xsl:call-template>
      </xsl:if>
    </图:图形_8062>
  </xsl:template>
  
  
  <xsl:template match="p:grpSp">
    <xsl:param name="type"/>
    <!--
    
    <xsl:for-each select="p:cxnSp|p:graphicFrame|p:grpSp|p:pic|p:sp">
    
    -->
    <xsl:for-each select="p:cxnSp|p:grpSp|p:pic|p:sp|p:graphicFrame">
      <xsl:apply-templates select=".">
        <xsl:with-param name="type" select="$type"/>
      </xsl:apply-templates>
    </xsl:for-each>
	  <图:图形_8062>
	  <!--<图:图形 uof:locID="g0000" uof:attrList="层次 标识符 组合列表 其他对象 OLE数据">-->
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="translate(@id,':','')"/>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <!--
        <xsl:for-each select="p:cxnSp|p:graphicFrame|p:grpSp|p:pic|p:sp">
        -->
        <xsl:for-each select="p:cxnSp|p:grpSp|p:pic|p:sp|p:graphicFrame">
          <xsl:value-of select="concat(' ',translate(@id,':',''))"/>
        </xsl:for-each>
      </xsl:attribute>
      <!--<xsl:apply-templates select="p:txBody"/>-->
      <xsl:apply-templates select="p:grpSpPr"/>
      <!--2011-5-26罗文甜，修改翻转bug-->
      <xsl:choose>
        <xsl:when test="p:grpSpPr/a:xfrm/@flipV and not(p:grpSpPr/a:xfrm/@flipH)">
			<图:翻转_803A>y</图:翻转_803A>
          <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="y"/>-->
        </xsl:when>
        <xsl:when test="p:grpSpPr/a:xfrm/@flipH and not(p:grpSpPr/a:xfrm/@flipV)">
			<图:翻转_803A>x</图:翻转_803A>
          <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="x"/>-->
        </xsl:when>
        <xsl:when test="p:grpSpPr/a:xfrm/@flipH and p:grpSpPr/a:xfrm/@flipV">
			<图:翻转_803A>xy</图:翻转_803A>
          <!--<图:翻转 uof:locID="g0040" uof:attrList="方向" 图:方向="xy"/>-->
        </xsl:when>
        <xsl:otherwise>
        </xsl:otherwise>
      </xsl:choose>
      <!--
       2010.2.5 黎美秀修改 增加组合位置
  
       -->
      <xsl:if test="parent::p:grpSp">
        <xsl:call-template name="groupLocation">
        </xsl:call-template>
      </xsl:if>
	  <!-- 添加图片数据引用 李娟11.12.09 ··@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
		  

	  </图:图形_8062>
  </xsl:template>


  <!--2012-12-20, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
  <xsl:template match="p:spPr|dsp:spPr">
    <xsl:call-template name="spPr"/>
  </xsl:template>
  <xsl:template match="p:grpSpPr|dsp:grpSpPr">
    <xsl:call-template name="spPr"/>
  </xsl:template>
  <!--end-->

  <!--4.15黎美秀修改-->
  <xsl:template name="spPr">
    <!--xsl:if test="a:prstGeom"-->
	  <图:预定义图形_8018>
		  <图:类别_8019>
        <xsl:choose>
          <xsl:when test="a:prstGeom/@prst='rect'">11</xsl:when>
          <xsl:when test="a:prstGeom/@prst='parallelogram'">12</xsl:when>
          <xsl:when test="a:prstGeom/@prst='trapezoid'">13</xsl:when>
          <xsl:when test="a:prstGeom/@prst='diamond'">14</xsl:when>
          <xsl:when test="a:prstGeom/@prst='roundRect'">15</xsl:when>
          <xsl:when test="a:prstGeom/@prst='octagon'">16</xsl:when>
          <xsl:when test="a:prstGeom/@prst='triangle'">17</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rtTriangle'">18</xsl:when>
          <!--罗文甜20110107斜纹，半闭框，L形转换为三角形-->
          <xsl:when test="a:prstGeom/@prst='diagStripe'">18</xsl:when>
          <xsl:when test="a:prstGeom/@prst='corner'">18</xsl:when>
          <xsl:when test="a:prstGeom/@prst='halfFrame'">18</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipse'">19</xsl:when>
          <!--2011-1-17罗文甜，7边形转换为8边形-->
          <xsl:when test="a:prstGeom/@prst='heptagon'">16</xsl:when>
          <!--2011-1-7罗文甜：十边形，十二边形，云，转换为圆-->
          <xsl:when test="a:prstGeom/@prst='decagon'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='dodecagon'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='pie'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='chord'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='teardrop'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cloud'">19</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightArrow'">21</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftArrow'">22</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upArrow'">23</xsl:when>
          <xsl:when test="a:prstGeom/@prst='downArrow'">24</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightArrow'">25</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upDownArrow'">26</xsl:when>
          <xsl:when test="a:prstGeom/@prst='quadArrow'">27</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightUpArrow'">28</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bentArrow'">29</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartProcess'">31</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartAlternateProcess'">32</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDecision'">33</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartInputOutput'">34</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPredefinedProcess'">35</xsl:when>
          <!--<xsl:when test="a:prstGeom/@prst='flowChartProcess'">35</xsl:when>过程：这块有错误-->
          <xsl:when test="a:prstGeom/@prst='flowChartInternalStorage'">36</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDocument'">37</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMultidocument'">38</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartTerminator'">39</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star4'">43</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star5'">44</xsl:when>
          <!--2011-1-7罗文甜：6、7角形转换为8边形-->
          <xsl:when test="a:prstGeom/@prst='star6'">45</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star7'">45</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star8'">45</xsl:when>
          <!--2011-1-7罗文甜：10、12角形转换为16角形-->
          <xsl:when test="a:prstGeom/@prst='star10'">46</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star12'">46</xsl:when>       
          <!--2010.10.12 罗文甜：修改bug，增加16星与旗帜对应-->
          <xsl:when test="a:prstGeom/@prst='star16'">46</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star24'">47</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star32'">48</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ribbon'">410</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cloudCallout'">54</xsl:when>
          <xsl:when test="a:prstGeom/@prst='line'">61</xsl:when>
          <xsl:when test="a:prstGeom/@prst='straightConnector1'">71</xsl:when>
          <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">72</xsl:when-->
          <xsl:when test="a:prstGeom/@prst='hexagon'">110</xsl:when>
          <xsl:when test="a:prstGeom/@prst='pentagon'">112</xsl:when>
          <xsl:when test="a:prstGeom/@prst='can'">113</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cube'">114</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bevel'">115</xsl:when>
          <xsl:when test="a:prstGeom/@prst='foldedCorner'">116</xsl:when>
          <xsl:when test="a:prstGeom/@prst='smileyFace'">117</xsl:when>
          <xsl:when test="a:prstGeom/@prst='donut'">118</xsl:when>
          <xsl:when test="a:prstGeom/@prst='blockArc'">120</xsl:when>
          <xsl:when test="a:prstGeom/@prst='heart'">121</xsl:when>
          <xsl:when test="a:prstGeom/@prst='lightningBolt'">122</xsl:when>
          <xsl:when test="a:prstGeom/@prst='sun'">123</xsl:when>
          <xsl:when test="a:prstGeom/@prst='moon'">124</xsl:when>
          <xsl:when test="a:prstGeom/@prst='arc'">125</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bracketPair'">126</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bracePair'">127</xsl:when>
          <xsl:when test="a:prstGeom/@prst='plaque'">128</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftBracket'">129</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightBracket'">130</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftBrace'">131</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightBrace'">132</xsl:when>
          <xsl:when test="a:prstGeom/@prst='uturnArrow'">210</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftUpArrow'">211</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bentUpArrow'">212</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedRightArrow'">213</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedLeftArrow'">214</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedUpArrow'">215</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedDownArrow'">216</xsl:when>
          <xsl:when test="a:prstGeom/@prst='stripedRightArrow'">217</xsl:when>
          <xsl:when test="a:prstGeom/@prst='notchedRightArrow'">218</xsl:when>
          <xsl:when test="a:prstGeom/@prst='chevron'">220</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightArrowCallout'">221</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftArrowCallout'">222</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upArrowCallout'">223</xsl:when>
          <xsl:when test="a:prstGeom/@prst='downArrowCallout'">224</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightArrowCallout'">225</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upDownArrowCallout'">226</xsl:when>
          <!--liuyangyang 2014-12-09 添加左圆箭头转圆箭头-->
          <xsl:when test="a:prstGeom/@prst='circularArrow' or a:prstGeom/@prst='leftCircularArrow'">228</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPreparation'">310</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartManualInput'">311</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartManualOperation'">312</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartConnector'">313</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOffpageConnector'">314</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPunchedCard'">315</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPunchedTape'">316</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartSummingJunction'">317</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOr'">318</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartCollate'">319</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartSort'">320</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartExtract'">321</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMerge'">322</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDelay'">324</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticDisk'">326</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDisplay'">328</xsl:when>
          <!--<xsl:when test="a:prstGeom/@prst='ribbon2'">410</xsl:when>-->
          <xsl:when test="a:prstGeom/@prst='ribbon2'">49</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipseRibbon'">412</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipseRibbon2'">411</xsl:when>
          <xsl:when test="a:prstGeom/@prst='verticalScroll'">413</xsl:when>
          <xsl:when test="a:prstGeom/@prst='horizontalScroll'">414</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wave'">415</xsl:when>
          <xsl:when test="a:prstGeom/@prst='doubleWave'">416</xsl:when>
          <!--4月10日蒋俊彦改-->
          <xsl:when test="a:prstGeom/@prst='noSmoking'">119</xsl:when>
          <xsl:when test="a:prstGeom/@prst='plus'">111</xsl:when>
          <!--2011-1-7罗文甜：加号、乘、除转换为十字-->
          <xsl:when test="a:prstGeom/@prst='mathPlus'">111</xsl:when>
          <xsl:when test="a:prstGeom/@prst='mathMultiply'">111</xsl:when>
          <xsl:when test="a:prstGeom/@prst='mathDivide'">111</xsl:when>
          <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">71</xsl:when-->
          <xsl:when test="a:prstGeom/@prst='irregularSeal1'">41</xsl:when>
          <xsl:when test="a:prstGeom/@prst='irregularSeal2'">42</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeRectCallout'">51</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeRoundRectCallout'">52</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeEllipseCallout'">53</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout1'">56</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout2'">57</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout3'">58</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout1'">510</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout2'">511</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout3'">512</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout1'">514</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout2'">515</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout3'">516</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout1'">518</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout2'">519</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout3'">520</xsl:when>
          <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">62</xsl:when-->
          <!--xsl:when test="a:prstGeom/@prst='curvedConnector3'">64</xsl:when-->
          <xsl:when test="a:prstGeom/@prst='bentConnector3' or a:prstGeom/@prst='bentConnector4' or a:prstGeom/@prst='bentConnector2'">74</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedConnector3'">77</xsl:when>
          <!--2011-12-08 罗文甜，增加类别类型-->
          <xsl:when test="a:prstGeom/@prst='curvedConnector4'">77</xsl:when>
          <xsl:when test="a:prstGeom/@prst='homePlate'">219</xsl:when>
          <xsl:when test="a:prstGeom/@prst='quadArrowCallout'">227</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOnlineStorage'">323</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticTape'">325</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticDrum'">327</xsl:when>
          <!--liuyangyang 添加对gear6 gear9转换为star5 star8-->
          <xsl:when test="a:prstGeom/@prst='gear6'">44</xsl:when>
          <xsl:when test="a:prstGeom/@prst='gear9'">45</xsl:when>
          <!--end -->
          <xsl:when test="a:custGeom">64</xsl:when>
          <xsl:otherwise>11</xsl:otherwise>
        </xsl:choose>
      </图:类别_8019>
		  <图:名称_801A>
        <xsl:choose>
          <!--4月10日蒋俊彦改-->
          <xsl:when test="a:prstGeom/@prst='noSmoking'">No Symbol</xsl:when>
          <xsl:when test="a:prstGeom/@prst='plus'">Cross</xsl:when>
          <!--2011-1-7罗文甜：加号、减号、除号转换为十字-->
          <xsl:when test="a:prstGeom/@prst='mathPlus'">Cross</xsl:when>
          <xsl:when test="a:prstGeom/@prst='mathMultiply'">Cross</xsl:when>
          <xsl:when test="a:prstGeom/@prst='mathDivide'">Cross</xsl:when>
          <xsl:when test="a:prstGeom/@prst='irregularSeal1'">Explosion 1</xsl:when>
          <xsl:when test="a:prstGeom/@prst='irregularSeal2'">Explosion 2</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeRectCallout'">Rectangular Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeRoundRectCallout'">Rounded Rectangular Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wedgeEllipseCallout'">Oval Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout1'">Line Callout2</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout2'">Line Callout3</xsl:when>
          <xsl:when test="a:prstGeom/@prst='borderCallout3'">Line Callout4</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout1'">Line Callout2(Accent Bar)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout2'">Line Callout3(Accent Bar)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentCallout3'">Line Callout4(Accent Bar)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout1'">Line Callout2(No Border)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout2'">Line Callout3(No Border)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='callout3'">Line Callout4(No Border)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout1'">Line Callout2(Border and Accent Bar)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout2'">Line Callout3(Border and Accent Bar)</xsl:when>
          <xsl:when test="a:prstGeom/@prst='accentBorderCallout3'">Line Callout4(Border and Accent Bar)</xsl:when>
          <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">Arrow</xsl:when-->
          <!--xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curve</xsl:when-->
          <xsl:when test="a:prstGeom/@prst='bentConnector3' or a:prstGeom/@prst='bentConnector4' or a:prstGeom/@prst='bentConnector2'">Elbow Connector</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curved Connector</xsl:when>
          <!--2010-12-08 罗文甜,增加名称-->
          <xsl:when test="a:prstGeom/@prst='curvedConnector4'">Curved Connector</xsl:when>
          <xsl:when test="a:prstGeom/@prst='homePlate'">Pentagon Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='quadArrowCallout'">Quad Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOnlineStorage'">Stored Data</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticTape'">Sequential Access Storage</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticDrum'">Direct Access Storage</xsl:when>
          <!---->
          <xsl:when test="a:prstGeom/@prst='rect' or not(a:prstGeom/@prst)">Rectangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='parallelogram'">Parallelogram</xsl:when>
          <xsl:when test="a:prstGeom/@prst='trapezoid'">Trapezoid</xsl:when>
          <xsl:when test="a:prstGeom/@prst='diamond'">Diamond</xsl:when>
          <xsl:when test="a:prstGeom/@prst='roundRect'">Rounded Rectangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='octagon'">Octagon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='triangle'">Isosceles Triangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rtTriangle'">Right Triangle</xsl:when>
          <!--罗文甜20110107斜纹，半闭框，L形转换为三角形-->
          <xsl:when test="a:prstGeom/@prst='diagStripe'">Right Triangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='corner'">Right Triangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='halfFrame'">Right Triangle</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipse'">Oval</xsl:when>
          <!--2011-1-17罗文甜，7边形转换为8边形-->
          <xsl:when test="a:prstGeom/@prst='heptagon'">Octagon</xsl:when>
          <!--2011-1-7罗文甜：十边形，十二边形，饼，泪滴，弦月，云，转换为圆-->
          <xsl:when test="a:prstGeom/@prst='decagon'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='dodecagon'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='pie'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='chord'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='teardrop'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cloud'">Oval</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightArrow'">Right Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftArrow'">Left Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upArrow'">Up Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='downArrow'">Down Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightArrow'">Left-Right Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upDownArrow'">Up-Down Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='quadArrow'">Quad Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightUpArrow'">Left-Right-Up Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bentArrow'">Bent Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartAlternateProcess'">Alternate Process</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDecision'">Decision</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPredefinedProcess'">Predefined Process</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartInputOutput'">Data</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartInternalStorage'">Internal Storage</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDocument'">Document</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMultidocument'">Multidocument</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartTerminator'">Terminator</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star4'">4-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star5'">5-Point Star</xsl:when>
          <!--2011-1-7罗文甜：6、7角形转换为8角形-->
          <xsl:when test="a:prstGeom/@prst='star6'">8-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star7'">8-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star8'">8-Point Star</xsl:when>
          <!--2011-1-7罗文甜：10、12角形转换为16角形-->
          <xsl:when test="a:prstGeom/@prst='star10'">16-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star12'">16-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star16'">16-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star24'">24-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='star32'">32-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ribbon'">Down Ribbon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cloudCallout'">Parallelogram</xsl:when>
          <xsl:when test="a:prstGeom/@prst='line'">Line</xsl:when>
          <xsl:when test="a:prstGeom/@prst='straightConnector1'">Straight Connector</xsl:when>
          <xsl:when test="a:prstGeom/@prst='hexagon'">Hexagon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='pentagon'">Regual Pentagon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='can'">Can</xsl:when>
          <xsl:when test="a:prstGeom/@prst='cube'">Cube</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bevel'">Bevel</xsl:when>
          <xsl:when test="a:prstGeom/@prst='foldedCorner'">Folded Corner</xsl:when>
          <xsl:when test="a:prstGeom/@prst='smileyFace'">Smiley Face</xsl:when>
          <xsl:when test="a:prstGeom/@prst='donut'">Donut</xsl:when>
          <xsl:when test="a:prstGeom/@prst='blockArc'">Block Arc</xsl:when>
          <xsl:when test="a:prstGeom/@prst='heart'">Heart</xsl:when>
          <xsl:when test="a:prstGeom/@prst='lightningBolt'">Lightning</xsl:when>
          <xsl:when test="a:prstGeom/@prst='sun'">Sun</xsl:when>
          <xsl:when test="a:prstGeom/@prst='moon'">Moon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='arc'">Arc</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bracketPair'">Double Bracket</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bracePair'">Double Brace</xsl:when>
          <xsl:when test="a:prstGeom/@prst='plaque'">Plaque</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftBracket'">Left Bracket</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightBracket'">Right Bracket</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftBrace'">Left Brace</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightBrace'">Right Brace</xsl:when>
          <xsl:when test="a:prstGeom/@prst='uturnArrow'">U-Turn Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftUpArrow'">Left-Up Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='bentUpArrow'">Bent-Up Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedRightArrow'">Curved Right Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedLeftArrow'">Curved Left Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedUpArrow'">Curved Up Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='curvedDownArrow'">Curved Down Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='stripedRightArrow'">Striped Right Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='notchedRightArrow'">Notched Right Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='chevron'">Chevron Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='rightArrowCallout'">Right Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftArrowCallout'">Left Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upArrowCallout'">Up Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='downArrowCallout'">Down Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='leftRightArrowCallout'">Left-Right Arrow Callout</xsl:when>
          <xsl:when test="a:prstGeom/@prst='upDownArrowCallout'">Up-Down Arrow Callout</xsl:when>
          <!--liuyangyang 2014-12-09 添加左圆箭头转圆箭头-->
          <xsl:when test="a:prstGeom/@prst='circularArrow' or a:prstGeom/@prst='leftCircularArrow'">Circular Arrow</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPreparation'">Preparation</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartManualInput'">Manual Input</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartManualOperation'">Manual Operation</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartConnector'">Connector</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOffpageConnector'">Off-page Connector</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPunchedCard'">Card</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartPunchedTape'">Punched Tape</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartSummingJunction'">Summing Junction</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartOr'">Or</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartCollate'">Collate</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartSort'">Sort</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartExtract'">Extract</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMerge'">Merge</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDelay'">Delay</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartMagneticDisk'">Magnetic Disk</xsl:when>
          <xsl:when test="a:prstGeom/@prst='flowChartDisplay'">Display</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ribbon2'">Up Ribbon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipseRibbon'">Curved Down Ribbon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='ellipseRibbon2'">Curved Up Ribbon</xsl:when>
          <xsl:when test="a:prstGeom/@prst='verticalScroll'">Vertical Scroll</xsl:when>
          <xsl:when test="a:prstGeom/@prst='horizontalScroll'">Horizontal Scroll</xsl:when>
          <xsl:when test="a:prstGeom/@prst='wave'">Wave</xsl:when>
          <xsl:when test="a:prstGeom/@prst='doubleWave'">Double Wave</xsl:when>
          <xsl:when test="a:prstGeom/@prst='gear6'">5-Point Star</xsl:when>
          <xsl:when test="a:prstGeom/@prst='gear9'">8-Point Star</xsl:when>
          <xsl:when test="a:custGeom/*">Curve</xsl:when>
          <xsl:otherwise>Rectangle</xsl:otherwise>
        </xsl:choose>
      </图:名称_801A>
		  <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
      <!--2010-12-03罗文甜：增加自由曲线的关键点坐标（思路：调用两个模板，如果是直线，传两个参数，调用模板1，如果是curve，传6个参数参，调用模板2）：-->
      <xsl:if test="a:custGeom/a:pathLst/a:path">
        <!--2014-05-05, tangjiang, 修复多路径转换 start -->
        <xsl:for-each select="a:custGeom/a:pathLst/a:path">
          <图:路径_801C>
            <xsl:call-template name="keyCorCommen"/>
          </图:路径_801C>
        </xsl:for-each>
        <!--end 2014-05-05, tangjiang, 修复多路径转换 -->
      </xsl:if>
	 <图:属性_801D>
        <!--填充-->
        <!--2011-6-24 罗文甜 组合图形填充-添加无填充情况-->
        <xsl:choose>
          <xsl:when test="a:grpFill">
            <xsl:choose>
              <xsl:when test="ancestor::p:grpSp/p:grpSpPr[a:noFill]/a:noFill">
              </xsl:when>
              <xsl:otherwise>
				  <图:填充_804C>
                  <xsl:apply-templates select="parent::p:grpSp/p:grpSpPr/a:gradFill
                                 |parent::p:grpSp/p:grpSpPr/a:pattFill
                                 |parent::p:grpSp/p:grpSpPr/a:solidFill
                                 |parent::p:grpSp/p:grpSpPr/a:blipFill
                                 |ancestor::p:grpSp/p:grpSpPr[a:gradFill]/a:gradFill
                                 |ancestor::p:grpSp/p:grpSpPr[a:pattFill]/a:pattFill
                                 |ancestor::p:grpSp/p:grpSpPr[a:solidFill]/a:solidFill
                                 |ancestor::p:grpSp/p:grpSpPr[a:blipFill]/a:blipFill"/>
                </图:填充_804C>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="a:gradFill">
				  <图:填充_804C>
                  <xsl:apply-templates select="a:gradFill"/>
                </图:填充_804C>
              </xsl:when>
              <xsl:when test="a:pattFill">
				  <图:填充_804C>
                  <xsl:apply-templates select="a:pattFill"/>
                </图:填充_804C>
              </xsl:when>
              <xsl:when test="a:solidFill">
				  <图:填充_804C>
					  <!--<图:颜色_8004>#FF0000</图:颜色_8004>-->
                  <xsl:apply-templates select="a:solidFill"/>
                </图:填充_804C>
              </xsl:when>
              <xsl:when test="a:blipFill">
				  <图:填充_804C>
                  <xsl:apply-templates select="a:blipFill"/>
                </图:填充_804C>
              </xsl:when>
              <xsl:when test="a:noFill"/>
              <!--有style的情况-->
              <!--10.02.04 马有旭 BugID:2944417 <xsl:when test="following-sibling::p:style/a:fillRef">-->
              <xsl:when test="following-sibling::p:style/a:fillRef[@idx!='0' and @idx!='1000']">
                <xsl:apply-templates select="following-sibling::p:style/a:fillRef"/>
              </xsl:when>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>


     <!--2014-04-27, tangjiang, 修复图形多出边框线 start -->
        <!--2011-01-07罗文甜：线-->
        <xsl:choose>
          <!--优先考虑组合图形-->
          <xsl:when test="../../p:grpSp/p:spPr/a:ln">
            <xsl:if test="not(../../p:grpSp/p:spPr/a:ln/a:noFill)">
              <图:线_8057>
                <xsl:apply-templates select="../../p:grpSp/p:spPr/a:ln"/>
              </图:线_8057>
            </xsl:if>
          </xsl:when>             
          <!--一般情况-->      
          <xsl:when test="a:ln/* or a:ln/@*">
            <xsl:if test="not(a:ln/a:noFill)">
              <图:线_8057>
                <xsl:apply-templates select="a:ln"/>
              </图:线_8057>
            </xsl:if>
          </xsl:when>
          <!--有style的情况-->
          <xsl:when test="following-sibling::p:style/a:lnRef">
            <xsl:apply-templates select="following-sibling::p:style/a:lnRef"/>
          </xsl:when>
		 </xsl:choose>
     <!--2014-04-27, tangjiang, 修复图形多出边框线 -->

     <xsl:if test="./a:ln/a:headEnd|./a:ln/a:tailEnd">
				  <图:箭头_805D>
					  <xsl:if test="./a:ln/a:headEnd/@type">
						  <!--修改箭头 李娟 11.11.30-->

						  <图:前端箭头_805E>
							  <图:式样_8000>
								  <xsl:choose>
									  <xsl:when test="./a:ln/a:headEnd/@type='triangle'">normal</xsl:when>
									  <xsl:when test="./a:ln/a:headEnd/@type='diamond'">diamond</xsl:when>
									  <xsl:when test="./a:ln/a:headEnd/@type='arrow'">open</xsl:when>
									  <xsl:when test="./a:ln/a:headEnd/@type='oval'">oval</xsl:when>
									  <xsl:when test="./a:ln/a:headEnd/@type='stealth'">stealth</xsl:when>
								  </xsl:choose>
							  </图:式样_8000>
							  <!--2011-3-20罗文甜，箭头大小-->
                <!--2013-1-1, liqiuling, 修复bug箭头大小  start -->
							  <xsl:choose>
								  <xsl:when test="./a:ln/a:headEnd/@w='sm' and ./a:ln/a:headEnd/@len='sm'">
									  <图:大小_8001>1</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='sm' and ./a:ln/a:headEnd/@len='med'">
									  <图:大小_8001>2</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='sm' and ./a:ln/a:headEnd/@len='lg'">
									  <图:大小_8001>3</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='med' and ./a:ln/a:headEnd/@len='sm'">
									  <图:大小_8001>4</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='med' and ./a:ln/a:headEnd/@len='lg'">
									  <图:大小_8001>6</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='lg' and ./a:ln/a:headEnd/@len='sm'">
									  <图:大小_8001>7</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='lg' and ./a:ln/a:headEnd/@len='med'">
									  <图:大小_8001>8</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:headEnd/@w='lg' and ./a:ln/a:headEnd/@len='lg'">
									  <图:大小_8001>9</图:大小_8001>
								  </xsl:when>
								  <xsl:otherwise>
								 <图:大小_8001>5</图:大小_8001>
                  
								  </xsl:otherwise>
							  </xsl:choose>
						  </图:前端箭头_805E>

					  </xsl:if>
					  <!--09.12.02 马有旭<xsl:if test="a:tailEnd">-->
					  <xsl:if test="./a:ln/a:tailEnd/@type">

						  <图:后端箭头_805F>
							  <图:式样_8000>
								  <xsl:choose>
									  <xsl:when test="./a:ln/a:tailEnd/@type='triangle'">normal</xsl:when>
									  <xsl:when test="./a:ln/a:tailEnd/@type='diamond'">diamond</xsl:when>
									  <xsl:when test="./a:ln/a:tailEnd/@type='arrow'">open</xsl:when>
									  <xsl:when test="./a:ln/a:tailEnd/@type='oval'">oval</xsl:when>
									  <xsl:when test="./a:ln/a:tailEnd/@type='stealth'">stealth</xsl:when>
								  </xsl:choose>
							  </图:式样_8000>
							  <!--2011-3-20罗文甜，箭头大小-->
							  <xsl:choose>
								  <xsl:when test="./a:ln/a:tailEnd/@w='sm' and ./a:ln/a:tailEnd/@len='sm'">
									  <图:大小_8001>1</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='sm' and ./a:ln/a:tailEnd/@len='med'">
									  <图:大小_8001>2</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='sm' and ./a:ln/a:tailEnd/@len='lg'">
									  <图:大小_8001>3</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='med' and ./a:ln/a:tailEnd/@len='sm'">
									  <图:大小_8001>4</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='med' and ./a:ln/a:tailEnd/@len='lg'">
									  <图:大小_8001>6</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='lg' and ./a:ln/a:tailEnd/@len='sm'">
									  <图:大小_8001>7</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='lg' and ./a:ln/a:tailEnd/@len='med'">
									  <图:大小_8001>8</图:大小_8001>
								  </xsl:when>
								  <xsl:when test="./a:ln/a:tailEnd/@w='lg' and ./a:ln/a:tailEnd/@len='lg'">
									  <图:大小_8001>9</图:大小_8001>
								  </xsl:when>
								  <xsl:otherwise>
									 <图:大小_8001>5</图:大小_8001>
								  </xsl:otherwise>
							  </xsl:choose>
						  </图:后端箭头_805F>

					  </xsl:if>
				  </图:箭头_805D>
			  </xsl:if>
     <!--2013-1-1, liqiuling, 箭头大小  end -->
       <!--阴影,透明度,艺术字,三维效果等没有添加 李娟 11.11.11-->
 
        <xsl:if test="a:xfrm/a:ext">
			<图:大小_8060>
        <xsl:attribute name="长_C604">
          <xsl:value-of select="a:xfrm/a:ext/@cy div 12700"/>

        </xsl:attribute>
				<xsl:attribute name="宽_C605">
					<xsl:value-of select="a:xfrm/a:ext/@cx div 12700"/>
					
				</xsl:attribute>
			</图:大小_8060>
        </xsl:if>
			  <!--11.11.28 添加三维效果 李娟-->
     <!--修复三维效果转换不正确，2013-2-25  liqiuling start-->
		<xsl:if test="a:scene3d">
				  <图:三维效果_8061>
				  <xsl:if test="not(../a:sp3d/@extrusionH)">
					  <!--无深度-->
					  <xsl:for-each select="a:scene3d/a:camera">
						  <xsl:if test="./a:rot">
							  <uof:角度_C635>
								  <uof:x方向_C636>
									  <xsl:value-of select="./a:rot/@lon div 60000"/>
								  </uof:x方向_C636>
								  <uof:y方向_C637>
									  <xsl:value-of select="./a:rot/@lat div 60000"/>
								  </uof:y方向_C637>
							  </uof:角度_C635>
							  <uof:方向_C63C>
								  <uof:角度_C639>
									  <xsl:value-of select="'none'"/>
								  </uof:角度_C639>
								  <xsl:choose>
									  <xsl:when test="./@prst='orthographicFront'">
										  <!--无旋转，只有在设置了三维效果时会出现-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricLeftDown'">
										  <!--等轴左下-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricRightUp'">
										  <!--等轴右上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricTopUp'">
										  <!--等长顶部朝上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricBottomDown'">
										  <!--等长底部朝下-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis1Left'">
										  <!--离轴1左-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis1Right'">
										  <!--离轴1右-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis1Top'">
										  <!--离轴1上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis2Left'">
										  <!--离轴2左-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis2Right'">
										  <!--离轴2右-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='isometricOffAxis2Top'">
										  <!--离轴2上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveFront'">
										  <!--前透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveLeft'">
										  <!--左透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveRight'">
										  <!--右透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveBelow'">
										  <!--下透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveAbove'">
										  <!--上透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveRelaxedModerately'">
										  <!--适度宽松透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveRelaxed'">
										  <!--宽松透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveContrastingLeftFacing'">
										  <!--左向对比透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveContrastingRightFacing'">
										  <!--右向对比透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'">
										  <!--极左极大透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'">
										  <!--极右极大透视-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='obliqueTopLeft'">
										  <!--倾斜左上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='obliqueTopRight'">
										  <!--倾斜右上-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='obliqueBottomLeft'">
										  <!--倾斜左下-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:when test="./@prst='obliqueBottomRight'">
										  <!--倾斜右下-->
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </xsl:when>
									  <xsl:otherwise>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </xsl:otherwise>
								  </xsl:choose>
							  </uof:方向_C63C>
						  </xsl:if>
						  <xsl:if test="not(./a:rot)">
							  <!--如果没有这个标签，则根据预设属性将值写死-->
							  <xsl:choose>
								  <xsl:when test="./@prst='orthographicFront'">
									  <!--无旋转，只有在设置了三维效果时会出现-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricLeftDown'">
									  <!--等轴左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>45.0</uof:x方向_C636>
										  <uof:y方向_C637>35.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricRightUp'">
									  <!--等轴右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>315.0</uof:x方向_C636>
										  <uof:y方向_C637>35.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricTopUp'">
									  <!--等长顶部朝上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>315.0</uof:x方向_C636>
										  <uof:y方向_C637>325.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricBottomDown'">
									  <!--等长底部朝下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>315.0</uof:x方向_C636>
										  <uof:y方向_C637>35.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Left'">
									  <!--离轴1左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>64.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Right'">
									  <!--离轴1右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>334.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Top'">
									  <!--离轴1上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>307.0</uof:x方向_C636>
										  <uof:y方向_C637>301.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Left'">
									  <!--离轴2左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>26.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Right'">
									  <!--离轴2右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>296.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Top'">
									  <!--离轴2上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>54.0</uof:x方向_C636>
										  <uof:y方向_C637>301.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveFront'">
									  <!--前透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveLeft'">
									  <!--左透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>20.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRight'">
									  <!--右透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>340.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveBelow'">
									  <!--下透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>20.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveAbove'">
									  <!--上透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>340.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxedModerately'">
									  <!--适度宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>324.8</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxed'">
									  <!--宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>310.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingLeftFacing'">
									  <!--左向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>44.0</uof:x方向_C636>
										  <uof:y方向_C637>10.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingRightFacing'">
									  <!--右向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>316.0</uof:x方向_C636>
										  <uof:y方向_C637>10.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'">
									  <!--极左极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>35.0</uof:x方向_C636>
										  <uof:y方向_C637>8.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'">
									  <!--极右极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>326.0</uof:x方向_C636>
										  <uof:y方向_C637>8.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopLeft'">
									  <!--倾斜左上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopRight'">
									  <!--倾斜右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomLeft'">
									  <!--倾斜左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomRight'">
									  <!--倾斜右下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:otherwise>
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:if>
					  </xsl:for-each>
				  </xsl:if>
				  <xsl:if test="../a:sp3d/@extrusionH">
					  <!--深度-->
					  <xsl:for-each select="./a:camera">
						  <xsl:if test="./a:rot">
							  <uof:角度_C635>
								  <uof:x方向_C636>
									  <xsl:value-of select="./a:rot/@lon div 60000"/>
								  </uof:x方向_C636>
								  <uof:y方向_C637>
									  <xsl:value-of select="./a:rot/@lat div 60000"/>
								  </uof:y方向_C637>
							  </uof:角度_C635>
							  <xsl:choose>
								  <xsl:when test="./@prst='orthographicFront'">
									  <!--无旋转，只有在设置了三维效果时会出现-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'none'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricLeftDown'">
									  <!--等轴左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricRightUp'">
									  <!--等轴右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricTopUp'">
									  <!--等长顶部朝上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricBottomDown'">
									  <!--等长底部朝下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Left'">
									  <!--离轴1左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Right'">
									  <!--离轴1右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Top'">
									  <!--离轴1上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Left'">
									  <!--离轴2左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Right'">
									  <!--离轴2右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Top'">
									  <!--离轴2上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveFront'">
									  <!--前透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'none'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveLeft'">
									  <!--左透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRight'">
									  <!--右透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveBelow'">
									  <!--下透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveAbove'">
									  <!--上透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxedModerately'">
									  <!--适度宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxed'">
									  <!--宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingLeftFacing'">
									  <!--左向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingRightFacing'">
									  <!--右向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'">
									  <!--极左极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'">
									  <!--极右极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopLeft'">
									  <!--倾斜左上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopRight'">
									  <!--倾斜右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomLeft'">
									  <!--倾斜左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-left'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomRight'">
									  <!--倾斜右下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-bottom-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:when>
								  <xsl:otherwise>
									  <uof:方向_C63C>
										  <uof:角度_C639>
											  <xsl:value-of select="'to-top-right'"/>
										  </uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:if>
						  <xsl:if test="not(./a:rot)">
							  <!--如果没有这个标签，则根据预设属性将值写死-->
							  <xsl:choose>
								  <xsl:when test="./@prst='orthographicFront'">
									  <!--无旋转，只有在设置了三维效果时会出现-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricLeftDown'">
									  <!--等轴左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>45.0</uof:x方向_C636>
										  <uof:y方向_C637>35.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricRightUp'">
									  <!--等轴右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>315.0</uof:x方向_C636>
										  <uof:y方向_C637>35.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricTopUp'">
									  <!--等长顶部朝上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>314.7</uof:x方向_C636>
										  <uof:y方向_C637>324.6</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricBottomDown'">
									  <!--等长底部朝下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>314.7</uof:x方向_C636>
										  <uof:y方向_C637>35.4</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Left'">
									  <!--离轴1左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>64.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Right'">
									  <!--离轴1右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>334.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis1Top'">
									  <!--离轴1上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>306.5</uof:x方向_C636>
										  <uof:y方向_C637>301.3</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Left'">
									  <!--离轴2左-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>26.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Right'">
									  <!--离轴2右-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>296.0</uof:x方向_C636>
										  <uof:y方向_C637>18.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='isometricOffAxis2Top'">
									  <!--离轴2上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>53.5</uof:x方向_C636>
										  <uof:y方向_C637>301.3</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveFront'">
									  <!--前透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>none</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveLeft'">
									  <!--左透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>20.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRight'">
									  <!--右透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>340.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveBelow'">
									  <!--下透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>20.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveAbove'">
									  <!--上透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>340.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxedModerately'">
									  <!--适度宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>324.8</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveRelaxed'">
									  <!--宽松透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>309.6</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingLeftFacing'">
									  <!--左向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>43.9</uof:x方向_C636>
										  <uof:y方向_C637>10.4</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveContrastingRightFacing'">
									  <!--右向对比透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>316.1</uof:x方向_C636>
										  <uof:y方向_C637>10.4</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'">
									  <!--极左极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>34.5</uof:x方向_C636>
										  <uof:y方向_C637>8.1</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'">
									  <!--极右极大透视-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'perspective'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>325.5</uof:x方向_C636>
										  <uof:y方向_C637>8.1</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopLeft'">
									  <!--倾斜左上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueTopRight'">
									  <!--倾斜右上-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomLeft'">
									  <!--倾斜左下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-left</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:when test="./@prst='obliqueBottomRight'">
									  <!--倾斜右下-->
									  <uof:方向_C63C>
										  <uof:角度_C639>to-bottom-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:when>
								  <xsl:otherwise>
									  <uof:方向_C63C>
										  <uof:角度_C639>to-top-right</uof:角度_C639>
										  <uof:方式_C63D>
											  <xsl:value-of select="'parallel'"/>
										  </uof:方式_C63D>
									  </uof:方向_C63C>
									  <uof:角度_C635>
										  <uof:x方向_C636>0.0</uof:x方向_C636>
										  <uof:y方向_C637>0.0</uof:y方向_C637>
									  </uof:角度_C635>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:if>
					  </xsl:for-each>
				  </xsl:if>
            <!--修复三维效果转换不正确，2013-2-25 2013-5-22   liqiuling end-->
				  <xsl:if test="child::a:sp3d">
					  <xsl:if test="child::a:sp3d/a:extrusionClr/a:srgbClr">
						  <uof:颜色_C63F>
							  <xsl:value-of select="../a:sp3d/a:extrusionClr/a:srgbClr/@val"/>
						  </uof:颜色_C63F>
					  </xsl:if>
					  <xsl:if test="child::a:sp3d/@extrusionH">
						  <uof:深度_C63B>
							  <xsl:value-of select="child::a:sp3d/@extrusionH div 12700"/>
						  </uof:深度_C63B>
					  </xsl:if>
					  <xsl:for-each select="../a:sp3d">
						  <uof:表面效果_C63E>
							  <xsl:choose>
								  <xsl:when test="../a:sp3d/@prstMaterial='legacyWireframe'">wire-frame</xsl:when>
								  <!--线框-->
								  <!--以下效果均转化为UOF中亚光效果-->
								  <xsl:when test="../a:sp3d/@prstMaterial='matte'">matte</xsl:when>
								  <!--亚光效果-->
								  <xsl:when test="not(../a:sp3d/@prstMaterial)">matte</xsl:when>
								  <!--默认为暖色粗糙-->
								  <xsl:when test="../a:sp3d/@prstMaterial='dkEdge'">matte</xsl:when>
								  <!--硬边缘-->
								  <!--以下效果均转化为UOF中塑料效果-->
								  <xsl:when test="../a:sp3d/@prstMaterial='plastic'">plastic</xsl:when>
								  <!--塑料效果-->
								  <xsl:when test="../a:sp3d/@prstMaterial='flat'">plastic</xsl:when>
								  <!--平面-->
								  <xsl:when test="../a:sp3d/@prstMaterial='powder'">plastic</xsl:when>
								  <!--粉-->
								  <xsl:when test="../a:sp3d/@prstMaterial='translucentPowder'">plastic</xsl:when>
								  <!--半透明粉-->
								  <!--以下效果均转化为UOF中金属效果-->
								  <xsl:when test="../a:sp3d/@prstMaterial='metal'">metal</xsl:when>
								  <!--金属效果-->
								  <xsl:when test="../a:sp3d/@prstMaterial='softEdge'">metal</xsl:when>
								  <!--柔边缘-->
								  <xsl:when test="../a:sp3d/@prstMaterial='clear'">metal</xsl:when>
								  <!--最浅-->
								  <xsl:otherwise>matte</xsl:otherwise>
							  </xsl:choose>
						  </uof:表面效果_C63E>
					  </xsl:for-each>
				  </xsl:if>
				  <xsl:if test="not(child::a:sp3d)">
					  <uof:深度_C63B>0</uof:深度_C63B>
					  <uof:表面效果_C63E>matte</uof:表面效果_C63E>
				  </xsl:if>
				  <xsl:for-each select="a:lightRig">
					  <uof:照明_C638>
						  <uof:角度_C639>
							  <!--照明角度-->
							  <xsl:choose>
								  <!--<xsl:when test="@dir='t'">0</xsl:when>-->
								  <!--dir属性只发现了取t值的情况-->
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;= 0 and (./a:rot/@rev div 60000)&lt;=30 ">0</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=31 and (./a:rot/@rev div 60000)&lt;=70">45</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=71 and (./a:rot/@rev div 60000)&lt;=120">90</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=121 and (./a:rot/@rev div 60000)&lt;=160">135</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=161 and (./a:rot/@rev div 60000)&lt;=200">180</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=201 and (./a:rot/@rev div 60000)&lt;=240">225</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=241 and (./a:rot/@rev div 60000)&lt;=300">270</xsl:when>
								  <xsl:when test="(./a:rot/@rev div 60000)&gt;=301 and (./a:rot/@rev div 60000)&lt;=360">315</xsl:when>
								  <xsl:otherwise>0</xsl:otherwise>
							  </xsl:choose>
						  </uof:角度_C639>
						  <uof:强度_C63A>
							  <!--照明效果-->
							  <xsl:choose>
								  <!--以下效果转化为UOF中的bright效果-->
								  <xsl:when test="@rig='balanced'">bright</xsl:when>
								  <!--平衡-->
								  <xsl:when test="@rig='soft'">bright</xsl:when>
								  <!--柔和-->
								  <xsl:when test="@rig='contrasting'">bright</xsl:when>
								  <!--对比-->
								  <xsl:when test="@rig='sunrise'">bright</xsl:when>
								  <!--日出-->
								  <xsl:when test="@rig='flat'">bright</xsl:when>
								  <!--平面-->
								  <xsl:when test="@rig='glow'">bright</xsl:when>
								  <!--发光-->
								  <xsl:when test="@rig='brightRoom'">bright</xsl:when>
								  <!--明亮的房间-->
								  <!--以下效果转化为UOF中的normal效果-->
								  <xsl:when test="@rig='threePt'">normal</xsl:when>
								  <!--三点-->
								  <xsl:when test="@rig='flood'">normal</xsl:when>
								  <!--强烈-->
								  <xsl:when test="@rig='morning'">normal</xsl:when>
								  <!--早晨-->
								  <xsl:when test="@rig='chilly'">normal</xsl:when>
								  <!--寒冷-->
								  <xsl:when test="@rig='twoPt'">normal</xsl:when>
								  <!--两点-->
								  <!--以下效果转化为UOF中的dim效果-->
								  <xsl:when test="@rig='harsh'">dim</xsl:when>
								  <!--粗糙-->
								  <xsl:when test="@rig='sunset'">dim</xsl:when>
								  <!--日落-->
								  <xsl:when test="@rig='freezing'">dim</xsl:when>
								  <!--冰冻-->
								  <xsl:otherwise>normal</xsl:otherwise>
							  </xsl:choose>
						  </uof:强度_C63A>
					  </uof:照明_C638>
				  </xsl:for-each>
				  <uof:是否显示效果_C640>true</uof:是否显示效果_C640>
				  </图:三维效果_8061>
			  </xsl:if>
			 
	    
        <!--4月15日 黎美秀修改-->
        <xsl:if test="a:xfrm/@rot">
			<图:旋转角度_804D>
            <!--2011-5-26  luowentian-->
            <xsl:choose>
              <xsl:when test="(a:xfrm/@flipV and not(a:xfrm/@flipH)) or(a:xfrm/@flipH and not(a:xfrm/@flipV))">
                <xsl:value-of select="360-(a:xfrm/@rot div 60000)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="a:xfrm/@rot div 60000"/>
              </xsl:otherwise>
            </xsl:choose>
          </图:旋转角度_804D>
        
	</xsl:if>
     <!--liuyangyang 2014-12-10 添加blockarc旋转，修复控制点差异引起的图形不一致-->
     <xsl:if test="a:prstGeom[@prst = 'blockArc']/a:avLst/a:gd">
         <xsl:variable name="adj1">
           <xsl:value-of select="a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
         </xsl:variable>
         <xsl:variable name="tmp1">
           <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
         </xsl:variable>
         <xsl:variable name="x1">
           <xsl:value-of select="number($tmp1)"/>
         </xsl:variable>
         <xsl:variable name="adj2">
           <xsl:value-of select="a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
         </xsl:variable>
         <xsl:variable name="tmp2">
           <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
         </xsl:variable>
         <xsl:variable name="x2">
           <xsl:value-of select="number($tmp2)"/>
         </xsl:variable>
         <xsl:variable name="angle">
           <xsl:choose>
             <xsl:when test="$x1 &lt; $x2">
               <xsl:value-of select="($x1 - $x2) div 21600000 + 1"/>
             </xsl:when>
             <xsl:otherwise>
               <xsl:value-of select="($x1 - $x2) div 21600000"/>
             </xsl:otherwise>
           </xsl:choose>
         </xsl:variable>
       <xsl:variable name ="angle1" select="$x1 div 21600000 - $angle div 2 - 0.25"/>
       <图:旋转角度_804D>
         <xsl:choose>
         <xsl:when test="$angle1 &lt; 0">
           <xsl:value-of select="(1 + $angle1 ) * 360"/>
         </xsl:when>
           <xsl:when test="$angle1 &gt; 0">
             <xsl:value-of select=" $angle1  * 360"/>
           </xsl:when>
       </xsl:choose>
       </图:旋转角度_804D>
     </xsl:if>
     <!--end-->
        <xsl:if test="a:scene3d/a:camera/a:rot">
			<图:旋转角度_804D>
            <xsl:value-of select="(21600000-a:scene3d/a:camera/a:rot/@rev) div 60000"/>
          </图:旋转角度_804D>
        </xsl:if>
        
        <!--11.11.11 注释缩放比例 相对原始比例 李娟-->
			  
       
          <!--<图:X-缩放比例 uof:locID="g0026">1</图:X-缩放比例>     
         
          <图:Y-缩放比例 uof:locID="g0027">1</图:Y-缩放比例>-->
     <!--2013-1-1 liqiuling  修复bug锁定纵横比丢失  start -->    
        <xsl:choose>
          <xsl:when test="../p:nvSpPr/p:cNvSpPr/a:spLocks/@noChangeAspect='1'">
			  <图:缩放是否锁定纵横比_8055>true</图:缩放是否锁定纵横比_8055>
          </xsl:when>
          <xsl:when test="../p:nvGrpSpPr/p:cNvGrpSpPr/a:grpSpLocks/@noChangeAspect='1'">
            <图:缩放是否锁定纵横比_8055>true</图:缩放是否锁定纵横比_8055>
          </xsl:when>
          <xsl:otherwise>
			    <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
          </xsl:otherwise>
        </xsl:choose>
     <!--2013-1-1 liqiuling  修复bug锁定纵横比丢失  start -->
     
     <!--<图:相对原始比例 uof:locID="g0029">1</图:相对原始比例>-->

        <xsl:if test="name(..)='p:sp'">
          <xsl:if test="../p:nvSpPr/p:cNvPr/@descr">
			  <图:Web文字_804F>
              <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:cNvPr/@descr"/>
            </图:Web文字_804F>
          </xsl:if>
        </xsl:if>
        <!--2011-4-5罗文甜-->
        <xsl:if test="name(..)='p:pic'">
          <xsl:if test="../p:nvPicPr/p:cNvPr/@descr">
			  <图:Web文字_804F>
              <xsl:value-of select="ancestor::p:pic/p:nvPicPr/p:cNvPr/@descr"/>
            </图:Web文字_804F>
          </xsl:if>
        </xsl:if>
        <xsl:if test="name(..)='p:cxnSp'">
          <xsl:if test="../p:nvCxnSpPr/p:cNvPr/@descr">
			  <图:Web文字_804F>
              <xsl:value-of select="ancestor::p:cxnSp/p:nvCxnSpPr/p:cNvPr/@descr"/>
            </图:Web文字_804F>
          </xsl:if>
        </xsl:if>
        <xsl:if test="name(..)='p:grpSp'">
          <xsl:if test="../p:nvGrpSpPr/p:cNvPr/@descr">
			  <图:Web文字_804F>
              <xsl:value-of select="ancestor::p:grpSp/p:nvGrpSpPr/p:cNvPr/@descr"/>
            </图:Web文字_804F>
          </xsl:if>
        </xsl:if>

        <!--4月10日蒋俊彦修改-->
        <!--xsl:if test="./*/a:alpha">
					<图:透明度 uof:locID="g0038">
						<xsl:value-of select="substring(./*/a:alpha/@val,1,string-length(./*/a:alpha/@val)-3)"/>
					</图:透明度>
				</xsl:if-->

     <!--2013-11-20, tangjiang, Strict OOXML到UOF图形透明度转换 start -->
     <xsl:choose>
          <xsl:when test="a:grpFill">            
            <xsl:if test="ancestor::p:grpSp/p:grpSpPr//a:alpha or ancestor::p:grpSp/p:grpSpPr//a:alphaModFix">
              <图:透明度_8050>
              <xsl:choose>
                <xsl:when test="ancestor::p:grpSp/p:grpSpPr//a:alpha">
                  <xsl:variable name="grpSpPrVal" select="ancestor::p:grpSp/p:grpSpPr//a:alpha/@val" />
                  <xsl:if test="contains($grpSpPrVal,'%')">
                    <xsl:variable name="alphaVal" select="100-number(substring($grpSpPrVal,1,string-length($grpSpPrVal)-1))"/>
                    <xsl:choose>
                      <xsl:when test="$alphaVal &lt; 0">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$alphaVal"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <xsl:if test="not(contains($grpSpPrVal,'%'))">
                    <xsl:variable name="alphaVal" select="round(100 - ($grpSpPrVal div 1000))"/>
                    <xsl:choose>
                      <xsl:when test="$alphaVal &lt; 0">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$alphaVal"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="ancestor::p:grpSp/p:grpSpPr//a:alphaModFix">
                  <xsl:variable name="grpSpPrVal" select="ancestor::p:grpSp/p:grpSpPr//a:alphaModFix/@amt" />
                  <xsl:if test="contains($grpSpPrVal,'%')">
                    <xsl:variable name="alphaVal" select="100-number(substring($grpSpPrVal,1,string-length($grpSpPrVal)-1))"/>
                    <xsl:choose>
                      <xsl:when test="$alphaVal &lt; 0">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$alphaVal"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <xsl:if test="not(contains($grpSpPrVal,'%'))">
                    <xsl:variable name="alphaVal" select="round(100 - ($grpSpPrVal div 1000))"/>
                    <xsl:choose>
                      <xsl:when test="$alphaVal &lt; 0">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="$alphaVal"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                </xsl:when>
              </xsl:choose>
              </图:透明度_8050>
            </xsl:if>
          </xsl:when>
          <xsl:when test="a:solidFill//a:alpha or a:gradFill//a:alpha or a:pattFill//a:alpha or a:blipFill//a:alphaModFix">
              <图:透明度_8050>
                <xsl:choose>
                  <xsl:when test="a:solidFill//a:alpha">
                    <xsl:variable name="solidFillVal" select="a:solidFill//a:alpha/@val" />
                    <xsl:if test="contains($solidFillVal,'%')">
                      <xsl:variable name="alphaVal" select="100-number(substring($solidFillVal,1,string-length($solidFillVal)-1))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                    <xsl:if test="not(contains($solidFillVal,'%'))">
                      <xsl:variable name="alphaVal" select="round(100 - ($solidFillVal div 1000))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="a:gradFill//a:alpha">
                    <xsl:variable name="gradFillVal" select="a:gradFill//a:alpha/@val" />
                    <xsl:if test="contains($gradFillVal,'%')">
                      <xsl:variable name="alphaVal" select="100-number(substring($gradFillVal,1,string-length($gradFillVal)-1))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                    <xsl:if test="not(contains($gradFillVal,'%'))">
                      <xsl:variable name="alphaVal" select="round(100 - ($gradFillVal div 1000))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="a:pattFill//a:alpha">
                    <xsl:variable name="pattFillVal" select="a:pattFill//a:alpha/@val" />
                    <xsl:if test="contains($pattFillVal,'%')">
                      <xsl:variable name="alphaVal" select="100-number(substring($pattFillVal,1,string-length($pattFillVal)-1))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                    <xsl:if test="not(contains($pattFillVal,'%'))">
                      <xsl:variable name="alphaVal" select="round(100 - ($pattFillVal div 1000))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="a:blipFill//a:alphaModFix">
                    <xsl:variable name="blipFillVal" select="a:blipFill//a:alphaModFix/@amt" />
                    <xsl:if test="contains($blipFillVal,'%')">
                      <xsl:variable name="alphaVal" select="100-number(substring($blipFillVal,1,string-length($blipFillVal)-1))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                    <xsl:if test="not(contains($blipFillVal,'%'))">
                      <xsl:variable name="alphaVal" select="round(100 - ($blipFillVal/@amt div 1000))"/>
                      <xsl:choose>
                        <xsl:when test="$alphaVal &lt; 0">
                          <xsl:value-of select="'0'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$alphaVal"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:if>
                  </xsl:when>
                </xsl:choose>
              </图:透明度_8050>
          </xsl:when>
       <!--2014.10.30 liuyangyang 注释掉lummod和lumoff对透明度调节，lummod和lumoff与透明度不相关
         <xsl:otherwise>
            <xsl:if test="a:solidFill//a:lumMod or a:solidFill//a:lumOff ">
               <图:透明度_8050>
                 <xsl:choose>
                 <xsl:when test="a:solidFill//a:lumOff">
                   <xsl:variable name="solidFillOffVal" select="a:solidFill//a:lumOff/@val" />
                   <xsl:if test="contains($solidFillOffVal,'%')">
                     <xsl:variable name="alphaVal" select="number(substring($solidFillOffVal,1,string-length($solidFillOffVal)-1))"/>
                     <xsl:choose>
                       <xsl:when test="$alphaVal &lt; 0 or $alphaVal = 0">
                         <xsl:value-of select="'0'"/>
                       </xsl:when>
                       <xsl:otherwise>
                         <xsl:value-of select="round(100-$alphaVal)"/>
                       </xsl:otherwise>
                     </xsl:choose>
                   </xsl:if>
                   <xsl:if test="not(contains($solidFillOffVal,'%'))">
                     <xsl:variable name="alphaVal" select="round(($solidFillOffVal div 1000))"/>
                     <xsl:choose>
                       <xsl:when test="$alphaVal &lt; 0 or $alphaVal = 0">
                         <xsl:value-of select="'0'"/>
                       </xsl:when>
                       <xsl:otherwise>
                         <xsl:value-of select="round(100-$alphaVal)"/>
                       </xsl:otherwise>
                     </xsl:choose>
                   </xsl:if>
                 </xsl:when>
                   <xsl:otherwise>
                     <xsl:variable name="solidFillModVal" select="a:solidFill//a:lumMod/@val" />
                     <xsl:if test="contains($solidFillModVal,'%')">
                       <xsl:variable name="alphaVal" select="number(substring($solidFillModVal,1,string-length($solidFillModVal)-1))"/>
                       <xsl:choose>
                         <xsl:when test="$alphaVal &lt; 0 or $alphaVal = 0">
                           <xsl:value-of select="'0'"/>
                         </xsl:when>
                         <xsl:otherwise>
                           <xsl:value-of select="round(100-$alphaVal)"/>
                         </xsl:otherwise>
                       </xsl:choose>
                     </xsl:if>
                     <xsl:if test="not(contains($solidFillModVal,'%'))">
                       <xsl:variable name="alphaVal" select="round(($solidFillModVal div 1000))"/>
                       <xsl:choose>
                         <xsl:when test="$alphaVal &lt; 0 or $alphaVal = 0">
                           <xsl:value-of select="'0'"/>
                         </xsl:when>
                         <xsl:otherwise>
                           <xsl:value-of select="round(100-$alphaVal)"/>
                         </xsl:otherwise>
                       </xsl:choose>
                     </xsl:if>
                   </xsl:otherwise>
                 </xsl:choose>
               </图:透明度_8050>
            </xsl:if>
          </xsl:otherwise>
     end 2013-11-20, tangjiang, Strict OOXML到UOF图形透明度转换 -->

        </xsl:choose>
     
        <!--2010-12-15罗文甜：增加阴影对应-->
			  <!--11.11.28 添加内部阴影 李娟-->
     <!--2013-1-4 组合图形阴影丢失 liqiuling start-->
        <xsl:if test="a:effectLst/a:outerShdw|a:effectLst/a:innerShdw|ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw|ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:innerShdw">
          <xsl:call-template name="shdwCommen"/>            
        </xsl:if>
     <!--修改图形阴影效果丢失  2013-03-08 liqiuling start-->
     <xsl:if test="following-sibling::p:style/a:effectRef">
       <xsl:if test="following-sibling::p:style/a:effectRef/@idx">
         <xsl:variable name="idx" select="following-sibling::p:style/a:effectRef/@idx"/>
         <xsl:for-each select="//a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle">
           <xsl:if test="position()=$idx">
             <xsl:if test="a:effectLst/a:outerShdw">
               <xsl:call-template name="shdwCommen"/>
             </xsl:if>
           </xsl:if>
         </xsl:for-each>
       </xsl:if>
     </xsl:if>
     <!--修改图形阴影效果丢失  2013-03-08 liqiuling end-->
     <!--2013-1-4 组合图形阴影丢失 liqiuling end-->
	     <xsl:if test="parent::p:pic">
				  <图:图片属性_801E>
					  <!--颜色在uof中有四个枚举值：auto(automatic)，greyscale(greyscale)，monochrome(blackwhite)，erosion(watermark)。OOXML中枚举值很多，所以部分颜色值无法对应-->
            <!--2013-1-7 图片颜色转换不一致 liqiuling start biLevel-->
					  <图:颜色模式_801F>
						  <xsl:choose>
							  <xsl:when test="parent::p:pic/p:blipFill/a:blip/a:grayscl">
								  <xsl:value-of select="'greyscale'"/>
							  </xsl:when>
							  <xsl:when test="parent::p:pic/p:blipFill/a:blip/a:lum">
								  <xsl:value-of select="'erosion'"/>
							  </xsl:when>
							  <xsl:when test="parent::p:pic/p:blipFill/a:blip/a:biLevel">
								  <xsl:value-of select="'monochrome'"/>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:value-of select="'auto'"/>
							  </xsl:otherwise>
						  </xsl:choose>
					  </图:颜色模式_801F>
            <!--2013-1-7 图片颜色转换不一致 liqiuling start-->
					  <xsl:choose>
						  <xsl:when test=" parent::p:pic/p:blipFill/a:blip/a:lum/@bright &gt;= 0">
							  <图:亮度_8020>
								  <xsl:value-of select=" parent::p:pic/p:blipFill/a:blip/a:lum/@bright div 1000"/>
							  </图:亮度_8020>
						  </xsl:when>
						  <xsl:when test=" parent::p:pic/p:blipFill/a:blip/a:extLst/a:ext//a14:brightnessContrast/@bright &gt;= 0">
							  <图:亮度_8020>
								  <xsl:value-of select=" parent::p:pic/p:blipFill/a:blip/a:extLst/a:ext//a14:brightnessContrast/@bright div 1000"/>
							  </图:亮度_8020>
						  </xsl:when>
						  <xsl:otherwise>
							  <图:亮度_8020>50</图:亮度_8020>
						  </xsl:otherwise>
					  </xsl:choose>
					  <xsl:choose>
						  <xsl:when test="parent::p:pic/p:blipFill/a:blip/a:lum/@contrast &gt;= 0">
							  <图:对比度_8021>
								  <xsl:value-of select="parent::p:pic/p:blipFill/a:blip/a:lum/@contrast div 1000"/>
							  </图:对比度_8021>
						  </xsl:when>
						  <xsl:when test="parent::p:pic/p:blipFill/a:blip/a:extLst/a:ext//a14:brightnessContrast/@contrast &gt;= 0">
							  <图:对比度_8021>
								  <xsl:value-of select="parent::p:pic/p:blipFill/a:blip/a:extLst/a:ext//a14:brightnessContrast/@contrast div 1000"/>
							  </图:对比度_8021>
						  </xsl:when>
						  <xsl:otherwise>
							  <图:对比度_8021>50</图:对比度_8021>
						  </xsl:otherwise>
					  </xsl:choose>
            <!--2012-1-10 增加图片剪裁 liqiuling start   只是缩小没有剪裁 -->
            <xsl:if test="parent::p:pic/p:blipFill/a:srcRect">
              <图:图片裁剪_8022>
                <图:上_8023>
                  <xsl:choose>
                    <xsl:when test="parent::p:pic/p:blipFill/a:srcRect/@t">
                      <xsl:value-of select ="substring-before((parent::p:pic/p:blipFill/a:srcRect/@t),'%') div 1000"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </图:上_8023>
                <图:下_8024>
                  <xsl:choose>
                    <xsl:when test="parent::p:pic/p:blipFill/a:srcRect/@b">
                      <xsl:value-of select ="substring-before((parent::p:pic/p:blipFill/a:srcRect/@b),'%') div 1000"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </图:下_8024>
                <图:左_8025>
                  <xsl:choose>
                    <xsl:when test="parent::p:pic/p:blipFill/a:srcRect/@l">
                      <xsl:value-of select ="substring-before((parent::p:pic/p:blipFill/a:srcRect/@l),'%') div 1000"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </图:左_8025>
                <图:右_8026>
                  <xsl:choose>
                    <xsl:when test="parent::p:pic/p:blipFill/a:srcRect/@r">
                      <xsl:value-of select ="substring-before((parent::p:pic/p:blipFill/a:srcRect/@r),'%') div 1000"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="'0'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </图:右_8026>
              </图:图片裁剪_8022>
            </xsl:if>
            <!--2012-1-10 增加图片剪裁 liqiuling start-->
				  </图:图片属性_801E>
			  </xsl:if>
    
   </图:属性_801D>
      <!--2010-12-08罗文甜：增加连接线规则-->
      <xsl:if test="../p:nvCxnSpPr/p:cNvCxnSpPr">
        <xsl:call-template name="lineRegulrCommen"/>
      </xsl:if>
    </图:预定义图形_8018>
  </xsl:template>

	<!--<xsl:template match="wps:spPr/a:effectLst/a:innerShdw" mode="shadow">
		<图:阴影_8051 是否显示阴影_C61C="true" 类型_C61D="single">

			<xsl:choose>
				<xsl:when test="a:srgbClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:prstClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:prstClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:schemeClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:schemeClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="'auto'"/>
					</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:if test="//a:alpha">
				<xsl:choose>
					--><!--<xsl:value-of select="round((//a:alpha/@val div (1000 * 100))*256)"/>--><!--

					<xsl:when test="//a:alpha/@val &lt; 80000">                                                                                                                                                                                                                                                
						<xsl:attribute name="透明度_C61F">50</xsl:attribute>
					</xsl:when>
					<xsl:when test="//a:alpha/@val &gt; 80000">
						<xsl:attribute name="透明度_C61F">100</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</xsl:if>

			<uof:偏移量_C61B>
				<xsl:attribute name="x_C606">
					<xsl:value-of select="-6.0"/>
				</xsl:attribute>
				<xsl:attribute name="y_C607">                                                                              
					<xsl:value-of select="-6.0"/>
				</xsl:attribute>
			</uof:偏移量_C61B>

		</图:阴影_8051>
	</xsl:template>-->
  <!--2011-1罗文甜：关键点坐标-->
  <xsl:template name="keyCorCommen">

    <!--2014-05-05, tangjiang, 修复多路径转换 start -->
    <!--<图:关键点坐标 uof:locID="g0009" uof:attrList="路径">-->
	  <!--<图:路径_801C>-->
	  <图:路径值_8069>
      
      <!--2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 start -->
      <xsl:if test="./@w">
        <xsl:attribute name="宽度">
          <xsl:value-of select="./@w"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./@h">
        <xsl:attribute name="高度">
          <xsl:value-of select="./@h"/>
        </xsl:attribute>
      </xsl:if>
      <!--end 2014-04-27, tangjiang, 修复互操作之后组合图形丢失，这里的宽度、高度属性并不符合UOF标准，临时这样处理有待完善 -->
      
        <xsl:if test="./a:moveTo">
          <xsl:variable name="xStart">
            <xsl:choose>
              <xsl:when test="./a:moveTo/a:pt/@x!='0'">
                <xsl:value-of select="./a:moveTo/a:pt/@x div 12700 * 1.34"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'0'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="yStart">
            <xsl:choose>
              <xsl:when test="./a:moveTo/a:pt/@y!='0'">
                <xsl:value-of select="./a:moveTo/a:pt/@y div 12700 * 1.34"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'0'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="Start" select="concat('M',' ',$xStart,' ',$yStart)"/>
          <xsl:value-of select="concat($Start,' ')"/>
        </xsl:if>
        <xsl:for-each select="./a:cubicBezTo|./a:lnTo">
          <xsl:if test="count(a:pt)=3">
            <xsl:variable name="xCur1">
              <xsl:choose>
                <xsl:when test="./a:pt[1]/@x!='0'">
                  <xsl:value-of select="./a:pt[1]/@x div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="yCur1">
              <xsl:choose>
                <xsl:when test="./a:pt[1]/@y!='0'">
                  <xsl:value-of select="./a:pt[1]/@y div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="xCur2">
              <xsl:choose>
                <xsl:when test="./a:pt[2]/@x!='0'">
                  <xsl:value-of select="./a:pt[2]/@x div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="yCur2">
              <xsl:choose>
                <xsl:when test="./a:pt[2]/@y!='0'">
                  <xsl:value-of select="./a:pt[2]/@y div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="xCur3">
              <xsl:choose>
                <xsl:when test="./a:pt[3]/@x!='0'">
                  <xsl:value-of select="./a:pt[3]/@x div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="yCur3">
              <xsl:choose>
                <xsl:when test="./a:pt[3]/@y!='0'">
                  <xsl:value-of select="./a:pt[3]/@y div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select ="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of  select="concat('C',' ',$xCur1,' ',$yCur1,' ',$xCur2,' ',$yCur2,' ',$xCur3,' ',$yCur3,' ')"/>
            <!--/xsl:for-each-->
          </xsl:if>
          <xsl:if test="count(a:pt)=1">
            <xsl:variable name="xLn">
              <xsl:choose>
                <xsl:when test="./a:pt/@x!='0'">
                  <xsl:value-of select="./a:pt/@x div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="yLn">
              <xsl:choose>
                <xsl:when test="./a:pt/@y!='0'">
                  <xsl:value-of select="./a:pt/@y div 12700 * 1.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:value-of select="concat('L',' ',$xLn,' ',$yLn,' ')"/>
          </xsl:if>
        </xsl:for-each>
	  </图:路径值_8069>
    <!--</图:路径_801C>-->
    <!--</图:关键点坐标>-->
    <!--end 2014-05-05, tangjiang, 修复多路径转换 -->

  </xsl:template>
  
  <!--2014-04-16, tangjiang, 修复连线规则互操作转换需打开修复 start -->
  <!--2010-12-08罗文甜：增加连接线规则-->
  <xsl:template name="lineRegulrCommen">
	  <图:连接线规则_8027>
      <xsl:if test="../p:nvCxnSpPr/p:cNvPr/@id">
        <xsl:attribute name="连接线引用_8028">
          <xsl:value-of  select="../@id"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="../p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn">
        <xsl:attribute name="始端对象引用_8029">
          <xsl:variable name="CxnSTid" select="../p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@id"/>
          <xsl:value-of select="../../p:sp[p:nvSpPr/p:cNvPr/@id=$CxnSTid]/@id"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="../p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn">
        <xsl:attribute name="终端对象引用_802A">
          <xsl:variable name="CxnENDid" select="../p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@id"/>
          <xsl:value-of select="../../p:sp[p:nvSpPr/p:cNvPr/@id=$CxnENDid]/@id"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="../p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@idx">
          <xsl:attribute name="始端对象连接点索引_802B">
            <xsl:value-of select="../p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn/@idx"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="始端对象连接点索引_802B">
            <xsl:value-of select="'-1'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="../p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@idx">
          <xsl:attribute name="终端对象连接点索引_802C">
            <xsl:value-of select="../p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn/@idx"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="终端对象连接点索引_802C">
            <xsl:value-of select="'-1'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </图:连接线规则_8027>
  </xsl:template>
  <!--end 2014-04-16, tangjiang, 修复连线规则互操作转换需打开修复 -->
  
  <!--2010-12-15罗文甜：增加阴影对应，内部阴影不能对应，UOF中没有-->
  <!--2013-1-4 组合图形阴影丢失 liqiuling start-->
  <xsl:template name="shdwCommen">
    <xsl:if test="a:effectLst/a:outerShdw|ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw">
		<图:阴影_8051>
			<xsl:attribute name="是否显示阴影_C61C">true</xsl:attribute>
      <!--<图:阴影 uof:locID="g0048" uof:attrList="是否显示阴影 阴影类型 X偏移量 Y偏移量 阴影颜色 阴影透明度" uof:是否显示阴影="true">-->
        <!--xsl:variable name="ang" select="a:effectLst/a:outerShdw/dir div 60000"/-->
        <xsl:choose>
          <xsl:when test="a:effectLst/a:outerShdw/@kx='-1200000'or ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/@kx='-1200000'">
            <xsl:attribute name="类型_C61D">3</xsl:attribute>
          </xsl:when>
          <xsl:when test="a:effectLst/a:outerShdw/@kx='1200000' or ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/@kx='1200000'">
            <xsl:attribute name="类型_C61D">2</xsl:attribute>
          </xsl:when>
          <xsl:when test="a:effectLst/a:outerShdw/@kx='-800400'or ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/@kx='-800400'">
            <xsl:attribute name="类型_C61D">7</xsl:attribute>
          </xsl:when>
          <xsl:when test="a:effectLst/a:outerShdw/@kx='800400' or ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/@kx='800400'">
            <xsl:attribute name="类型_C61D">6</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="类型_C61D">13</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
		
        <xsl:choose>
          <xsl:when test="a:effectLst/a:outerShdw/a:srgbClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:value-of select="concat('#',a:effectLst/a:outerShdw/a:srgbClr/@val)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:effectLst/a:outerShdw/a:prstClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:apply-templates select="a:effectLst/a:outerShdw/a:prstClr"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="a:effectLst/a:outerShdw/a:schemeClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:apply-templates select="a:effectLst/a:outerShdw/a:schemeClr"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:srgbClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:value-of select="concat('#',ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:srgbClr/@val)"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:prstClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:apply-templates select="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:prstClr"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:schemeClr">
            <xsl:attribute name="颜色_C61E">
              <xsl:apply-templates select="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/a:schemeClr"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="颜色_C61E">
              <xsl:value-of select="'auto'"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      <xsl:choose>
        <!--2013-01-09 修改阴影的透明度  李秋玲 start-->
         <!-- <xsl:when test="a:effectLst/a:outerShdw//a:alpha">
         <xsl:attribute name="透明度_C61F">
            <xsl:value-of select="round((a:effectLst/a:outerShdw//a:alpha/@val div (1000 * 100))*256)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw//a:alpha">
          <xsl:attribute name="透明度_C61F">
            <xsl:value-of select="round((ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw//a:alpha/@val div (1000 * 100))*256)"/>
          </xsl:attribute>
        </xsl:when>-->
        <xsl:when test="a:effectLst/a:outerShdw//a:alpha">
          <xsl:attribute name="透明度_C61F">
            <xsl:value-of select="substring-before(a:effectLst/a:outerShdw//a:alpha/@val, '%')"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw//a:alpha">
          <xsl:attribute name="透明度_C61F">
            <xsl:value-of select="substring-before(ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw//a:alpha/@val, '%')"/>
          </xsl:attribute>
        </xsl:when>
        <!--2013-01-09 修改阴影的透明度  李秋玲 end-->
      </xsl:choose>

      <!--2014-03-17, tangjiang, 修复Strict OOXML到UOF 偏移量的指数转换 start -->
			<uof:偏移量_C61B>
        <xsl:choose>
          <xsl:when test="a:effectLst/a:outerShdw/x">
            <xsl:variable name="xVal" select="a:effectLst/a:outerShdw/x"/>
            <xsl:attribute name="x_C606">
              <!-- 指数处理,这里均按‘E-6’和‘E+6’处理，有待完善 -->
            <xsl:choose>
              <xsl:when test="contains( $xVal, 'E-' )">
                <xsl:value-of select="round(substring-before($xVal, 'E-') div 1000000 )"/>
              </xsl:when>
              <xsl:when test="contains( $xVal, 'E+' )">
                <xsl:value-of select="round(substring-before($xVal, 'E+') * 1000000 )"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="round( $xVal )"/>
              </xsl:otherwise> 
            </xsl:choose>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/x">
            <xsl:variable name="xVal" select="ancestor::p:grpSp/p:grpSpPr/a:effectLst/a:outerShdw/x"/>
            <xsl:attribute name="x_C606">
              <xsl:choose>
                <xsl:when test="contains( $xVal, 'E-' )">
                  <xsl:value-of select="round(substring-before($xVal, 'E-') div 1000000 )"/>
                </xsl:when>
                <xsl:when test="contains( $xVal, 'E+' )">
                  <xsl:value-of select="round(substring-before($xVal, 'E+') * 1000000 )"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round( $xVal )"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        
        <xsl:choose>
          <xsl:when test="a:effectLst/a:outerShdw/y">
            <xsl:variable name="yVal" select="a:effectLst/a:outerShdw/y"/>
            <xsl:attribute name="y_C607">
              <xsl:choose>
                <xsl:when test="contains( $yVal, 'E-' )">
                  <xsl:value-of select="round(substring-before($yVal, 'E-') div 1000000 )"/>
                </xsl:when>
                <xsl:when test="contains( $yVal, 'E+' )">
                  <xsl:value-of select="round(substring-before($yVal, 'E+') * 1000000 )"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round( $yVal )"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:grpSp/p:grpSpPr/a:effectLst/y">
            <xsl:variable name="yVal" select="ancestor::p:grpSp/p:grpSpPr/a:effectLst/y"/>
            <xsl:attribute name="y_C607">
              <xsl:choose>
                <xsl:when test="contains( $yVal, 'E-' )">
                  <xsl:value-of select="round(substring-before($yVal, 'E-') div 1000000 )"/>
                </xsl:when>
                <xsl:when test="contains( $yVal, 'E+' )">
                  <xsl:value-of select="round(substring-before($yVal, 'E+') * 1000000 )"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round( $yVal )"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
			</uof:偏移量_C61B>
      <!-- end 2013-11-20, tangjiang -->
      
	  </图:阴影_8051>
    </xsl:if>
     </xsl:template>
  <!--2013-1-4 组合图形阴影丢失 liqiuling end-->
  <!--2010.04.29 myx add -->
	<!--先注销这块李娟2012 03.07-->
  <!--<xsl:template match="a:ln" mode="linecolor">
	  <图:线_8057>
	  <xsl:choose>
      <xsl:when test="a:solidFill">
		  <图:线颜色_8058>
          <xsl:apply-templates select="a:solidFill" mode="ln"/>
        </图:线颜色_8058>
      </xsl:when>
      
      <xsl:when test="a:pattFill">
		  <图:线颜色_8058>
          <xsl:call-template name="colorChoice"/>
        </图:线颜色_8058>
      </xsl:when>
     
      <xsl:when test="a:gradFill">
		  <图:线颜色_8058>
          <xsl:apply-templates select="a:gradFill/a:gsLst/a:gs" mode="firstClr"/>
        </图:线颜色_8058>
      </xsl:when>
    </xsl:choose>
	  </图:线_8057>
  </xsl:template>-->
  
  
  <xsl:template match="a:ln">
	 
    <xsl:choose>
      <!-- 2010.04.29 myx add -->
      <xsl:when test="a:solidFill | a:pattFill |a:gradFill">
		
        <xsl:apply-templates select="." mode="ln"/>
		 
      </xsl:when>
      <xsl:otherwise>
        <!--2011-1-14罗文甜，修改颜色bug-->
        <xsl:choose>
          <!--xsl:when test="parent::p:spPr/parent::p:sp/p:style/a:lnRef/*"-->
          <xsl:when test="../../p:style/a:lnRef/*">
            <xsl:for-each select="../../p:style/a:lnRef">
				<图:线颜色_8058>
                <xsl:call-template name="colorChoice"/>
              </图:线颜色_8058>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="../../p:style/a:lnRef/@idx">
              <xsl:variable name="idx" select ="../../p:style/a:lnRef/@idx"/>
              <xsl:apply-templates select="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]"
                                    mode="linecolor"/>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
    <!--10.10.8 罗文甜 修改线型虚实对应-->
    <!--4月15日蒋俊彦改-->
    <xsl:if test="current()">
		<图:线类型_8059>
        <xsl:attribute name="线型_805A">
          <xsl:choose>
            <xsl:when test="@cmpd='sng'">single</xsl:when>
            <xsl:when test="@cmpd='dbl'">double</xsl:when>
            <xsl:when test="@cmpd='thickThin'">thin-thick</xsl:when>
            <xsl:when test="@cmpd='thinThick'">thick-thin</xsl:when>
            <xsl:when test="@cmpd='tri'">thick-between-thin</xsl:when>
            <xsl:otherwise>single</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:choose>
          <xsl:when test="a:prstDash/@val">
            <xsl:attribute name="虚实_805B">
              <xsl:choose>
                <xsl:when test="a:prstDash/@val='solid'">
                  <xsl:value-of select="'solid'"/>
                </xsl:when>
                <!--圆点线-->
                <xsl:when test="a:prstDash/@val='sysDot'">
                  <xsl:value-of select="'round-dot'"/>
                </xsl:when>
                <!--方点线-->
                <xsl:when test="a:prstDash/@val='sysDash'">
                  <xsl:value-of select="'square-dot'"/>
                </xsl:when>
                <!--短划线->虚线-->
                <xsl:when test="a:prstDash/@val='dash'">
                  <xsl:value-of select="'dash'"/>
                </xsl:when>
                <!--划线-点->点虚线-->
                <xsl:when test="a:prstDash/@val='dashDot'">
                  <xsl:value-of select="'dash-dot'"/>
                </xsl:when>
                <!--长划线->长虚线-->
                <xsl:when test="a:prstDash/@val='lgDash'">
                  <xsl:value-of select="'long-dash'"/>
                </xsl:when>
                <!--长划线-点-->
                <xsl:when test="a:prstDash/@val='lgDashDot'">
                  <xsl:value-of select="'long-dash-dot'"/>
                </xsl:when>
                <!--长划线-点-点-->
                <xsl:when test="a:prstDash/@val='lgDashDotDot'">
                  <xsl:value-of select="'dash-dot-dot'"/>
                </xsl:when>
                <!--其它情况处理>
                <xsl:when test="a:prstDash/@val='sysDashDot'">
                  <xsl:value-of select="'dash-dot-heavy'"/>
                </xsl:when>
                <xsl:when test="a:prstDash/@val='dot'">
                  <xsl:value-of select="'dotted'"/>
                </xsl:when>
                <xsl:when test="a:prstDash/@val='sysDashDotDot'">
                  <xsl:value-of select="'dash-dot-dot-heavy'"/>
                </xsl:when-->
              </xsl:choose>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="虚实_805B">solid</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>  
    </图:线类型_8059>
    </xsl:if>   
    <xsl:if test="@w">
		<图:线粗细_805C>
        <xsl:value-of select="@w div 12700"/>
      </图:线粗细_805C>
    </xsl:if>
    <!-- 2010.04.28 马有旭-->
    <xsl:if test="not(@w)">     
      <xsl:choose>
        <xsl:when test="../../p:style/a:lnRef/@idx">
          <xsl:variable name="idx" select="../../p:style/a:lnRef/@idx"/>
          <xsl:choose>
            <xsl:when test="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]/@w">
				<图:线粗细_805C>
                <xsl:value-of select="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]/@w div 12700"/>
              </图:线粗细_805C>
            
		</xsl:when>
            <xsl:otherwise>
				<图:线粗细_805C>2</图:线粗细_805C>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
			<图:线粗细_805C>2</图:线粗细_805C>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <!--4月13日改线型箭头-->
    <!--09.12.02 马有旭<xsl:if test="a:headEnd">-->
	
</xsl:template>
	<!--<xsl:template match="a:solidFill" mode="linecolor">
		<图:线颜色_8058>
			<xsl:variable name="col">
				<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
			</xsl:variable>
			<xsl:value-of select="$col"/>
		</图:线颜色_8058>
	</xsl:template>-->
  <!--对style的处理 add by linyaohu-->
	<!--线条颜色,只需得到颜色值-->
	<xsl:template match="a:solidFill" mode="linecolor">
		
		<xsl:variable name="lncolor">
			<xsl:choose>
				<xsl:when test="a:srgbClr">
					<xsl:value-of select="a:srgbClr/@val"/>
				</xsl:when>
				<xsl:when test="a:hslClr">
					<xsl:value-of select="a:hslClr/@val"/>
				</xsl:when>
				<xsl:when test="a:prstClr">
					<xsl:value-of select="a:prstClr/@val"/>
				</xsl:when>
				<xsl:when test="a:schemeClr">
					<xsl:value-of select="a:schemeClr/@val"/>
				</xsl:when>
				<xsl:when test="a:scrgbClr">
					<xsl:value-of select="a:scrgbClr/@val"/>
				</xsl:when>
				<xsl:when test="a:sysClr">
					<xsl:value-of select="a:sysClr/@val"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<图:线颜色_8058>
			<xsl:value-of select="concat('#',$lncolor)"/>
		</图:线颜色_8058>
	</xsl:template>
	<xsl:template match="a:gradFill" mode="ln">
		<图:线颜色_8058>auto</图:线颜色_8058>
	</xsl:template>
  <xsl:template match="a:lnRef">
    <!--2010-11-19 罗文甜 增加判断-->
	  <图:线_8057>
    <xsl:choose>
      <xsl:when test="./*">
        <图:线颜色_8058>
          <xsl:call-template name="colorChoice"/>
        </图:线颜色_8058>
      </xsl:when>
      <xsl:otherwise>
        <!-- 2010.04.29 myx add -->
        <xsl:variable name="idx" select="@idx"/>
        <xsl:for-each select="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln">
          <xsl:if test="position()=$idx">
            <xsl:apply-templates select="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
    <!--2010-0-7罗文甜，修改预定义图形边框线粗细BUG-->
    <xsl:if test="./@idx">
      <xsl:variable name="idx" select="./@idx"/>
      <xsl:choose>
        <xsl:when test="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]/@w">
			<图:线粗细_805C>
            <xsl:value-of select="//a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]/@w div 12700"/>
          </图:线粗细_805C>
        </xsl:when>
        <xsl:otherwise>
			<图:线粗细_805C>2</图:线粗细_805C>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
	  </图:线_8057>
  </xsl:template>
  <!--修改bug图形背景色转换不正确 2013-03-08；预定义图形背景转换不正确 2013-03-19；图形填充不正确2014-03-20 liqiuling start-->
  <xsl:template match="a:fillRef">
    <xsl:if test="./@idx">
      <xsl:variable name="idx" select="./@idx"/>
     <xsl:choose>
        <xsl:when test="$idx=1 or $idx=2 or $idx=3">
          <图:填充_804C>
            <图:颜色_8004>
              <xsl:apply-templates select="a:schemeClr"/>  
            </图:颜色_8004>
          </图:填充_804C>
           </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="//a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/*">
            <xsl:if test="position()=$idx">
              <图:填充_804C>
                <xsl:apply-templates select="."/>
              </图:填充_804C>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
      </xsl:if>
<!--
	  <图:填充_804C>
		  <图:颜色_8004>
        <xsl:call-template name="colorChoice"/>
      </图:颜色_8004>
    </图:填充_804C>-->
  </xsl:template>
  <!--修改bug图形背景色转换不正确 2013-03-08 liqiuling end-->
  <xsl:template match="p:txBody|dsp:txBody">
	  <图:文本_803C>
    <!--<图:文本内容 uof:locID="g0002" uof:attrList="文本框 左边距 右边距 上边距 下边距 水平对齐 垂直对齐 文字排列方向 自动换行 大小适应文字 前一链接 后一链接">-->
      <xsl:apply-templates select="a:bodyPr"/>
		  <图:内容_8043>
      <xsl:apply-templates select="a:p"/>
			  <!--<xsl:if test =".//a:tbl">
				  <xsl:apply-templates select=".//a:tbl"/>
			  </xsl:if>-->
			  </图:内容_8043>
    <!--</图:文本内容>-->
		  </图:文本_803C>
  </xsl:template>
  <xsl:template match="a:bodyPr">
	  <!--<图:文本_803C>-->
    <xsl:attribute name="是否为文本框_8046">
      <xsl:value-of select="'true'"/>
    </xsl:attribute>
	  <xsl:attribute name="是否自动换行_8047">
		  <xsl:if test="./@wrap='none'">
			  <xsl:value-of select="'false'"/>
		  </xsl:if>
		  <xsl:if test="./@wrap='square' or name(./@wrap)=''">
			  <xsl:value-of select="'true'"/>
		  </xsl:if>
	  </xsl:attribute>
	  <xsl:attribute name="是否大小适应文字_8048">
		  <!--	<xsl:if test="name(./a:spAutoFit) !=''or name(./a:normAutofit)!=''">	
			  <xsl:value-of select="'true'"/>
			</xsl:if>
			<xsl:if test="name(./a:spAutoFit) =''and name(./a:normAutofit)='' or name(./a:noAutofit) !=''">
				<xsl:value-of select="'false'"/>
			</xsl:if>前一链接后一链接还没转-->
		  <xsl:if test="not(name(./a:spAutoFit)) or a:noAutofit">
			  <xsl:value-of select="'false'"/>
		  </xsl:if>
		  <xsl:if test="./a:spAutoFit">
			  <xsl:value-of select="'true'"/>
		  </xsl:if>
	  </xsl:attribute>
	  <!--属性:是否文字随对象旋转,是否文字随边框自动缩放_804A,是否文字绕排外部对象_8068没有添加 修改 下面属性变为元素 李娟 11.11.11 -->
    <!--liuyangyang 2015-04-01 添加文字随图形旋转 start-->
    <xsl:attribute name="是否文字随对象旋转_8049">
      <xsl:value-of select="'true'"/>
    </xsl:attribute>
    <!--end liuyangyang 2015-04-01 添加文字随图形旋转-->
    <!--2014-03-09, pengxin, 修改Strict OOXML到UOF图形文本框过大的错误和图形内部文本中边距转换有误 start -->
    <图:边距_803D>
      <xsl:choose>
        <xsl:when test="./@lIns">
          <xsl:variable name="lInsStr" select="./@lIns"/>
          <xsl:attribute name="左_C608">
            <xsl:choose>
              <xsl:when test="string-length($lInsStr) >= 8 ">
                <xsl:value-of select="$lInsStr div 1270000"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$lInsStr div 12700"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="左_C608">
            <xsl:value-of select="90000 div 12700"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="./@rIns">
          <xsl:variable name="rInsStr" select="./@rIns"/>
          <xsl:attribute name="右_C60A">
            <xsl:choose>
              <xsl:when test="string-length($rInsStr) >= 8 ">
                <xsl:value-of select="$rInsStr div 1270000"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$rInsStr div 12700"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="右_C60A">
            <xsl:value-of select="90000 div 12700"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>  
            
      <xsl:choose>
        <xsl:when test="./@tIns">
          <xsl:variable name="tInsStr" select="./@tIns"/>
          <xsl:attribute name="上_C609">
            <xsl:choose>
              <xsl:when test="string-length($tInsStr) >= 8 ">
                <xsl:value-of select="$tInsStr div 1270000"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$tInsStr div 12700"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="上_C609">
            <xsl:value-of select="46800 div 12700"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>      
      
      <xsl:choose>
        <xsl:when test="./@bIns">
          <xsl:variable name="bInsStr" select="./@bIns"/>
          <xsl:attribute name="下_C60B">
            <xsl:choose>
              <xsl:when test="string-length($bInsStr) >= 8 ">
                <xsl:value-of select="$bInsStr div 1270000"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$bInsStr div 12700"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="下_C60B">
            <xsl:value-of select="46800 div 12700"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
	  </图:边距_803D>
    <!-- end 2014-03-09, pengxin -->
    
    <!--2013-1-15 修改文本框对齐不正确 2013-04-25  2013-5-19  liqiuling start -->
	  <图:对齐_803E>
      <xsl:variable name="phType" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
      <xsl:attribute name="水平对齐_421D">
        <xsl:choose>
          <xsl:when test=" $phType = 'title' ">
            <!-- 这里应该取a:pPr/@lvl对应的级别剧中，先按lvl1pPr处理 add by tangjiang 2014-03-16 -->
            <xsl:variable name="tileAlign" select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/@algn"/>
            <xsl:choose>
              <xsl:when test=" $tileAlign = 'l' ">
                <xsl:value-of select="'left'"/>
              </xsl:when>
              <xsl:when test=" $tileAlign = 'ctr' ">
                <xsl:value-of select="'center'"/>
              </xsl:when>
              <xsl:when test=" $tileAlign = 'r' ">
                <xsl:value-of select="'right'"/>
              </xsl:when>
              <xsl:when test=" $tileAlign = 'just' ">
                <xsl:value-of select="'auto'"/>
              </xsl:when>
              <xsl:when test=" $tileAlign = 'dist' ">
                <xsl:value-of select="'center'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'left'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test=" $phType = 'body' ">
            <!-- 这里应该取a:pPr/@lvl对应的级别剧中，先按lvl1pPr处理 add by tangjiang 2014-03-16 -->
            <xsl:variable name="bodyAlign" select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/@algn"/>
            <xsl:choose>
              <xsl:when test=" $bodyAlign = 'l' ">
                <xsl:value-of select="'left'"/>
              </xsl:when>
              <xsl:when test=" $bodyAlign = 'ctr' ">
                <xsl:value-of select="'center'"/>
              </xsl:when>
              <xsl:when test=" $bodyAlign = 'r' ">
                <xsl:value-of select="'right'"/>
              </xsl:when>
              <xsl:when test=" $bodyAlign = 'just' ">
                <xsl:value-of select="'auto'"/>
              </xsl:when>
              <xsl:when test=" $bodyAlign = 'dist' ">
                <xsl:value-of select="'center'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'left'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test=" $phType = 'dt' or $phType = 'ftr' or $phType = 'sldNum' ">
            <!-- 这里应该在p:txStyles/p:otherStyle节点下获得样式，先不做处理 add by tangjiang 2014-03-16-->
            <xsl:value-of select="'left'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'left'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--2013-1-15 修改文本框对齐不正确 liqiuling end-->
      <!--liuyin 20121120 修改bug 字体对齐方式不正确-->
    <xsl:if test="not(./@anchor) or ./@anchor">
      <!--end  liuyin 20121120 修改bug 字体对齐方式不正确-->
      <xsl:attribute name="文字对齐_421E">
        <xsl:choose>
          <xsl:when test="./@anchor='b'">
            <xsl:value-of select="'bottom'"/>
          </xsl:when>
          <xsl:when test="./@anchor='ctr' or ./@anchor='dist'">
            <xsl:value-of select="'center'"/>
          </xsl:when>
          <xsl:when test="./@anchor='t'">
            <xsl:value-of select="'top'"/>
          </xsl:when>
          <xsl:when test="./@anchor='just'">
            <xsl:value-of select="'auto'"/>
          </xsl:when>
          <!--2014-03-27, tangjiang, 修改bug 分别设置标题、正文等的默认值对齐方式 start -->
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test=" $phType = 'title' ">
                <xsl:value-of select="'center'"/>
              </xsl:when>
              <xsl:when test=" $phType = 'body' ">
                <xsl:value-of select="'top'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'top'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
          <!-- end 2014-03-27, tangjiang, 分别设置标题、正文等的默认值对齐方式 -->
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
	  </图:对齐_803E>
      <xsl:if test="./@vert">
		<图:文字排列方向_8042>
      <!--<xsl:attribute name="图:文字排列方向">-->
        <!--2010.10.25 罗文甜 修改文字排列方向 -->
        <xsl:choose>
          <xsl:when test="./@vert='horz' or not(./@vert)">
            <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
          </xsl:when>
          <xsl:when test="./@vert='eaVert'">
            <xsl:value-of select="'r2l-t2b-0e-90w'"/>
          </xsl:when>
          <xsl:when test="./@vert='vert'">
            <xsl:value-of select="'r2l-t2b-90e-90w'"/>
          </xsl:when>
          <!--2011-2-17罗文甜，修改文字堆积bug-->
          <xsl:when test="./@vert='vert270'">
            <xsl:value-of select="'l2r-b2t-270e-270w'"/>
          </xsl:when>
          <xsl:when test="./@vert='wordArtVertRtl'">
            <xsl:value-of select="'r2l-t2b-0e-90w'"/>
          </xsl:when>
          <!--liuyin 堆积，这么转换更合适啊
          <xsl:when test="./@veert='wordArtVert'">-->
          <xsl:when test="./@vert='wordArtVert'">
            <xsl:value-of select="'l2r-t2b-0e-90w'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'t2b-l2r-0e-0w'"/>
          </xsl:otherwise>
        </xsl:choose>

        <!--</xsl:attribute>-->
		</图:文字排列方向_8042>
    </xsl:if>
		  <!--<图:内容_8043>
			  <xsl:apply-templates select="a:p"/>
			</图:内容_8043>-->
		 
    <!--4.15黎美秀修改-->
	  <!--</图:文本_803C>-->
  </xsl:template>
   <!--2012-12-20, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
  <xsl:template name="dspgroupLocation">
    <xsl:param name="framexfrm"/>
    <图:组合位置_803B>
      <xsl:attribute name="x_C606">
        <xsl:variable name="xVal" select="( .//a:xfrm/a:off/@x + $framexfrm/@x ) div 12700"/>
        <xsl:choose>
          <xsl:when test="$xVal &lt; 0">
            <xsl:value-of select="'0'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$xVal"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="y_C607">
        <xsl:variable name="yVal" select="(  .//a:xfrm/a:off/@y + $framexfrm/@y ) div 12700"/>
        <xsl:choose>
          <xsl:when test="$yVal &lt; 0">
            <xsl:value-of select="'0'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$yVal"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </图:组合位置_803B>

  </xsl:template>
  <!--end-->
  <xsl:template name="groupLocation">
    <!--<图:组合位置 uof:locID="g0041" uof:attrList="x坐标 y坐标">
      <xsl:attribute name="图:x坐标">
        <xsl:value-of select="($basex+.//a:xfrm/a:off/@x) div 12700"/>
      </xsl:attribute>
      <xsl:attribute name="图:y坐标">
        <xsl:value-of select="($basey+.//a:xfrm/a:off/@y) div 12700"/>
      </xsl:attribute>
    </图:组合位置>-->
	  <图:组合位置_803B>
      <xsl:attribute name="x_C606">
        <xsl:variable name="xVal" select=".//a:xfrm/a:off/@x div 12700"/>
        <xsl:choose>
          <xsl:when test="$xVal &lt; 0">
            <xsl:value-of select="'0'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$xVal"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="y_C607">
        <xsl:variable name="yVal" select=".//a:xfrm/a:off/@y div 12700"/>
        <xsl:choose>
          <xsl:when test="$yVal &lt; 0">
            <xsl:value-of select="'0'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$yVal"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </图:组合位置_803B>
  
  </xsl:template>
  <!-- 09.11.09 马有旭 添加 -->
  <xsl:template match="p:grpSpPr" mode="colorNode">
    <xsl:choose>
      <xsl:when test="a:gradFill">
		  <图:填充_804C>
          <xsl:apply-templates select="a:gradFill"/>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:pattFill">
		  <图:填充_804C>
          <xsl:apply-templates select="a:pattFill"/>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:solidFill">
		  <图:填充_804C>
          <xsl:apply-templates select="a:solidFill"/>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:blipFill">
		  <图:填充_804C>
          <xsl:apply-templates select="a:blipFill"/>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:noFill"/>
      <!--有style的情况-->
      <xsl:when test="following-sibling::p:style/a:fillRef">
        <xsl:apply-templates select="following-sibling::p:style/a:fillRef"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="groupBaseSize">
  </xsl:template>
</xsl:stylesheet>
