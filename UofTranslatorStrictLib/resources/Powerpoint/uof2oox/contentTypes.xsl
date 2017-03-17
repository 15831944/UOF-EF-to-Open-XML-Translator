<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:对象="http://schemas.uof.org/cn/2009/objects"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:图="http://schemas.uof.org/cn/2009/graph">
  <xsl:template name="contentTypes">
	  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
		  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
		  <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="jpeg" ContentType="image/jpeg"/>
      <!--liuyangyang 2015.03.04 添加ole对象的contenttype start-->
      <xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/图:图形_8062/ole">
        <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
        <Default Extension="xlsm" ContentType="application/vnd.ms-excel.sheet.macroEnabled.12"/>
        <Default Extension="sldm" ContentType="application/vnd.ms-powerpoint.slide.macroEnabled.12"/>
        <Default Extension="xls" ContentType="application/vnd.ms-excel"/>
        <Default Extension="xlsb" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.12"/>
        <Default Extension="ppt" ContentType="application/vnd.ms-powerpoint"/>
        <Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>
        <Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
        <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
        <Default Extension="doc" ContentType="application/msword"/>
        <Default Extension="docm" ContentType="application/vnd.ms-word.document.macroEnabled.12"/>
        <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
        <Default Extension="sldx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide"/>
      </xsl:if>
        <!-- end liuyangyang 2015.03.04 添加ole对象的contenttype-->
      <!--下边两个未验证-->
		  <!--修改标签 李娟11.12.26-->
		  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
		  <xsl:for-each select=".//演:母版集_6C0C/演:母版_6C0D">
			  <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml">
				  <xsl:attribute name="PartName">
					  <xsl:value-of select="'/ppt/theme/theme'"/>
					  <!-- 09.10.12 马有旭 修改
          <xsl:value-of select="substring-after(@演:配色方案引用,'ID')"/>
          -->
					  <xsl:choose>
						  
							  <xsl:when test="@配色方案引用_6BEB">
								  <xsl:choose>
									  <xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
										<xsl:value-of select="substring-after(@配色方案引用_6BEB,'clr_')"/>
									</xsl:when>
									  <xsl:otherwise>
										 <xsl:value-of select="substring-after(@配色方案引用_6BEB,'Id')"/>
										</xsl:otherwise>
								  </xsl:choose>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:value-of select="substring-after(@标识符_6BE8,'ID')"/>
							  </xsl:otherwise>
						
						  
						  
					  </xsl:choose>
					  <xsl:value-of select="'.xml'"/>
				  </xsl:attribute>
			  </Override>
		  </xsl:for-each>


		  <!--picture-->
		  <!--修改标签 李娟 11.12.26-->
		  <xsl:if test=".//对象:对象数据_D701[@公共类型_D706='jpg'] or //对象:对象数据集_D700/对象:对象数据_D701[@私有类型_D707]">
			  <Default Extension="jpg" ContentType="image/jpg"/>
		  </xsl:if>
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='png']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:公共类型='png']">-->
			  <Default Extension="png" ContentType="image/png"/>
		  </xsl:if>
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='gif'] or .//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='GIF']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:公共类型='gif']">-->
			  <Default Extension="gif" ContentType="image/gif"/>
		  </xsl:if>
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='bmp']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:公共类型='bmp']">-->
			  <Default Extension="bmp" ContentType="image/bmp"/>
		  </xsl:if>
		  <!-- 09.12.18 马有旭 -->
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@私有类型_D707='wmf'] or .//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='wmf']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:私有类型='wmf'] or .//uof:其他对象[@uof:公共类型='wmf']">-->
			  <Default Extension="wmf" ContentType="image/x-wmf"/>
		  </xsl:if>
		  <!-- 09.12.21 马有旭 添加-->
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='emf'] or .//对象:对象数据集_D700/对象:对象数据_D701[@私有类型_D707='emf']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:公共类型='emf']">-->
			  <Default Extension="emf" ContentType="image/x-emf"/>
		  </xsl:if>
		  <!--09.12.22 马有旭 添加-->
		  <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='wav']">
			  <!--<xsl:if test=".//uof:其他对象[@uof:公共类型='wav']">-->
			  <Default Extension="wav" ContentType="audio/wav"/>
		  </xsl:if>
      
      <!--start liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->
      <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@私有类型_D707='eof']">
        <Default Extension="eof" ContentType="eof/eof"/>
      </xsl:if>
      <!--end liuyin 20130111 修改OLE对象经过转换后，ooxml文档无法打开-->

      <!--start liuyin 20130316 修改OLE对象经过转换后，ooxml文档无法打开-->
      <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@私有类型_D707='bin']">
        <Default Extension="bin" ContentType="bin/bin"/>
      </xsl:if>
      <!--end liuyin 20130316 修改OLE对象经过转换后，ooxml文档无法打开-->

      <!--start liuyin 20121203 修改音视频丢失-->
      <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='mp3']">
        <Default Extension="mp3" ContentType="audio/unknown"/>
      </xsl:if>
      <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='avi']">
        <Default Extension="avi" ContentType="video/avi"/>
      </xsl:if>
      <xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='wmv']">
        <Default Extension="wmv" ContentType="video/x-ms-wmv"/>
      </xsl:if>
      <!--<xsl:if test=".//对象:对象数据集_D700/对象:对象数据_D701[@公共类型_D706='avi']">
        <Default Extension="avi" ContentType="video/avi"/>
      </xsl:if>-->
      <!--end liuyin 20121203 修改音视频丢失-->
      
		  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
		  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>

		  <!--oox中有个默认的表格式样-->
		  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
		  <!--matadata-->
		  <!--李娟 修改标签 11.12.26 不知道路径是否对-->
		  <!--<xsl:if test="uof:元数据/uof:编辑时间|uof:元数据/uof:创建应用程序|uof:元数据/uof:文档模板|uof:元数据/uof:公司名称|uof:元数据/uof:经理名称|uof:元数据/uof:页数|uof:元数据/uof:字数|uof:元数据/uof:行数|uof:元数据/uof:段落数">-->
		  <xsl:if test="uof:元数据/元:元数据_5200/元:编辑时间_5209|uof:元数据/元:元数据_5200/元:创建应用程序_520A|uof:元数据/元:元数据_5200/元:文档模板_520C|uof:元数据/元:元数据_5200/元:公司名称_5213|uof:元数据/元:元数据_5200/元:经理名称_5214|uof:元数据/元:元数据_5200/元:页数_5215|uof:元数据/元:元数据_5200/元:字数_5216|uof:元数据/元:元数据_5200/元:字数_5216/元:行数_5219|uof:元数据/元:元数据_5200/元:字数_5216/元:段落数_521A">
			  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
		  </xsl:if>
		  <!--<xsl:if test="uof:元数据/uof:标题|uof:元数据/uof:主题|uof:元数据/uof:创建者|uof:元数据/uof:最后作者|uof:元数据/uof:摘要|uof:元数据/uof:创建日期|uof:元数据/uof:编辑次数|uof:元数据/uof:分类|uof:元数据/uof:关键字集">-->

		  <xsl:if test="uof:元数据/元:元数据_5200/元:标题_5201|uof:元数据/元:元数据_5200/元:主题_5202|uof:元数据/元:元数据_5200/元:创建者_5203|uof:元数据/元:元数据_5200/元:作者_5204|uof:元数据/元:元数据_5200/元:最后作者_5205">
			  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
		  </xsl:if>
		  <xsl:if test="uof:元数据/元:元数据_5200/元:用户自定义元数据集_520F">
			  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
		  </xsl:if>
		  <!--<xsl:for-each select=".//演:页面版式集/演:页面版式">-->
		  <xsl:for-each select=".//规则:页面版式集_B651/规则:页面版式_B652">
			 <Override PartName="/ppt/slideLayouts" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml">
				  <xsl:attribute name="PartName">
					  <xsl:value-of select="'/ppt/slideLayouts/'"/>
					  <xsl:value-of select="@标识符_6B0D"/>
					  <!--xsl:number count="演:页面版式" level="any" format="1"/-->
					  <xsl:value-of select="'.xml'"/>
				  </xsl:attribute>
			  </Override>
		  </xsl:for-each>
		  <xsl:for-each select=".//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide' or not(@类型_6BEA)]">
			  <!--<xsl:for-each select=".//演:母版集/演:母版[@演:类型='slide' or not(@演:类型)]">-->
			  <Override PartName="/ppt/slideMasters" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml">
				  <xsl:attribute name="PartName">
					  <xsl:value-of select="'/ppt/slideMasters/'"/>
					  <xsl:value-of select="@标识符_6BE8"/>
					  <xsl:value-of select="'.xml'"/>
				  </xsl:attribute>
			  </Override>
		  </xsl:for-each>

		  <xsl:if test=".//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
			  <xsl:for-each select=".//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
				  <Override PartName="/ppt/notesMasters" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml">
					  <xsl:attribute name="PartName">
						  <xsl:value-of select="'/ppt/notesMasters/'"/>
						  <!-- 09.10.12
                <xsl:value-of select="concat('notesMaster',substring-after(@演:标识符,'ID'))"/>              
                -->
						  <xsl:choose>
							  <xsl:when test="contains(@标识符_6BE8,'ID')">
								  <xsl:value-of select="concat('notesMaster',substring-after(@标识符_6BE8,'ID'))"/>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:value-of select="@标识符_6BE8"/>
							  </xsl:otherwise>
						  </xsl:choose>
						  <xsl:value-of select="'.xml'"/>
					  </xsl:attribute>
				  </Override>
			  </xsl:for-each>
		  </xsl:if>
		  <xsl:if test="not(.//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']) and .//演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
			  <!--<xsl:if test ="not(.//演:母版集/演:母版[@演:类型='notes']) and .//演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
			  <xsl:for-each select=".//演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
				  <Override PartName="/ppt/notesMasters" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml">
					  <xsl:attribute name="PartName">
						  <xsl:value-of select="'/ppt/notesMasters/'"/>
						  <!--<xsl:value-of select="concat('notesMaster',substring-after(uof:锚点/@uof:图形引用,'OBJ'))"/>-->
						  <xsl:value-of select="concat('notesMaster',substring-after(uof:锚点_C644/@图形引用_C62E,'Obj'))"/>
						  <xsl:value-of select="'.xml'"/>
					  </xsl:attribute>
				  </Override>
				  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml">
					  <xsl:attribute name="PartName">
						  <!--<xsl:value-of select ="concat('/ppt/notesSlides/','notesSlide',substring-after(uof:锚点/@uof:图形引用,'OBJ'),'.xml')"/>-->
						  <xsl:value-of select ="concat('/ppt/notesSlides/','notesSlide',substring-after(uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml')"/>
					  </xsl:attribute>
				  </Override>
			  </xsl:for-each>
		  </xsl:if>
		  <xsl:if test=".//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes'] and .//演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
			  <!--<xsl:if test=".//演:母版集/演:母版[@演:类型='notes'] and .//演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
			  <xsl:for-each select=" .//演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
				  <!--<xsl:for-each select=".//演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
				  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml">
					  <xsl:attribute name="PartName">
						  <xsl:value-of select ="concat('/ppt/notesSlides/','notesSlide',substring-after(uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml')"/>
					  </xsl:attribute>
				  </Override>
			  </xsl:for-each>
		  </xsl:if>
		  <!--/xsl:choose-->

		  <xsl:for-each select=".//演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='handout']">
			  <!--<xsl:for-each select=".//演:母版集/演:母版[@演:类型='handout']">-->
			  <Override PartName="/ppt/handoutMasters" ContentType="application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml">
				  <!--10.30 黎美秀修改-->
				  <xsl:attribute name="PartName">
					  <xsl:choose>
						  <xsl:when test="contains(@标识符_6BE8,'ID')">
							  <xsl:value-of select="concat('/ppt/handoutMasters/handoutMaster',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
						  </xsl:when>
						  <xsl:otherwise>
							  <xsl:value-of select="concat('/ppt/handoutMasters/',@标识符_6BE8,'.xml')"/>
						  </xsl:otherwise>
					  </xsl:choose>

				  </xsl:attribute>
			  </Override>
		  </xsl:for-each>
		  <xsl:for-each select=".//演:幻灯片集_6C0E/演:幻灯片_6C0F">
			  <!--<xsl:for-each select=".//演:幻灯片集/演:幻灯片">-->
			  <Override PartName="/ppt/slide" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml">
				  <xsl:attribute name="PartName">
					  <xsl:value-of select="'/ppt/slides/'"/>
					  <xsl:value-of select="@标识符_6B0A"/>
					  <xsl:value-of select="'.xml'"/>
				  </xsl:attribute>
			  </Override>
		  </xsl:for-each>
      
      <!--2014-04-19, tangjiang,  修复comment的PartName重复 start -->
		  <!--添加批逐集 李娟 2012.02.25-->
      <xsl:if test=".//规则:批注集_B669">
        <!--guoyongbin 2014-01-27 修改批注-->
        <!--<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml">
          <xsl:attribute name="PartName">
            <xsl:value-of select="'/ppt/comments/comment.xml'"/>
          </xsl:attribute>
        </Override>-->
        <xsl:for-each select="//演:幻灯片_6C0F">
          <xsl:if test="uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165">
            <Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml">
              <xsl:attribute name="PartName">
                <xsl:value-of select="concat('/ppt/comments/comment',substring-after(@标识符_6B0A,'slideId'),'.xml')"/>
              </xsl:attribute>
            </Override>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <!--2014-04-19, tangjiang,  修复comment的PartName重复 -->
      
		  <xsl:if test=".//规则:用户集_B667">
			  <Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>
		  </xsl:if>
	  </Types>
  </xsl:template>
</xsl:stylesheet>
