<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
        xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
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
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
        xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
        xmlns:演="http://schemas.uof.org/cn/2009/presentation"
        xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
        xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="PPr-commen.xsl"/>
  <!--xsl:import href="txStyles.xsl"/-->

  <xsl:import href="numbering.xsl"/>
	<xsl:import href="txStyles.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="styles">
	 
    <xsl:call-template name="fonts"/>
    <xsl:call-template name="numbering"/>
		<!--修改式样 李娟 2012.04.25-->
	  <!--<xsl:if test="p:presentation/p:notesMaster/p:notesStyle |p:presentation//a:lstStyle">
		 <xsl:for-each select="p:presentation/p:notesMaster/p:notesStyle/* |p:presentation//a:lstStyle/*">
		 <xsl:call-template name="noteStyle"/>
		 </xsl:for-each>
	  </xsl:if>-->
	  <xsl:for-each select="p:presentation/p:notesMaster">
		  <xsl:for-each select="p:notesStyle/*|.//a:lstStyle/*">
			  <xsl:call-template name="noteStyle"/>
		  </xsl:for-each>
	  </xsl:for-each>
	  <xsl:for-each select="p:presentation/p:notes//a:lstStyle/*">
		  <xsl:call-template name="noteStyle"/>
	  </xsl:for-each>
	</xsl:template>
  <xsl:template name="fonts">
 <式样:字体集_990C>
	 
      <xsl:for-each select="//node()[@typeface]">
        <xsl:if test="@typeface!='' and @typeface!='+mn-lt' and @typeface!='+mj-lt' and @typeface!='+mj-ea' and @typeface!='+mn-ea' and @typeface!='+mj-cs' and @typeface!='+mn-cs' ">

          <!--2013-11-5, tangjiang, 修改Strict OOXML到UOF字体集不全的转换 start -->
          <!--修改前
          <xsl:if test="not(current()/@typeface = preceding::node()[@typeface]/@typeface) and name(.)!='a:font'">
          -->
            <xsl:if test="not(current()/@typeface = preceding::node()[@typeface]/@typeface) and (name(.)='a:font' or name(.)='a:latin' )">
            <!--end-->
              
			  <式样:字体声明_990D>
			  <xsl:attribute name="标识符_9902">
				  <!--<xsl:variable name="id" select="position()"/>
				  <xsl:value-of select="concat('font_',$id)"/>-->
                <xsl:value-of select="@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="名称_9903">
				  <!--<xsl:value-of select="@val"/>-->
                <xsl:value-of select="@typeface"/>
              </xsl:attribute>
				  <!--2011-11-08 李娟--><!--
              --><!--<xsl:attribute name="uof:字体族">
                <xsl:value-of select="'auto'"/>
              </xsl:attribute>-->                             
              <xsl:attribute name="替换字体_9904">
                <xsl:value-of select="@typeface"/>
              </xsl:attribute>
				 <式样:字体族_9900>
				  <xsl:value-of select="'auto'"/>
			  </式样:字体族_9900>
			  </式样:字体声明_990D>
          </xsl:if>
        </xsl:if>
      </xsl:for-each>

   <!--2014-04-05, tangjiang, 添加Wingdings、Segoe UI字体，修改项目符号转换不正确 start -->
   <式样:字体声明_990D>
     <xsl:attribute name="标识符_9902">Wingdings</xsl:attribute>
     <xsl:attribute name="名称_9903">Wingdings</xsl:attribute>
     <xsl:attribute name="替换字体_9904">Wingdings</xsl:attribute>
     <式样:字体族_9900>auto</式样:字体族_9900>
   </式样:字体声明_990D>
   <式样:字体声明_990D>
     <xsl:attribute name="标识符_9902">Wingdings2</xsl:attribute>
     <xsl:attribute name="名称_9903">Wingdings 2</xsl:attribute>
     <xsl:attribute name="替换字体_9904">Wingdings 2</xsl:attribute>
     <式样:字体族_9900>auto</式样:字体族_9900>
   </式样:字体声明_990D>
   <式样:字体声明_990D>
     <xsl:attribute name="标识符_9902">Segoe UI</xsl:attribute>
     <xsl:attribute name="名称_9903">Segoe UI</xsl:attribute>
     <xsl:attribute name="替换字体_9904">Segoe UI</xsl:attribute>
     <式样:字体族_9900>auto</式样:字体族_9900>
   </式样:字体声明_990D>
   <!--2011-3-21罗文甜，上边的转换已存在“宋体”，故先去掉-->
   <!-- 
      <xsl:if test="not(//node[name(.)!=a:font and @typeface='宋体'])">
		  <式样:字体声明_990D>
			  <xsl:attribute name="标识符_9902">宋体</xsl:attribute>
			  <xsl:attribute name="名称_9903">宋体</xsl:attribute>
			  <xsl:attribute name="替换字体_9904">宋体</xsl:attribute>
			  <式样:字体族_9900>auto</式样:字体族_9900>
		  </式样:字体声明_990D>
				</xsl:if>
   -->
   <!--end 2014-04-05, tangjiang, 添加Wingdings、Segoe UI字体，修改项目符号转换不正确 -->
   
   <!--修改字体转换不正确 liqiuling 2013-3-25 -->
   <xsl:for-each select="//a:theme/a:themeElements/a:fontScheme//a:font[@script='Hans']">
     <xsl:if test="not(@typeface='宋体')">
       <式样:字体声明_990D>
         <xsl:attribute name="标识符_9902">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="名称_9903">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="名称_9903">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="替换字体_9904">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <式样:字体族_9900>
           <xsl:value-of select="'auto'"/>
         </式样:字体族_9900>
       </式样:字体声明_990D>
     </xsl:if>
   </xsl:for-each>

   <xsl:for-each select="//a:theme/a:themeElements/a:fontScheme//a:latin">
     <xsl:if test="not(@typeface='宋体')">
       <式样:字体声明_990D>
         <xsl:attribute name="标识符_9902">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="名称_9903">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="名称_9903">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <xsl:attribute name="替换字体_9904">
           <xsl:value-of select="@typeface"/>
         </xsl:attribute>
         <式样:字体族_9900>
           <xsl:value-of select="'auto'"/>
         </式样:字体族_9900>
       </式样:字体声明_990D>
     </xsl:if>
   </xsl:for-each>
   
   <!--修改字体转换不正确 liqiuling 2013-3-25 end-->
		 <!--<式样:替换字体族_9901>
			 <xsl:value-of select="@typeface"/>
		 </式样:替换字体族_9901>-->
    </式样:字体集_990C>
  </xsl:template>


  <xsl:template name="noteStyle">
	  <式样:段落式样集_9911>
		  <式样:段落式样_9912>
			  <xsl:attribute name="标识符_4100">
				  <xsl:value-of select="@id"/>
			  </xsl:attribute>
			  <xsl:attribute name="名称_4101">notes</xsl:attribute>
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
						  <xsl:otherwise>0</xsl:otherwise>
					  </xsl:choose>
				  </字:大纲级别_417C>
			  </xsl:if>
			  <!--PPr-commen 与 defRPr均在PPr-commen.xsl中定义-->
			  <xsl:call-template name="PPr-commen"/>
			  <xsl:apply-templates select="a:defRPr" mode="pPr-child"/>
		  </式样:段落式样_9912>
	  </式样:段落式样集_9911>
  </xsl:template>

	
	
	


</xsl:stylesheet>
