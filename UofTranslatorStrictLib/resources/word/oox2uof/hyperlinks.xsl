<?xml version="1.0" encoding="UTF-8"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:a15="http://schemas.microsoft.com/office/drawing/2012/main"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>

  <xsl:template name="hyperlink">
    <xsl:param name ="filename"/>
    <xsl:variable name="num" select="position()"/>
    <超链:超级链接_AA0C>
      <xsl:attribute name="标识符_AA0A">
        <xsl:value-of select="concat('hlkref_',$num)"/>
      </xsl:attribute>
      <!--<xsl:attribute name="名称_4166">
        <xsl:value-of select="'hyperlink'"/>
      </xsl:attribute>
      <xsl:attribute name="类型_413B">
        <xsl:value-of select="'hyperlink'"/>
      </xsl:attribute>-->

      <xsl:if test="@r:id">
        <xsl:variable name="rid">
          <xsl:value-of select="@r:id"/>
        </xsl:variable>
        <xsl:variable name ="goal">
          <!--超链接几种要转移到的目标-->
          <xsl:choose>           
            <xsl:when test ="$filename='document'">

              <!--2013-11-12，wudi，处理Strict标准下，@Target包含file:///的情况，start-->
              <xsl:variable name="tmp">
                <xsl:value-of select ="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:variable>
              <xsl:if test ="contains($tmp,':///')">
                <xsl:value-of select ="substring-after($tmp,':///')"/>
              </xsl:if>
              <xsl:if test ="not(contains($tmp,':///'))">
                <xsl:value-of select="$tmp"/>
              </xsl:if>
              <!--end-->
              
            </xsl:when>
            <xsl:when test ="$filename='comments'">
              <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
            </xsl:when>
            <xsl:when test ="$filename='endnotes'">
              <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
            </xsl:when>      
            <xsl:when test ="$filename='footnotes'">
              <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
            </xsl:when>
            <xsl:when test ="contains($filename,'header')">
              <xsl:variable name ="doc">
                <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
              </xsl:variable>        
              <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
            </xsl:when>
            <xsl:when test ="contains($filename,'footer')">
              <xsl:variable name ="doc">
                <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
              </xsl:variable>
              <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <超链:目标_AA01>
          <xsl:if test="contains($goal,'&quot;')">

            <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
            <xsl:if test ="contains($goal,':\\')">
              <xsl:call-template name ="pathOutput">
                <xsl:with-param name ="str1" select ="$goal"/>
                <xsl:with-param name ="str2" select ="''"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(contains($goal,':\\'))">
              <xsl:value-of select="translate($goal,'&quot;','')"/>
            </xsl:if>
            <!--end-->
            
          </xsl:if>
          <xsl:if test="not(contains($goal,'&quot;'))">

            <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
            <xsl:if test ="contains($goal,':\\')">
              <xsl:call-template name ="pathOutput">
                <xsl:with-param name ="str1" select ="$goal"/>
                <xsl:with-param name ="str2" select ="''"/>
              </xsl:call-template>
            </xsl:if>
            <xsl:if test ="not(contains($goal,':\\'))">
              <xsl:value-of select="translate($goal,'&quot;','')"/>
            </xsl:if>
            <!--end-->
            
          </xsl:if>
        </超链:目标_AA01>
      </xsl:if>
      <xsl:if test="$filename='field'"><!--2012.3.14-->
        <xsl:variable name="temp">
          <xsl:for-each select="parent::w:p//w:instrText">
            <xsl:value-of select="concat(.,'')"/>
          </xsl:for-each>
        </xsl:variable>
        
        <!--2013-03-27，wudi，修复超链接到书签(或标题)出错的BUG，start-->
        <xsl:variable name ="temp1">
          <xsl:value-of select ="substring-before($temp,'&quot;')"/>
        </xsl:variable>
        <xsl:if test ="contains($temp1,'\l')">
          <超链:书签_AA0D>
            <xsl:variable name="tmp">
              <xsl:value-of select="substring-after($temp,'&quot;')"/>
            </xsl:variable>
            <xsl:variable name ="fieldgoal">
              <xsl:value-of select ="translate($tmp,'&quot;','')"/>
            </xsl:variable>
            <xsl:value-of select ="normalize-space($fieldgoal)"/>
          </超链:书签_AA0D>
        </xsl:if>
        <xsl:if test ="not(contains($temp1,'\l'))">
          <超链:目标_AA01>
            <xsl:variable name="fieldgoal">
              <xsl:value-of select="substring-after($temp,'HYPERLINK')"/>
            </xsl:variable>
            <xsl:if test="contains($fieldgoal,' &quot;')">

              <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
              <xsl:if test ="contains($fieldgoal,':\\')">
                <xsl:call-template name ="pathOutput">
                  <xsl:with-param name ="str1" select ="substring-after(translate($fieldgoal,'&quot;',''),' ')"/>
                  <xsl:with-param name ="str2" select ="''"/>
                </xsl:call-template>
              </xsl:if>
              <xsl:if test ="not(contains($fieldgoal,':\\'))">
                <xsl:value-of select="substring-after(translate($fieldgoal,'&quot;',''),' ')"/>
              </xsl:if>
              <!--end-->
              
            </xsl:if>
            <xsl:if test="not(contains($fieldgoal,' &quot;'))">
              <xsl:if test="not(contains($fieldgoal,'&quot;'))">

                <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
                <xsl:if test ="contains($fieldgoal,':\\')">
                  <xsl:call-template name ="pathOutput">
                    <xsl:with-param name ="str1" select ="$fieldgoal"/>
                    <xsl:with-param name ="str2" select ="''"/>
                  </xsl:call-template>
                </xsl:if>
                <xsl:if test ="not(contains($fieldgoal,':\\'))">
                  <xsl:value-of select="$fieldgoal"/>
                </xsl:if>
                <!--end-->
                
              </xsl:if>
              <xsl:if test="contains($fieldgoal,'&quot;')">

                <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
                <xsl:if test ="contains($fieldgoal,':\\')">
                  <xsl:call-template name ="pathOutput">
                    <xsl:with-param name ="str1" select ="translate($fieldgoal,'&quot;','')"/>
                    <xsl:with-param name ="str2" select ="''"/>
                  </xsl:call-template>
                </xsl:if>
                <xsl:if test ="not(contains($fieldgoal,':\\'))">
                  <xsl:value-of select="translate($fieldgoal,'&quot;','')"/>
                </xsl:if>
                <!--end-->
                
              </xsl:if>
            </xsl:if>
          </超链:目标_AA01>
        </xsl:if>
        <!--end-->
        
      </xsl:if>
      <超链:链源_AA00>
        <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
      </超链:链源_AA00>
      <xsl:if test="@w:anchor">
        <超链:书签_AA0D>
          <xsl:value-of select="./@w:anchor"/>
        </超链:书签_AA0D>
      </xsl:if>
      <xsl:if test="@w:tooltip">
        <超链:提示_AA05>
          <xsl:value-of select="./@w:tooltip"/>
        </超链:提示_AA05>
      </xsl:if>
    </超链:超级链接_AA0C>
  </xsl:template>

  <!--2013-04-25，wudi，修复OOX到UOF方向超链接转换的BUG，文件路径问题，start-->
  <xsl:template name ="pathOutput">
    <xsl:param name ="str1"/>
    <xsl:param name ="str2"/>
    <xsl:if test ="contains($str1,'\\')">
      <xsl:variable name ="tmp1">
        <xsl:value-of select ="substring-before($str1,'\\')"/>
      </xsl:variable>
      <xsl:variable name ="tmp2">
        <xsl:value-of select ="substring-after($str1,'\\')"/>
      </xsl:variable>
      <xsl:variable name ="tmp3">
        <xsl:value-of select ="concat($str2,$tmp1,'\')"/>
      </xsl:variable>
      <xsl:call-template name ="pathOutput">
        <xsl:with-param name ="str1" select ="$tmp2"/>
        <xsl:with-param name ="str2" select ="$tmp3"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="not(contains($str1,'\\'))">
      <xsl:value-of select ="concat($str2,$str1)"/>
    </xsl:if>
  </xsl:template>
  <!--end-->
  
</xsl:stylesheet>
