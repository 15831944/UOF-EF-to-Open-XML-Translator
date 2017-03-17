<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects" 
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
  <xsl:import href="style.xsl"/>
  <xsl:import href="table.xsl"/>

  <xsl:template name="object_root">
    <!--sheetName:Sheet1-->
    <xsl:variable name="sheetName">
      <xsl:value-of select="./ws:spreadsheets/ws:spreadsheet/@sheetName"/>
    </xsl:variable>
    <xsl:for-each select=" .//ws:spreadsheets/ws:spreadsheet/pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']">
      <xsl:variable name="rID">
        <xsl:value-of select="./@Id"/>
      </xsl:variable>
      <xsl:variable name="target">
        <xsl:value-of select="./@Target"/>
      </xsl:variable>
      <xsl:variable name="imageName">
        <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
      </xsl:variable>

      <对象:对象数据_D701>
        <!--标识符:image1-->
        <xsl:attribute name="标识符_D704">
          <xsl:value-of select="$imageName"/>
        </xsl:attribute>
        <xsl:attribute name="是否内嵌_D705">false</xsl:attribute>
        <xsl:choose>
          <xsl:when test="contains($target,'emf')">
            <xsl:attribute name="私有类型_D707">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="公共类型_D706">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <!--<对象:数据_D702>
          --><!--用C# 转media中的图片，转成二进制代码--><!--
        </对象:数据_D702>-->
        
        <!--添加元素：路径，李杨2011-11-07-->
        <xsl:element name ="对象:路径_D703">
          <xsl:variable name="ptype">
            <xsl:value-of select="substring-after($target,'../media/')"/>
          </xsl:variable>
          <xsl:variable name ="pname">
            <xsl:value-of select ="concat($imageName,'.',substring-after($ptype,'.'))"/>
          </xsl:variable>
          <xsl:value-of select ="concat('.\data\',$pname)"/>
        </xsl:element>
      </对象:对象数据_D701>
    </xsl:for-each>
    <xsl:for-each select="./ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']">
      <xsl:variable name="rID">
        <xsl:value-of select="./@Id"/>
      </xsl:variable>
      <xsl:variable name="target">
        <xsl:value-of select="./@Target"/>
      </xsl:variable>
      <xsl:variable name="imageName">
        <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
      </xsl:variable>

      <对象:对象数据_D701>
        <!--标识符:image1-->
        <xsl:attribute name="标识符_D704">
          <xsl:value-of select="$imageName"/>
        </xsl:attribute>
        <xsl:attribute name="是否内嵌_D705">false</xsl:attribute>
        <xsl:choose>
          <xsl:when test="contains($target,'emf')">
            <xsl:attribute name="私有类型_D707">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="公共类型_D706">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <!--<对象:数据_D702>
          --><!--用C# 转media中的图片，转成二进制代码--><!--
        </对象:数据_D702>-->

        <!--添加元素：路径，李杨2011-12-07-->
        <xsl:element name ="对象:路径_D703">
          <xsl:variable name="ptype">
            <xsl:value-of select="substring-after($target,'../media/')"/>
          </xsl:variable>
          <xsl:variable name ="pname">
            <xsl:value-of select ="concat($imageName,'.',substring-after($ptype,'.'))"/>
          </xsl:variable>
          <xsl:value-of select ="concat('.\data\',$pname)"/>
        </xsl:element>
      </对象:对象数据_D701>
    </xsl:for-each>
    <!--添加图表中的数据系列 图片填充时，图片对象引用。李杨2012-3-20-->
    <xsl:for-each select ="./ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace/pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']">     
      <xsl:variable name="target">
        <xsl:value-of select="./@Target"/>
      </xsl:variable>
      <xsl:variable name="imageName">
        <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
      </xsl:variable>
      <对象:对象数据_D701>
        <xsl:attribute name="标识符_D704">
          <xsl:value-of select="$imageName"/>
        </xsl:attribute>
        <xsl:attribute name="是否内嵌_D705">false</xsl:attribute>
        <xsl:choose>
          <xsl:when test="contains($target,'emf')">
            <xsl:attribute name="私有类型_D707">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="公共类型_D706">
              <xsl:variable name="type">
                <xsl:value-of select="substring-after($target,'../media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($type,'.')"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>

        <xsl:element name ="对象:路径_D703">
          <xsl:variable name="ptype">
            <xsl:value-of select="substring-after($target,'../media/')"/>
          </xsl:variable>
          <xsl:variable name ="pname">
            <xsl:value-of select ="concat($imageName,'.',substring-after($ptype,'.'))"/>
          </xsl:variable>
          <xsl:value-of select ="concat('.\data\',$pname)"/>
        </xsl:element>
      </对象:对象数据_D701>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name ="drawings">
    <xsl:for-each select="./ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:oneCellAnchor">
      <!-- add by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致  -->
      <xsl:variable name="posit" select="position()"/>
      <!-- end -->
      <xsl:for-each select="xdr:cxnSp|xdr:grpSp|xdr:pic|xdr:sp">
        <xsl:apply-templates select=".">
          <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
          <xsl:with-param name="posi" select="$posit"/>
        </xsl:apply-templates>
      </xsl:for-each>
    </xsl:for-each>
    <xsl:for-each select="./ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:twoCellAnchor">
      <xsl:if test=".//xdr:grpSp">
        <xsl:apply-templates select=".//xdr:grpSp">
          <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:if test="./xdr:pic">
        <xsl:apply-templates select="./xdr:pic">
          <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:if test="./xdr:sp">
        <xsl:apply-templates select="./xdr:sp">
          <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:apply-templates>
      </xsl:if>

      <xsl:if test="./xdr:cxnSp">
        <xsl:apply-templates select="./xdr:cxnSp">
          <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:apply-templates>
      </xsl:if>

      <xsl:if test ="./xdr:graphicFrame">
        <xsl:choose>
          <xsl:when test ="contains(./xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name,'Diagram')">
            <!-- 20130506 add by xuzhenwei BUG_2909:第二轮回归oo-uof smartArt丢失 start  -->
            <xsl:if test="ancestor::xdr:wsDr/dsp:drawing/dsp:spTree/dsp:sp">
              <xsl:variable name="rcs" select="./xdr:graphicFrame/a:graphic/a:graphicData/dgm:relIds/@r:cs"/>
              <xsl:variable name="rels" select="substring($rcs,1,3)"/>
              <xsl:variable name="number" select="substring-after($rcs,'rId')"/>
              <xsl:variable name="dsmRel" select="concat($rels,$number + 1)"/>
              <xsl:call-template name="dspdrawing">
                <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
                <xsl:with-param name="dsmRelid" select="$dsmRel"/>
              </xsl:call-template>
            </xsl:if>
            <!-- end -->
          </xsl:when>
          <!-- update xuzhenwei 20121228 将smartArt转换成组合图形 -->
          <xsl:when test ="contains(./xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name,'图示')">
            <xsl:variable name="rcs" select="./xdr:graphicFrame/a:graphic/a:graphicData/dgm:relIds/@r:cs"/>
            <xsl:variable name="rels" select="substring($rcs,1,3)"/>
            <xsl:variable name="number" select="substring-after($rcs,'rId')"/>
            <xsl:variable name="dsmRel" select="concat($rels,$number + 1)"/>
            <xsl:call-template name="dspdrawing">
              <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
              <xsl:with-param name="dsmRelid" select="$dsmRel"/>
            </xsl:call-template>
          </xsl:when>
          <!-- end -->
          <xsl:otherwise>
            <xsl:variable name="grapFromWidth" select="./xdr:from/xdr:colOff"/>
            <xsl:variable name="grapToWidth" select="./xdr:to/xdr:colOff"/>
            <xsl:apply-templates select ="./xdr:graphicFrame">
              <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
              <xsl:with-param name="grapWidth" select="$grapToWidth - $grapFromWidth"/>
            </xsl:apply-templates>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:for-each>
    <!--批注集
    <xsl:for-each select="./ws:spreadsheets/ws:spreadsheet/ws:comments">
      <xsl:call-template name="pizhuji-root">
        <xsl:with-param name="comment-name" select="."/>
      </xsl:call-template>
    </xsl:for-each>-->
    <xsl:for-each select ="./ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:absoluteAnchor/xdr:graphicFrame/xdr:nvGraphicFramePr">
      <xsl:apply-templates select ="./xdr:cNvPr">
        <xsl:with-param name="s_name" select="ancestor::ws:spreadsheet/@sheetName"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match="xdr:cNvPr">
    <!--chartSheet-->
    <xsl:param name="s_name"/>
    <图:图形_8062>
      <xsl:variable name="ppid">
          <xsl:value-of select="./@id"/>
      </xsl:variable>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$ppid)"/>
      </xsl:attribute>
      <图:图表数据引用_8065>
        <xsl:variable name ="sheetNo">
          <xsl:value-of select ="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:variable>
        <xsl:variable name ="id">
          <xsl:value-of select ="ancestor::xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id"/>
        </xsl:variable>
        <xsl:value-of select ="concat($sheetNo,'_chart_',$id)"/>
      </图:图表数据引用_8065>
    </图:图形_8062>
  </xsl:template>
  <xsl:template name="PicProperties">
    <xsl:param name="picId"/>
    <图:预定义图形_8018>
      <xsl:apply-templates select="xdr:pic/xdr:spPr" mode="first"/>
      <图:属性_801D>
        <xsl:apply-templates select="xdr:pic/xdr:spPr" mode="second"/>
      </图:属性_801D>
      <xsl:if test="xdr:grpSp">
        <xsl:if test="xdr:grpSp/xdr:pic">
          <xsl:if test="xdr:grpSp/xdr:pic/xdr:blipFill/a:blip[@r:embed=$picId]">
            <xsl:apply-templates select="xdr:grpSp/xdr:pic/xdr:spPr" mode="first"/>
            <图:属性_801D>
              <xsl:apply-templates select="xdr:grpSp/xdr:pic/xdr:spPr" mode="second"/>
            </图:属性_801D>
          </xsl:if>
        </xsl:if>
      </xsl:if>
    </图:预定义图形_8018>
  </xsl:template>
  <xsl:template match="xdr:pic">
    <xsl:param name="s_name"/>
    <图:图形_8062>
      <!--层次在生成的xml文件中没有设置，是twoCellAnchor的位置-->
      <xsl:variable name="ppid">
        <xsl:if test="xdr:nvPicPr">
          <xsl:value-of select="xdr:nvPicPr/xdr:cNvPr/@id"/>
        </xsl:if>
        <!--yx,delete 10.22 16:08 xsl:if test="xdr:nvCxnSpPr">
          <xsl:value-of select="xdr:nvCxnSpPr/xdr:cNvPr/@id"/>
        </xsl:if-->
      </xsl:variable>
      <!--3yue27ri 夏艳霞checked-->
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$ppid)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <xsl:number count="xdr:twoCellAnchor" level="single"/>
      </xsl:attribute>
      <!--<xsl:attribute name="图:其他对象">
        <xsl:value-of select="substring-before(substring-after($xtuxm1,'media/'),'.')"/>
      </xsl:attribute>-->
      <!--标识符:Sheet1_OBJ1-->
      <图:预定义图形_8018>
        <xsl:apply-templates select="xdr:spPr" mode="first"/>
        <图:属性_801D>
          <xsl:apply-templates select="xdr:spPr" mode="second"/>
          <!-- 20130308 add by xuzhenwei BUG_2705: oo-uof 测试用例PageSize.xlsx 工作表名 Letter：图片丢失(图片剪裁) start -->
          <xsl:if test="xdr:blipFill/a:srcRect">
            <xsl:apply-templates select="xdr:blipFill/a:srcRect"/>
          </xsl:if>
          <!-- end -->
        </图:属性_801D>
      </图:预定义图形_8018>

      <!--添加：图片数据引用  李杨2011-12-07-->
      <xsl:element name ="图:图片数据引用_8037">
        <!-- 20130301 update by xuzhenwei BUG_2705:图片和图形组合，内容丢失 start -->
        <xsl:choose>
          <xsl:when test="xdr:blipFill/a:blip/@r:embed">
            <xsl:variable name="xtuxm" select="xdr:blipFill/a:blip/@r:embed"/>
            <xsl:variable name="xtuxm1" select="ancestor::xdr:twoCellAnchor/following::pr:Relationships/pr:Relationship[@Id=$xtuxm]/@Target"/>
            <xsl:value-of select="substring-before(substring-after($xtuxm1,'../media/'),'.')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="rId">
              <xsl:value-of select=".//a:blip/@r:embed"/>
            </xsl:variable>
            <xsl:variable name="target">
              <xsl:value-of select="../../pr:Relationships/pr:Relationship[@Id = $rId]/@Target"/>
            </xsl:variable>
            <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!-- end -->
      </xsl:element>
      <xsl:if test="./xdr:spPr/a:xfrm/@flipH = '1'">
        <图:翻转_803A>x</图:翻转_803A>
      </xsl:if>
      <xsl:if test="xdr:pic/xdr:spPr/a:xfrm[@rot]">
        <图:翻转_803A>
          <!--uof:attrList="方向" 图:方向="x">-->
          <xsl:variable name="rot" select="xdr:pic/xdr:spPr/a:xfrm/@rot"/>
          <xsl:value-of select="$rot div 60000"/>
        </图:翻转_803A>
      </xsl:if>
      <!--组合位置-->
      <xsl:if test="parent::xdr:grpSp">
        <图:组合位置_803B>
          <xsl:call-template name="groupLocation"/>
        </图:组合位置_803B>
      </xsl:if>
      <图:文本_803C>
        <!--uof:attrList="文本框 左边距 右边距 上边距 下边距 水平对齐 垂直对齐 文字排列方向 自动换行 大小适应文字 前一链接 后一链接">-->
        <xsl:apply-templates select="xdr:txBody/a:bodyPr"/>
        <图:内容_8043>
          <xsl:apply-templates select="xdr:txBody/a:p"/>
        </图:内容_8043>
      </图:文本_803C>
    </图:图形_8062>
  </xsl:template>
  <xsl:template match="xdr:pic" mode="levell">
    <xsl:value-of select="position()"/>
  </xsl:template>

  <!-- 20130308 add by xuzhenwei BUG_2705: oo-uof 测试用例PageSize.xlsx 工作表名 Letter：图片丢失(图片剪裁) start -->
  <xsl:template match="a:srcRect">
    <图:图片属性_801E>
      <图:颜色模式_801F>auto</图:颜色模式_801F>
      <图:亮度_8020>50.0</图:亮度_8020>
      <图:对比度_8021>50.0</图:对比度_8021>
      <图:图片裁剪_8022>
        <xsl:choose>
          <xsl:when test="./@t">
            <图:上_8023>
              <xsl:value-of select="./@t div 1000"/>
            </图:上_8023>
          </xsl:when>
          <xsl:otherwise>
            <图:上_8023>0.0</图:上_8023>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="./@b">
            <图:下_8024>
              <xsl:value-of select="./@b div 1000"/>
            </图:下_8024>
        </xsl:when>
          <xsl:otherwise>
            <图:下_8024>0.0</图:下_8024>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="./@l">
            <图:左_8025>
              <xsl:value-of select="./@l div 1000"/>
            </图:左_8025>
          </xsl:when>
          <xsl:otherwise>
            <图:左_8025>0.0</图:左_8025>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="./@t">
            <图:右_8026>
              <xsl:value-of select="./@r div 1000"/>
            </图:右_8026>
        </xsl:when>
          <xsl:otherwise>
            <图:右_8026>0.0</图:右_8026>
        </xsl:otherwise>
        </xsl:choose>
      </图:图片裁剪_8022>
    </图:图片属性_801E>
  </xsl:template>
  <!-- end -->
  <!--批注集-->
  <xsl:template name="pizhuji-root">
    <xsl:param name="comment-name"/>
    <xsl:variable name="a" select="concat('comments',./ancestor::ws:spreadsheet/@sheetId,'.')"/>
    <xsl:for-each select=".//ws:comment">
      <xsl:variable name="p">
        <xsl:number count="ws:comment" level="single"/>
      </xsl:variable>
      <图:图形_8062>
        <xsl:attribute name="标识符_804B">
          <xsl:value-of select="concat($a,$p)"/>
          <!--标识符_804B:comments1.1-->
        </xsl:attribute>
        <xsl:attribute name ="层次_8063">0</xsl:attribute>

        <图:预定义图形_8018>
          <xsl:call-template name="ydy-tuxing">
            <!--预定义图形-->
            <xsl:with-param name="cname" select="$a"/>
            <xsl:with-param name="pos">
              <xsl:number count="ws:comment" level="single"/>
            </xsl:with-param>
          </xsl:call-template>
        </图:预定义图形_8018>
        <!--文本内容-->
        <图:文本_803C 是否为文本框_8046="false" 是否自动换行_8047="true" 是否大小适应文字_8048="false"
            是否文字随边框自动缩放_804A="true" 是否文字绕排外部对象_8068="true">
          <图:边距_803D 左_C608="7.2" 上_C609="3.6" 右_C60A="7.2" 下_C60B="3.6"/>
          <图:对齐_803E 水平对齐_421D="left" 文字对齐_421E="top"/>
          <图:文字排列方向_8042>t2b-l2r-0e-0w</图:文字排列方向_8042>
          <xsl:call-template name="wbnr">
            <xsl:with-param name="cname" select="$a"/>
            <xsl:with-param name="pos">
              <xsl:number count="ws:comment" level="single"/>
            </xsl:with-param>
          </xsl:call-template>
        </图:文本_803C>
        <!--/xsl:if-->
      </图:图形_8062>
    </xsl:for-each>
  </xsl:template>
  
  <!--预定义图形-->
  <xsl:template name="ydy-tuxing">
    <xsl:param name="cname"/>
    <xsl:param name="pos"/>
    
      <图:类别_8019>11</图:类别_8019>
      <图:名称_801A>Rectangle</图:名称_801A>
      <图:生成软件_801B>Microsoft</图:生成软件_801B>
      <图:属性_801D>
        <xsl:variable name="vmlname" select="translate(concat('vmlDrawing',substring-after($cname,'s')),'x','v')"/>
        <!--vmlname:vmlDrawing1.vml-->
        <xsl:variable name="vml" select="concat('xlsx/xl/drawings/',$vmlname)"/>
        <!--vml:xlsx/xl/drawings/vmlDrawing1.vml-->
        <图:填充_804C 是否填充随图形旋转_8067="false">
          <图:颜色_8004>
            <xsl:value-of select="'#ffffe1'"/>
          </图:颜色_8004>
        </图:填充_804C>
        <图:旋转角度_804D>0.0</图:旋转角度_804D>
        <图:是否打印对象_804E>true</图:是否打印对象_804E>
        <图:阴影_8051 是否显示阴影_C61C="true" 类型_C61D="single" 颜色_C61E="#000000" 透明度_C61F="100">
          <uof:偏移量_C61B x_C606="2.0" y_C607="2.0"/>
        </图:阴影_8051>
        <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
        <图:线_8057>
          <图:线颜色_8058>#000000</图:线颜色_8058>
          <图:线类型_8059 线型_805A="single" 虚实_805B="solid"/>
          <图:线粗细_805C>0.75</图:线粗细_805C>
        </图:线_8057>
        <图:大小_8060 长_C604="59.999084" 宽_C605="97.49851"/>
        
        <!--<图:Web文字_804F>Graph</图:Web文字_804F>-->
      </图:属性_801D>
    
  </xsl:template>
  <xsl:template name="wbnr">
    <!--文本内容 修改  李杨2011-11-29-->
    <xsl:param name="cname"/>
    <xsl:param name="pos"/>
    <!--cname:comments1.xml-->
    <!--pos:第几个批注-->
    <xsl:variable name="commentsname" select="concat('xlsx/xl/',$cname)"/>
    <!--commentsname:xlsx/xl/comments1.xml-->
    <!--<xsl:for-each select="document($commentsname)/ws:comments/ws:commentList/ws:comment[position()=$pos]/ws:text">-->
    <xsl:for-each select="./ancestor-or-self::ws:spreadsheet/ws:comments/ws:commentList/ws:comment[position()=$pos]/ws:text">
      <xsl:variable name="textid">
        <xsl:number count="ws:text" level="single"/>
      </xsl:variable>
      <图:内容_8043>
        <字:段落_416B>
          <字:段落属性_419B>
            <字:对齐_417D 字:水平对齐="left"/>
            <!--<字:边框_4133 阴影类型_C645="">
              <uof:下_C616 线型_C60D="none"/>
            </字:边框_4133>-->
          </字:段落属性_419B>
          <xsl:for-each select="ws:r">
            <xsl:variable name="rid">
              <xsl:number count="ws:r" level="single"/>
            </xsl:variable>
            <xsl:variable name="fontid" select="concat(substring-before($cname,'.'),$pos,$textid,$rid)"/>
            <!--fontid:comments1.1.1.1-->
            <字:句_419D>
              <字:句属性_4158>
                <xsl:if test="ws:rPr/ws:rFont">
                  <字:字体_4128>
                    <!--<字:字体 uof:locID="t0088" uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色">-->
                    <!--字:字体:font_comments1.1.1.1-->
                    <xsl:if test="ws:rPr[ws:rFont]">
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat('font_',$fontid)"/>
                      </xsl:attribute>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat('font_',$fontid)"/>
                      </xsl:attribute>
                    </xsl:if>
                    <!--字号 -->
                    <xsl:if test="ws:rPr[ws:sz]">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="ws:rPr/ws:sz/@val"/>
                      </xsl:attribute>
                    </xsl:if>
                    <!--color-->
                    <xsl:if test="ws:rPr[ws:color]">
                      <xsl:if test="ws:rPr/ws:color[@theme]">
                        <xsl:variable name="the" select="ws:rPr/ws:color/@theme"/>
                        <!--<xsl:variable name="them" select="document('xlsx/xl/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:sysClr/@val"/>
                      <xsl:variable name="them2" select="document('xlsx/xl/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:srgbClr/@val"/>-->
                        <xsl:variable name="them" select="/ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:sysClr/@val"/>
                        <xsl:variable name="them2" select="/ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:srgbClr/@val"/>
                        <xsl:attribute name="颜色_412F">
                          <xsl:choose>
                            <xsl:when test="$them='windowText'">
                              <xsl:value-of select="'#000000'"/>
                            </xsl:when>
                            <xsl:when test="$them='window'">
                              <xsl:value-of select="'#000000'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:variable name="colll">
                                <xsl:value-of select="translate($them2,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                              </xsl:variable>
                              <xsl:value-of select="concat('#',$colll)"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                      </xsl:if>
                      <xsl:if test="ws:rPr/ws:color[@rgb]">
                        <xsl:attribute name="颜色_412F">
                          <xsl:variable name="c">
                            <xsl:value-of select="ws:rPr/ws:color/@rgb"/>
                          </xsl:variable>
                          <xsl:variable name="co">
                            <xsl:value-of select="substring($c,3,8)"/>
                          </xsl:variable>
                          <xsl:variable name="col">
                            <xsl:value-of select="translate($co,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                          </xsl:variable>
                          <xsl:value-of select="concat('#',$col)"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:if>
                  </字:字体_4128>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:b]">
                  <字:是否粗体_4130>true</字:是否粗体_4130>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:i]">
                  <字:是否斜体_4131>true</字:是否斜体_4131>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:strike]">
                  <字:删除线_4135>single</字:删除线_4135>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:u]">
                  <xsl:if test="ws:rPr/ws:u[@val]">
                    <字:下划线_4136>
                      <xsl:attribute name="线型_4137">
                        <xsl:if test="ws:rPr/ws:u[@val='double' or @val='doubleAccounting' ]">
                          <xsl:value-of select="'double'"/>
                        </xsl:if>
                        <xsl:if test="ws:rPr/ws:u[@val='single' or @val='singleAccounting' ]">
                          <xsl:value-of select="'single'"/>
                        </xsl:if>
                      </xsl:attribute>
                    </字:下划线_4136>
                  </xsl:if>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:outline]">
                  <字:是否空心_413E>true</字:是否空心_413E>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:shadow]">
                  <字:是否阴影_4140>true</字:是否阴影_4140>
                </xsl:if>
                <xsl:if test="ws:rPr[ws:vertAlign]">
                  <字:上下标类型_4143>
                    <xsl:if test="ws:rPr[ws:vertAlign]">
                      <xsl:if test="ws:rPr/ws:vertAlign[@val='superscript']">
                        <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="'sup'"/>
                        <!--</xsl:attribute>-->
                      </xsl:if>
                    </xsl:if>
                    <xsl:if test="ws:rPr[ws:vertAlign]">
                      <xsl:if test="ws:rPr/ws:vertAlign[@val='subscript']">
                        <!--<xsl:attribute name="字:值">-->
                          <xsl:value-of select="'sub'"/>
                        <!--</xsl:attribute>-->
                      </xsl:if>
                    </xsl:if>
                  </字:上下标类型_4143>
                </xsl:if>
              </字:句属性_4158>
              <字:文本串_415B>
                <xsl:value-of select="ws:t"/>
              </字:文本串_415B>
            </字:句_419D>
          </xsl:for-each>
        </字:段落_416B>
      </图:内容_8043>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="tupian-ydy-tuxing">
    <xsl:param name="picId"/>
    <图:预定义图形_8018>
      <xsl:apply-templates select="xdr:pic/xdr:spPr" mode="first"/>
      <图:属性_801D>
        <xsl:apply-templates select="xdr:pic/xdr:spPr" mode="second"/>
      </图:属性_801D>
      <!--/xsl:if-->
      <!--/xsl:if>
        </xsl:if-->
      <xsl:if test="xdr:grpSp">
        <xsl:if test="xdr:grpSp/xdr:pic">
          <xsl:if test="xdr:grpSp/xdr:pic/xdr:blipFill/a:blip[@r:embed=$picId]">
            <xsl:apply-templates select="xdr:grpSp/xdr:pic/xdr:spPr" mode="first"/>
            <图:属性_801D>
              <xsl:apply-templates select="xdr:grpSp/xdr:pic/xdr:spPr" mode="second"/>
            </图:属性_801D>
          </xsl:if>
        </xsl:if>
      </xsl:if>
    </图:预定义图形_8018>
    <!--/xsl:for-each-->
  </xsl:template>
  <xsl:template match="xdr:spPr|dsp:spPr" mode="first">
    <图:类别_8019>
      <xsl:call-template name="ShapeClass"/>
    </图:类别_8019>
    <图:名称_801A>
      <xsl:call-template name="ShapeName"/>
    </图:名称_801A>
    <图:生成软件_801B>Microsoft Office</图:生成软件_801B>
    <!--路径-->
    <xsl:if test="a:custGeom/a:pathLst/a:path">
      <xsl:call-template name="keyCorCommen"/>
    </xsl:if>
  </xsl:template>
  <xsl:template name="keyCorCommen">
    <!--<图:关键点坐标 uof:locID="g0009" uof:attrList="路径">
      <xsl:attribute name="图:路径">-->
    <图:路径_801C>
      <图:路径值_8069>
        <xsl:if test="a:custGeom/a:pathLst/a:path/a:moveTo">
          <xsl:variable name="xStart">
            <xsl:choose>
              <xsl:when test="a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@x!='0'">
                <xsl:value-of select="a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@x div 12700 * 1.34"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'0'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="yStart">
            <xsl:choose>
              <xsl:when test="a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@y!='0'">
                <xsl:value-of select="a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@y div 12700 * 1.34"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'0'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="Start" select="concat('M',' ',$xStart,' ',$yStart)"/>
          <xsl:value-of select="concat($Start,' ')"/>
        </xsl:if>
        <xsl:for-each select="a:custGeom/a:pathLst/a:path/a:cubicBezTo|a:custGeom/a:pathLst/a:path/a:lnTo">
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
    </图:路径_801C>
      <!--</xsl:attribute>
    </图:关键点坐标>-->
  </xsl:template>
  <xsl:template name="ShapeClass2">
    <xsl:if test="a:prstGeom/@prst">
      <xsl:choose>
        <xsl:when test="a:prstGeom/@prst='heptagon'">16</xsl:when>
        <xsl:when test="a:prstGeom/@prst='rect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='parallelogram'">12</xsl:when>
        <xsl:when test="a:prstGeom/@prst='trapezoid'">13</xsl:when>
        <xsl:when test="a:prstGeom/@prst='diamond'">14</xsl:when>
        <xsl:when test="a:prstGeom/@prst='roundRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='snip1Rect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='snip2SameRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='snipRoundRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='snip2DiagRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='round1Rect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='round2SameRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='round2DiagRect'">11</xsl:when>
        <xsl:when test="a:prstGeom/@prst='octagon'">16</xsl:when>
        <xsl:when test="a:prstGeom/@prst='triangle'">17</xsl:when>
        <xsl:when test="a:prstGeom/@prst='rtTriangle'">18</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipse'">19</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='flowChartInternalStorage'">36</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDocument'">37</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMultidocument'">38</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartTerminator'">39</xsl:when>
        <xsl:when test="a:prstGeom/@prst='irregularSeal1'">41</xsl:when>
        <xsl:when test="a:prstGeom/@prst='irregularSeal2'">42</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star4'">43</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star5'">44</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star8'">45</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star12'">46</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star24'">47</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star32'">48</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ribbon2'">49</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeRectCallout'">51</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeRoundRectCallout'">52</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeEllipseCallout'">53</xsl:when>
        <xsl:when test="a:prstGeom/@prst='cloudCallout'">54</xsl:when>
        <xsl:when test="a:prstGeom/@prst='line'">61</xsl:when>
        <xsl:when test="a:prstGeom/@prst='straightConnector1'">61</xsl:when>
        <xsl:when test="a:prstGeom/@prst='straightConnector1'">63</xsl:when>
        <xsl:when test="a:prstGeom/@prst='curvedConnector3'">64</xsl:when>
        <xsl:when test="a:prstGeom/@prst='straightConnector1'">71</xsl:when>
        <xsl:when test="a:prstGeom/@prst='bentConnector3'">74</xsl:when>
        <xsl:when test="a:prstGeom/@prst='curvedConnector3'">77</xsl:when>
        <xsl:when test="a:prstGeom/@prst='hexagon'">110</xsl:when>
        <xsl:when test="a:prstGeom/@prst='plus'">111</xsl:when>
        <xsl:when test="a:prstGeom/@prst='pentagon'">112</xsl:when>
        <xsl:when test="a:prstGeom/@prst='can'">113</xsl:when>
        <xsl:when test="a:prstGeom/@prst='cube'">114</xsl:when>
        <xsl:when test="a:prstGeom/@prst='bevel'">115</xsl:when>
        <xsl:when test="a:prstGeom/@prst='foldedCorner'">116</xsl:when>
        <xsl:when test="a:prstGeom/@prst='smileyFace'">117</xsl:when>
        <xsl:when test="a:prstGeom/@prst='donut'">118</xsl:when>
        <xsl:when test="a:prstGeom/@prst='noSmoking'">119</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='homePlate'">219</xsl:when>
        <xsl:when test="a:prstGeom/@prst='chevron'">220</xsl:when>
        <xsl:when test="a:prstGeom/@prst='rightArrowCallout'">221</xsl:when>
        <xsl:when test="a:prstGeom/@prst='leftArrowCallout'">222</xsl:when>
        <xsl:when test="a:prstGeom/@prst='upArrowCallout'">223</xsl:when>
        <xsl:when test="a:prstGeom/@prst='downArrowCallout'">224</xsl:when>
        <xsl:when test="a:prstGeom/@prst='leftRightArrowCallout'">225</xsl:when>
        <xsl:when test="a:prstGeom/@prst='upDownArrowCallout'">226</xsl:when>
        <xsl:when test="a:prstGeom/@prst='quadArrowCallout'">227</xsl:when>
        <xsl:when test="a:prstGeom/@prst='circularArrow'">228</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='flowChartOnlineStorage'">323</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDelay'">324</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticTape'">325</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticDisk'">326</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticDrum'">327</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDisplay'">328</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ribbon'">410</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipseRibbon'">412</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipseRibbon2'">411</xsl:when>
        <xsl:when test="a:prstGeom/@prst='verticalScroll'">413</xsl:when>
        <xsl:when test="a:prstGeom/@prst='horizontalScroll'">414</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wave'">415</xsl:when>
        <xsl:when test="a:prstGeom/@prst='doubleWave'">416</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout1'">513</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout2'">56</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout3'">58</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout1'">59</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout2'">511</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout3'">512</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout1'">55</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout2'">515</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout3'">516</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout1'">517</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout2'">519</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout3'">520</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star10' or a:prstGeom/@prst='star12'">46</xsl:when>
        <xsl:otherwise>11</xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="not(a:prstGeom)">
      <xsl:value-of select="'64'"/>
    </xsl:if>
  </xsl:template>

  <xsl:template name="ShapeClass">
    <xsl:if test="a:prstGeom/@prst">
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
        <xsl:when test="a:prstGeom/@prst='circularArrow'">228</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='bentConnector3'">74</xsl:when>
        <xsl:when test="a:prstGeom/@prst='curvedConnector3'">77</xsl:when>
        <!--2011-12-08 罗文甜，增加类别类型-->
        <xsl:when test="a:prstGeom/@prst='curvedConnector4'">77</xsl:when>
        <xsl:when test="a:prstGeom/@prst='homePlate'">219</xsl:when>
        <xsl:when test="a:prstGeom/@prst='quadArrowCallout'">227</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartOnlineStorage'">323</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticTape'">325</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticDrum'">327</xsl:when>

        <xsl:otherwise>11</xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="not(a:prstGeom)">
      <xsl:value-of select="'64'"/>
    </xsl:if>
  </xsl:template>
  <xsl:template name="ShapeName">
    <xsl:if test="a:prstGeom/@prst">
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
        <xsl:when test="a:prstGeom/@prst='bentConnector3'">Elbow Connector</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='circularArrow'">Circular Arrow</xsl:when>
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
        
        <xsl:otherwise>Rectangle</xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="not(a:prstGeom/@prst)">
      <xsl:value-of select="'Curve'"/>
    </xsl:if>
  </xsl:template>
  <xsl:template name="ShapeName2">
    <xsl:if test="a:prstGeom/@prst">
      <xsl:choose>
        <xsl:when test="a:prstGeom/@prst='flowChartSort'">Sort</xsl:when>
        <xsl:when test="a:prstGeom/@prst='rect' or not(a:prstGeom/@prst)">Rectangle</xsl:when>
        <xsl:when test="a:prstGeom/@prst='parallelogram'">Parallelogram</xsl:when>
        <xsl:when test="a:prstGeom/@prst='trapezoid'">Trapezoid</xsl:when>
        <xsl:when test="a:prstGeom/@prst='diamond'">Diamond</xsl:when>
        <xsl:when test="a:prstGeom/@prst='roundRect'">Rounded Rectangle</xsl:when>
        <xsl:when test="a:prstGeom/@prst='octagon'">Octagon</xsl:when>
        <xsl:when test="a:prstGeom/@prst='triangle'">Isosceles Triangle</xsl:when>
        <xsl:when test="a:prstGeom/@prst='rtTriangle'">Right Triangle</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipse'">Oval</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='flowChartInputOutput'">Data</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartPredefinedProcess'">Predefined Process</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartInternalStorage'">Internal Storage</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDocument'">Document</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMultidocument'">Multidocument</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartTerminator'">Terminator</xsl:when>
        <xsl:when test="a:prstGeom/@prst='irregularSeal1'">Explosion 1</xsl:when>
        <xsl:when test="a:prstGeom/@prst='irregularSeal2'">Explosion 2</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star4'">4-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star5'">5-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star8'">8-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star12'">16-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star24'">24-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star32'">32-Point Star</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ribbon2'">Up Ribbon</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeRectCallout'">Rectangular Callout</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeRoundRectCallout'">Rounded Rectangular Callout</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wedgeEllipseCallout'">Oval Callout</xsl:when>
        <xsl:when test="a:prstGeom/@prst='straightConnector1'">Line</xsl:when>
        <xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curve</xsl:when>
        <xsl:when test="a:prstGeom/@prst='bentConnector3'">Elbow Connector</xsl:when>
        <xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curved Connector</xsl:when>
        <xsl:when test="a:prstGeom/@prst='plus'">Cross</xsl:when>
        <xsl:when test="a:prstGeom/@prst='noSmoking'">No Symbol</xsl:when>
        <xsl:when test="a:prstGeom/@prst='homePlate'">Pentagon Arrow</xsl:when>
        <xsl:when test="a:prstGeom/@prst='quadArrowCallout'">Quad Arrow Callout</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartOnlineStorage'">Stored Data</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticTape'">Sequential Access Storage</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout1'">Line Callout1(No Border)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout2'">Line Callout2</xsl:when>
        <xsl:when test="a:prstGeom/@prst='borderCallout3'">Line Callout4</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout1'">Line Callout1(Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout2'">Line Callout3(Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentCallout3'">Line Callout4(Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout1'">Line Callout1</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout2'">Line Callout3(No Border)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='callout3'">Line Callout4(No Border)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout1'">Line Callout1(Border and Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout2'">Line Callout3(Border and Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='accentBorderCallout3'">Line Callout4(Border and Accent Bar)</xsl:when>
        <xsl:when test="a:prstGeom/@prst='cloudCallout'">Parallelogram</xsl:when>
        <xsl:when test="a:prstGeom/@prst='line'">Line</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='circularArrow'">Circular Arrow</xsl:when>
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
        <xsl:when test="a:prstGeom/@prst='flowChartCollate'">Collate</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticDrum'">Direct Access Storage</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartExtract'">Extract</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMerge'">Merge</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDelay'">Delay</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartMagneticDisk'">Magnetic Disk</xsl:when>
        <xsl:when test="a:prstGeom/@prst='flowChartDisplay'">Display</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ribbon'">Down Ribbon</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipseRibbon2'">Curved Up Ribbon</xsl:when>
        <xsl:when test="a:prstGeom/@prst='ellipseRibbon'">Curved Down Ribbon</xsl:when>
        <xsl:when test="a:prstGeom/@prst='verticalScroll'">Vertical Scroll</xsl:when>
        <xsl:when test="a:prstGeom/@prst='horizontalScroll'">Horizontal Scroll</xsl:when>
        <xsl:when test="a:prstGeom/@prst='wave'">Wave</xsl:when>
        <xsl:when test="a:prstGeom/@prst='doubleWave'">Double Wave</xsl:when>
        <xsl:when test="a:prstGeom/@prst='star10' or a:prstGeom/@prst='star12'">16-Point Star</xsl:when>

        <xsl:when test="a:prstGeom/@prst='mathPlus'">Cross</xsl:when>
        <xsl:when test="a:prstGeom/@prst='mathMultiply'">Cross</xsl:when>
        <xsl:when test="a:prstGeom/@prst='mathDivide'">Cross</xsl:when>


        <xsl:when test="a:prstGeom/@prst='arrow'">Straight Connector</xsl:when>
        
        <xsl:otherwise>Rectangle</xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="not(a:prstGeom/@prst)">
      <xsl:value-of select="'Curve'"/>
    </xsl:if>
  </xsl:template>
  <xsl:template match="xdr:spPr|dsp:spPr" mode="second">
    <!--填充-->
    <xsl:choose>
      <xsl:when test="a:pattFill">
        <图:填充_804C>
          <图:图案_800A>
            <xsl:if test="a:pattFill/@prst">
              <xsl:variable name="type">
                <xsl:value-of select="a:pattFill/@prst"/>
              </xsl:variable>
              <xsl:attribute name="类型_8008">
                <xsl:call-template name="ttype">
                  <xsl:with-param name="tttype" select="$type"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <!--图形引用属性-->
            <xsl:if test="a:pattFill/a:fgClr or a:pattFill/a:bgClr">
              <xsl:for-each select="a:pattFill/a:fgClr">
                <xsl:attribute name="前景色_800B">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select="a:pattFill/a:bgClr">
                <xsl:attribute name="背景色_800C">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
              </xsl:for-each>
            </xsl:if>
          </图:图案_800A>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:solidFill/a:srgbClr">
        <图:填充_804C>
          <图:颜色_8004>
            <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
          </图:颜色_8004>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="ancestor::xdr:sp/preceding-sibling::xdr:grpSpPr/a:solidFill">
        <图:填充_804C>
          <图:颜色_8004>
            <!-- update by xuzhenwei BUG_2727: 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。-->
            <xsl:choose>
              <xsl:when test="ancestor::xdr:sp/preceding-sibling::xdr:grpSpPr/a:solidFill/a:srgbClr">
                <xsl:value-of select="concat('#',ancestor::xdr:sp/preceding-sibling::xdr:grpSpPr/a:solidFill/a:srgbClr/@val)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:variable name="schemeid" select="ancestor::xdr:sp/preceding-sibling::xdr:grpSpPr/a:solidFill/a:schemeClr/@val"/>
                <xsl:choose>
                  <xsl:when test="$schemeid='dk1'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='lt1'">
                    <xsl:value-of select="'#FFFFFF'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='dk2'">
                    <xsl:value-of select="'#1F497D'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='lt2'">
                    <xsl:value-of select="'#EEECE1'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent1'">
                    <xsl:value-of select="'#4F81BD'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent2'">
                    <xsl:value-of select="'#C0504D'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent3'">
                    <xsl:value-of select="'#9BBB59'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent4'">
                    <xsl:value-of select="'#8064A2'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent5'">
                    <xsl:value-of select="'#4BACC6'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='accent6'">
                    <xsl:value-of select="'#F79646'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='hlink'">
                    <xsl:value-of select="'#0000FF'"/>
                  </xsl:when>
                  <xsl:when test="$schemeid='folHlink'">
                    <xsl:value-of select="'#800080'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'auto'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </图:颜色_8004>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:solidFill/a:schemeClr">
        <xsl:variable name="id">
          <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
        </xsl:variable>
        <图:填充_804C>
          <图:颜色_8004>
            <xsl:choose>
              <xsl:when test="$id='dk1'">
                <xsl:value-of select="'#000000'"/>
              </xsl:when>
              <xsl:when test="$id='lt1'">
                <xsl:value-of select="'#FFFFFF'"/>
              </xsl:when>
              <xsl:when test="$id='dk2'">
                <xsl:value-of select="'#1F497D'"/>
              </xsl:when>
              <xsl:when test="$id='lt2'">
                <xsl:value-of select="'#EEECE1'"/>
              </xsl:when>
              <xsl:when test="$id='accent1'">
                <xsl:value-of select="'#4F81BD'"/>
              </xsl:when>
              <xsl:when test="$id='accent2'">
                <xsl:value-of select="'#C0504D'"/>
              </xsl:when>
              <xsl:when test="$id='accent3'">
                <xsl:value-of select="'#9BBB59'"/>
              </xsl:when>
              <xsl:when test="$id='accent4'">
                <xsl:value-of select="'#8064A2'"/>
              </xsl:when>
              <xsl:when test="$id='accent5'">
                <xsl:value-of select="'#4BACC6'"/>
              </xsl:when>
              <xsl:when test="$id='accent6'">
                <xsl:value-of select="'#F79646'"/>
              </xsl:when>
              <xsl:when test="$id='hlink'">
                <xsl:value-of select="'#0000FF'"/>
              </xsl:when>
              <xsl:when test="$id='folHlink'">
                <xsl:value-of select="'#800080'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'auto'"/>
              </xsl:otherwise>
            </xsl:choose>
          </图:颜色_8004>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:gradFill">
        <图:填充_804C>
          <xsl:apply-templates select="a:gradFill"/>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:blipFill">
        <xsl:variable name="yuding1">
          <xsl:value-of select="a:blipFill/a:blip/@ r:embed"/>
        </xsl:variable>
        <xsl:variable name="yuding2">
          <xsl:value-of select="ancestor::xdr:wsDr/pr:Relationships/pr:Relationship[@Id=$yuding1]/@Target"/>
        </xsl:variable>
        <xsl:variable name="yuding3">
          <xsl:value-of select="substring-after($yuding2,'media/')"/>
        </xsl:variable>
        <图:填充_804C>
          <图:图片_8005>
            <xsl:if test="a:blipFill/a:stretch">
              <xsl:attribute name="位置_8006">
                <xsl:value-of select="'stretch'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="图形引用_8007">
              <xsl:value-of select="substring-before($yuding3,'.')"/>
            </xsl:attribute>
            <!--图片的 类型，名称 属性-->
          </图:图片_8005>
        </图:填充_804C>
      </xsl:when>
      <xsl:when test="a:noFill">
        <!--
          <图:填充 uof:locID="g0012">
            <图:颜色_8004>
              <xsl:value-of select="'auto'"/>
            </图:颜色_8004>
          </图:填充>
          -->
      </xsl:when>
      <xsl:when test="a:custGeom">
        <xsl:if test="./a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@x = ./a:custGeom/a:pathLst/a:path/a:lnTo[position()=last()]/a:pt/@x and ./a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@y = ./a:custGeom/a:pathLst/a:path/a:lnTo[position()=last()]/a:pt/@y">
          <xsl:if test="not(a:solidFill) and not(a:pattFill)">
            <图:填充_804C>
              <图:颜色_8004>
                <xsl:value-of select="'#4F81BD'"/>
              </图:颜色_8004>
            </图:填充_804C>
          </xsl:if>
        </xsl:if>
      </xsl:when>
      <xsl:when test="not(a:solidFill or a:gradFill) and ./following-sibling::xdr:style and a:prstGeom/@prst!='bracketPair' and a:prstGeom/@prst!='bracePair' and a:prstGeom/@prst!='leftBracket' and a:prstGeom/@prst!='rightBracket' and a:prstGeom/@prst!='leftBrace' and a:prstGeom/@prst!='rightBrace'">
        <xsl:choose>
          <xsl:when test="./following-sibling::xdr:style/a:fillRef/a:schemeClr">
            <xsl:variable name="id">
              <xsl:value-of select="./following-sibling::xdr:style/a:fillRef/a:schemeClr/@val"/>
            </xsl:variable>
            <图:填充_804C>
              <图:颜色_8004>
                <xsl:choose>
                  <xsl:when test="$id='dk1'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  <xsl:when test="$id='lt1'">
                    <xsl:value-of select="'#FFFFFF'"/>
                  </xsl:when>
                  <xsl:when test="$id='dk2'">
                    <xsl:value-of select="'#1F497D'"/>
                  </xsl:when>
                  <xsl:when test="$id='lt2'">
                    <xsl:value-of select="'#EEECE1'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent1'">
                    <xsl:value-of select="'#4F81BD'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent2'">
                    <xsl:value-of select="'#C0504D'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent3'">
                    <xsl:value-of select="'#9BBB59'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent4'">
                    <xsl:value-of select="'#8064A2'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent5'">
                    <xsl:value-of select="'#4BACC6'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent6'">
                    <xsl:value-of select="'#F79646'"/>
                  </xsl:when>
                  <xsl:when test="$id='hlink'">
                    <xsl:value-of select="'#0000FF'"/>
                  </xsl:when>
                  <xsl:when test="$id='folHlink'">
                    <xsl:value-of select="'#800080'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'auto'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </图:颜色_8004>
            </图:填充_804C>
          </xsl:when>
          <xsl:when test="./following-sibling::xdr:style/a:fillRef/a:srgbClr">
            <图:填充_804C>
              <图:颜色_8004>
                <xsl:value-of select="concat('#',./following-sibling::xdr:style/a:fillRef/a:srgbClr/@val)"/>
              </图:颜色_8004>
            </图:填充_804C>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
    <xsl:if test="a:xfrm/@rot">
      <图:旋转角度_804D>
        <xsl:value-of select="360-(21600000-a:xfrm/@rot) div 60000"/>
      </图:旋转角度_804D>
    </xsl:if>
    <xsl:if test="a:scene3d/a:camera/a:rot">
      <图:旋转角度_804D>
        <xsl:value-of select="(21600000-a:scene3d/a:camera/a:rot/@rev) div 60000"/>
      </图:旋转角度_804D>
    </xsl:if>
    <图:是否打印对象_804E>true</图:是否打印对象_804E>
    <xsl:if test="ancestor::p:sp/p:nvSpPr/p:cNvPr/@descr">
      <图:Web文字_804F>
        <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:cNvPr/@descr"/>
      </图:Web文字_804F>
    </xsl:if>
    <xsl:if test="ancestor::p:cxnSp/p:nvCxnSpPr/p:cNvPr/@descr">
      <图:Web文字_804F>
        <xsl:value-of select="ancestor::p:cxnSp/p:nvCxnSpPr/p:cNvPr/@descr"/>
      </图:Web文字_804F>
    </xsl:if>
    <!--添加图形背景透明度。李杨2012-3-14-->
    <xsl:if test ="a:solidFill/a:srgbClr/a:alpha/@val">
      <图:透明度_8050>
        <xsl:variable name ="tm">
          <xsl:value-of select ="a:solidFill/a:srgbClr/a:alpha/@val"/>
        </xsl:variable>
        <xsl:value-of select ="number(100000-$tm) div 1000"/>
      </图:透明度_8050>
    </xsl:if>
    <!--以下代码取的是线条的透明度，UOF中只有图背景的透明度。李杨2012-3-14-->
    <!--<xsl:if test=".//a:alpha">
      <图:透明度_8050>
        <xsl:variable name="tm">
          <xsl:value-of select=".//a:alpha/@val"/>
        </xsl:variable>
        <xsl:value-of select="number(100000 - $tm) div 1000"/>
      </图:透明度_8050>
    </xsl:if>-->
    <!--阴影  李杨2011-12-09-->
    <xsl:if test="a:effectLst/a:outerShdw">
      <!--外部阴影-->
      <xsl:apply-templates select="a:effectLst/a:outerShdw" mode="shadow"/>
    </xsl:if>
    <xsl:if test="a:effectLst/a:innerShdw">
      <!--内部阴影-->
      <xsl:apply-templates select="a:effectLst/a:innerShdw" mode="shadow"/>
    </xsl:if>
    <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
    <!--线-->
    <xsl:call-template name="xiantiao"/>
    <!--箭头 包含在xiantiao模板中-->
    <xsl:if test="a:xfrm/a:ext">
      <图:大小_8060>
        <xsl:attribute name ="长_C604">
          <!--修改图片高度（长）李杨 2012-5-24-->
          <!--<xsl:value-of select="a:xfrm/a:ext/@cy div 12800"/>-->
          <xsl:value-of select="a:xfrm/a:ext/@cy div 12700"/>
        </xsl:attribute>
        <xsl:attribute name ="宽_C605">
          <!-- 2012-11-27, xuzhenwei, 解决预定义图形的位置与原来不一致的BUG start -->
          <!--<xsl:choose>-->
            <!-- 判断预定义图形是否拉伸，暂时先通过width大于20来判断 -->
            <!--<xsl:when test="ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col/@width &gt; 20">
              <xsl:variable name="off" select="a:xfrm/a:off/@x"/>
              <xsl:variable name="etx" select="a:xfrm/a:ext/@cx"/>
              <xsl:variable name="distance" select="($etx - $off) div 12700"/>
              <xsl:variable name="distance01">
                <xsl:choose>
                  <xsl:when test="ancestor::ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:twoCellAnchor/xdr:to">
                    <xsl:value-of select="ancestor::ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:twoCellAnchor/xdr:to/xdr:colOff div 12700"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="ancestor::ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:oneCellAnchor/xdr:from/xdr:rowOff div 12700"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable> 
              <xsl:variable name="pull" select="ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col/@width div 8.38 * 54"/>
              <xsl:value-of select="$pull + $distance01"/>
            </xsl:when>
            <xsl:otherwise>-->
              <xsl:value-of select="a:xfrm/a:ext/@cx div 12700"/>
            <!--</xsl:otherwise>
          </xsl:choose>-->
          <!-- end -->
        </xsl:attribute>
      </图:大小_8060>
    </xsl:if>
    <!--三维效果-->
    <xsl:if test="a:scene3d">
      <图:三维效果_8061>
        <xsl:apply-templates select="a:scene3d"/>
      </图:三维效果_8061>
    </xsl:if>
    <!--图片属性-->
    <!--艺术字-->
  </xsl:template>
  <xsl:template name="xiantiao">
      <xsl:if test="a:ln">
        <xsl:for-each select="a:ln[not(a:noFill)]">
          <图:线_8057>
            <xsl:if test="a:solidFill">
              <图:线颜色_8058>
                <xsl:if test="a:solidFill[a:srgbClr]">
                  <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
                </xsl:if>
                <xsl:if test="a:solidFill[a:schemeClr]">
                  <xsl:variable name="id">
                    <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="$id='dk1'">
                      <xsl:value-of select="'#000000'"/>
                    </xsl:when>
                    <!-- 20130402 update by xuzhenwei BUG_2727 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start -->

                    <!--2014-4-24, update by Qihy, 修复bug3154，smartArt边框线颜色不正确， start-->
                    <xsl:when test="$id='lt1' and not(parent::xdr:spPr[a:sp3d])">
                      <!--<xsl:value-of select="'#385D8A'"/>-->
                      <xsl:value-of select="'#FFFFFF'"/>
                      <!--2014-4-24 end-->

                      <!--ffffff-->
                    </xsl:when>
                    <!-- end -->
                    <xsl:when test="$id='dk2'">
                      <xsl:value-of select="'#1F497D'"/>
                    </xsl:when>
                    <xsl:when test="$id='lt2'">
                      <xsl:value-of select="'#EEECE1'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent1'">
                       <xsl:value-of select="'#4F81BD'"/> 
                    </xsl:when>
                    <xsl:when test="$id='accent2'">
                      <xsl:value-of select="'#C0504D'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent3'">
                      <xsl:value-of select="'#9BBB59'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent4'">
                      <xsl:value-of select="'#8064A2'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent5'">
                      <xsl:value-of select="'#4BACC6'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent6'">
                      <xsl:value-of select="'#F79646'"/>
                    </xsl:when>
                    <xsl:when test="$id='hlink'">
                      <xsl:value-of select="'#0000FF'"/>
                    </xsl:when>
                    <xsl:when test="$id='folHlink'">
                      <xsl:value-of select="'#800080'"/>
                    </xsl:when>
                    <!-- 20130402 update by xuzhenwei BUG_2727 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start -->
                    <xsl:when test="parent::xdr:spPr/a:sp3d/a:extrusionClr or parent::xdr:spPr/a:sp3d/a:contourClr">
                      <xsl:choose>
                        <xsl:when test="parent::xdr:spPr/a:sp3d/a:contourClr/a:srgbClr/@val">
                          <xsl:value-of select ="concat('#',parent::xdr:spPr/a:sp3d/a:contourClr/a:srgbClr/@val)"/>
                        </xsl:when>
                        <xsl:when test="parent::xdr:spPr/a:sp3d/a:extrusionClr/a:srgbClr/@val">
                          <xsl:value-of select ="concat('#',parent::xdr:spPr/a:sp3d/a:extrusionClr/a:srgbClr/@val)"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="'auto'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!-- end -->
                    <xsl:otherwise>
                      <xsl:value-of select="'auto'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if>
              </图:线颜色_8058>
            </xsl:if>
            <xsl:if test="not(a:solidFill) and following::xdr:style/a:lnRef[a:schemeClr or a:srgbClr]">
              <图:线颜色_8058>
                <xsl:if test ="following::xdr:style/a:lnRef/a:schemeClr[@val]">
                  <xsl:variable name ="id">
                    <xsl:value-of select ="following::xdr:style/a:lnRef/a:schemeClr/@val"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="$id='dk1'">
                      <xsl:value-of select="'#000000'"/>
                    </xsl:when>
                    <xsl:when test="$id='lt1'">
                      <xsl:value-of select="'#FFFFFF'"/>
                    </xsl:when>
                    <xsl:when test="$id='dk2'">
                      <xsl:value-of select="'#1F497D'"/>
                    </xsl:when>
                    <xsl:when test="$id='lt2'">
                      <xsl:value-of select="'#EEECE1'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent1'">
                      <xsl:value-of select="'#4F81BD'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent2'">
                      <xsl:value-of select="'#C0504D'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent3'">
                      <xsl:value-of select="'#9BBB59'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent4'">
                      <xsl:value-of select="'#8064A2'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent5'">
                      <xsl:value-of select="'#4BACC6'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent6'">
                      <xsl:value-of select="'#F79646'"/>
                    </xsl:when>
                    <xsl:when test="$id='hlink'">
                      <xsl:value-of select="'#0000FF'"/>
                    </xsl:when>
                    <xsl:when test="$id='folHlink'">
                      <xsl:value-of select="'#800080'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'#385D8A'"/>
                      <!--<xsl:value-of select="'auto'"/>-->
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:if>
                <!--<xsl:value-of select="'#385D8A'"/>-->
              </图:线颜色_8058>
            </xsl:if>
            <!--添加OO中线条渐变色，转换为UOF中一种颜色，因为UOF中线颜色没有渐变色定义。李杨2012-3-14-->
            <xsl:if test ="not(a:solidFill) and a:gradFill">
              <图:线颜色_8058>
                <xsl:value-of select ="concat('#',a:gradFill/a:gsLst/a:gs[position()=2]/a:srgbClr/@val)"/>
              </图:线颜色_8058>
            </xsl:if>
            <xsl:if test="a:prstDash/@val and not(@cmpd)">
              <图:线类型_8059>
                <xsl:attribute name="线型_805A">
                  <xsl:value-of select="'single'"/>
                </xsl:attribute>
                <xsl:attribute name="虚实_805B">
                  <xsl:choose>
                    <xsl:when test="a:prstDash/@val='solid'">
                      <xsl:value-of select="'solid'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dot'">
                      <xsl:value-of select="'square-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dash'">
                      <xsl:value-of select="'dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDash'">
                      <xsl:value-of select="'long-dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDashDotDot'">
                      <xsl:value-of select="'dash-dot-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDash'">
                      <xsl:value-of select="'dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDot'">
                      <xsl:value-of select="'square-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDashDotDot'">
                      <xsl:value-of select="'dash-dot-dot'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
              </图:线类型_8059>
            </xsl:if>
            <xsl:if test="not(a:prstDash/@val) and @cmpd">
              <图:线类型_8059>
                <xsl:attribute name="线型_805A">
                  <xsl:choose>
                    <xsl:when test="@cmpd='dbl'">
                      <xsl:value-of select="'double'"/>
                    </xsl:when>
                    <xsl:when test="@cmpd='sng'">
                      <xsl:value-of select="'single'"/>
                    </xsl:when>
                    <!-- 20130402 update by xuzhenwei BUG_2727 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start -->
                    <!--上细下粗-->
                    <xsl:when test="@cmpd='thickThin'">
                      <xsl:value-of select="'thin-thick'"/>
                    </xsl:when>
                    <!--上粗下细-->
                    <xsl:when test="@cmpd='thinThick'">
                      <xsl:value-of select="'thick-thin'"/>
                    </xsl:when>
                    <!--三线-->
                    <xsl:when test="@cmpd='tri'">
                      <xsl:value-of select="'thick-between-thin'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="虚实_805B">
                  <xsl:value-of select="solid"/>
                </xsl:attribute>
                <!-- end -->
              </图:线类型_8059>
            </xsl:if>
            <xsl:if test="a:prstDash/@val and @cmpd">
              <图:线类型_8059>
                <xsl:attribute name="线型_805A">
                  <xsl:choose>
                    <xsl:when test="@cmpd='dbl'">
                      <xsl:value-of select="'double'"/>
                    </xsl:when>
                    <xsl:when test="@cmpd='sng'">
                      <xsl:value-of select="'single'"/>
                    </xsl:when>
                    <!-- 20130402 update by xuzhenwei BUG_2727 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start -->
                    <!--上细下粗-->
                    <xsl:when test="@cmpd='thickThin'">
                      <xsl:value-of select="'thin-thick'"/>
                    </xsl:when>
                    <!--上粗下细-->
                    <xsl:when test="@cmpd='thinThick'">
                      <xsl:value-of select="'thick-thin'"/>
                    </xsl:when>
                    <!--三线-->
                    <xsl:when test="@cmpd='tri'">
                      <xsl:value-of select="'thick-between-thin'"/>
                    </xsl:when>
                    <!-- end -->
                  </xsl:choose>
                </xsl:attribute>

                <xsl:attribute name="虚实_805B">
                  <xsl:choose>
                    <xsl:when test="a:prstDash/@val='solid'">
                      <xsl:value-of select="'solid'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dot'">
                      <xsl:value-of select="'square-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dash'">
                      <xsl:value-of select="'dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDash'">
                      <xsl:value-of select="'long-dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='dashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='lgDashDotDot'">
                      <xsl:value-of select="'dash-dot-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDash'">
                      <xsl:value-of select="'dash'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDot'">
                      <xsl:value-of select="'square-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDashDot'">
                      <xsl:value-of select="'dash-dot'"/>
                    </xsl:when>
                    <xsl:when test="a:prstDash/@val='sysDashDotDot'">
                      <xsl:value-of select="'dash-dot-dot'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
              </图:线类型_8059>
            </xsl:if>
            <xsl:if test="not(a:prstDash/@val) and not(@cmpd) and following::xdr:style">
              <图:线类型_8059>
                <xsl:variable name="lnstyleid">
                  <xsl:value-of select="following::xdr:style/a:lnRef/@idx"/>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/a:prstDash/@val">
                    <xsl:for-each select="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]">

                      <xsl:attribute name="线型_805A">
                        <xsl:value-of select="'single'"/>
                      </xsl:attribute>
                      <xsl:attribute name="虚实_805B">
                        <xsl:choose>
                          <xsl:when test="a:prstDash/@val='solid'">
                            <xsl:value-of select="'solid'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='dot'">
                            <xsl:value-of select="'square-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='dash'">
                            <xsl:value-of select="'dash'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='lgDash'">
                            <xsl:value-of select="'long-dash'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='dashDot'">
                            <xsl:value-of select="'dash-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='lgDashDot'">
                            <xsl:value-of select="'dash-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='lgDashDotDot'">
                            <xsl:value-of select="'dash-dot-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='sysDash'">
                            <xsl:value-of select="'dash'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='sysDot'">
                            <xsl:value-of select="'square-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='sysDashDot'">
                            <xsl:value-of select="'dash-dot'"/>
                          </xsl:when>
                          <xsl:when test="a:prstDash/@val='sysDashDotDot'">
                            <xsl:value-of select="'dash-dot-dot'"/>
                          </xsl:when>
                        </xsl:choose>
                      </xsl:attribute>
                    </xsl:for-each>
                    <!--<xsl:value-of select="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/a:prstDash/@val"/>-->
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="线型_805A">
                      <xsl:value-of select="'single'"/>
                    </xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
              </图:线类型_8059>
            </xsl:if>
            <xsl:if test="@w">
              <图:线粗细_805C>
                <xsl:value-of select="@w div 12700"/>
              </图:线粗细_805C>
            </xsl:if>
            <xsl:if test="not(@w) and following::xdr:style">
              <图:线粗细_805C>
                <xsl:variable name="lnstyleid">
                  <xsl:value-of select="following::xdr:style/a:lnRef/@idx"/>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/@w">
                    <xsl:value-of select="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/@w div 12700"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0.75'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </图:线粗细_805C>
            </xsl:if>
          </图:线_8057>
          <xsl:if test ="a:headEnd or a:tailEnd">
            <图:箭头_805D>
              <xsl:if test="a:headEnd">
                <图:前端箭头_805E>
                  <图:式样_8000>
                    <xsl:choose>
                      <xsl:when test="a:headEnd/@type='triangle'">normal</xsl:when>
                      <xsl:when test="a:headEnd/@type='diamond'">diamond</xsl:when>
                      <xsl:when test="a:headEnd/@type='arrow'">open</xsl:when>
                      <xsl:when test="a:headEnd/@type='oval'">oval</xsl:when>
                      <xsl:when test="a:headEnd/@type='stealth'">stealth</xsl:when>
                    </xsl:choose>
                  </图:式样_8000>
                  <图:大小_8001>5</图:大小_8001>
                </图:前端箭头_805E>
              </xsl:if>
              <xsl:if test="a:tailEnd">
                <图:后端箭头_805F>
                  <图:式样_8000>
                    <xsl:choose>
                      <xsl:when test="a:tailEnd/@type='triangle'">normal</xsl:when>
                      <xsl:when test="a:tailEnd/@type='diamond'">diamond</xsl:when>
                      <xsl:when test="a:tailEnd/@type='arrow'">open</xsl:when>
                      <xsl:when test="a:tailEnd/@type='oval'">oval</xsl:when>
                      <xsl:when test="a:tailEnd/@type='stealth'">stealth</xsl:when>
                    </xsl:choose>
                  </图:式样_8000>
                  <图:大小_8001>5</图:大小_8001>
                </图:后端箭头_805F>
              </xsl:if>
            </图:箭头_805D>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="not(a:ln) and following-sibling::xdr:style">
        <图:线_8057>
          <xsl:if test="following-sibling::xdr:style/a:lnRef[a:schemeClr or a:srgbClr]">
            <图:线颜色_8058>
              <!-- 20130402 update by xuzhenwei  BUG_2727 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start -->
              <xsl:choose>
                <xsl:when test="following-sibling::xdr:style/a:lnRef/a:srgbClr">
                  <xsl:value-of select="concat('#',following-sibling::xdr:style[1]/a:lnRef/a:srgbClr/@val)"/>
                </xsl:when>
                <xsl:when test="following-sibling::xdr:style/a:lnRef/a:schemeClr">
                  <xsl:variable name="id">
                    <xsl:value-of select="following-sibling::xdr:style[1]/a:lnRef/a:schemeClr/@val"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="$id='dk1'">
                      <xsl:value-of select="'#000000'"/>
                    </xsl:when>
                    <xsl:when test="$id='lt1'">
                      <xsl:value-of select="'#FFFFFF'"/>
                    </xsl:when>
                    <xsl:when test="$id='dk2'">
                      <xsl:value-of select="'#1F497D'"/>
                    </xsl:when>
                    <xsl:when test="$id='lt2'">
                      <xsl:value-of select="'#EEECE1'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent1'">

                      <!-- 2014-3-17 Qihy 默认图形轮廓颜色不正确 start-->
                      <xsl:value-of select="'#333399'"/>
                      <!-- <xsl:value-of select="'#4F81BD'"/> -->
                      <!-- 2014-3-17  end-->
                      
                    </xsl:when>
                    <xsl:when test="$id='accent2'">
                      <xsl:value-of select="'#C0504D'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent3'">
                      <xsl:value-of select="'#9BBB59'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent4'">
                      <xsl:value-of select="'#8064A2'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent5'">
                      <xsl:value-of select="'#4BACC6'"/>
                    </xsl:when>
                    <xsl:when test="$id='accent6'">
                      <xsl:value-of select="'#F79646'"/>
                    </xsl:when>
                    <xsl:when test="$id='hlink'">
                      <xsl:value-of select="'#0000FF'"/>
                    </xsl:when>
                    <xsl:when test="$id='folHlink'">
                      <xsl:value-of select="'#800080'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'auto'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
              </xsl:choose>
              <!-- end -->
            </图:线颜色_8058>
          </xsl:if>
          <xsl:if test="following-sibling::xdr:style">
            <图:线类型_8059>
              <xsl:attribute name ="线型_805A">
                <xsl:variable name="lnstyleid">
                  <xsl:value-of select="following-sibling::xdr:style/a:lnRef/@idx"/>
                </xsl:variable>
                <xsl:choose>
                  <!-- update by xuzhenwei BUG_2727: 集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。 start OO中的线型只有四种类型：sng、dbl、tri、thickThin和thinThick -->
                  <xsl:when test="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/a:prstDash[not(@val='solid')]">
                    <xsl:value-of select="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/a:prstDash/@val"/>
                  </xsl:when>
                  <!-- end -->
                  <xsl:otherwise>
                    <xsl:value-of select="'single'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </图:线类型_8059>
          </xsl:if>
          <xsl:if test="following-sibling::xdr:style">
            <图:线粗细_805C>
              <xsl:variable name="lnstyleid">
                <xsl:value-of select="following-sibling::xdr:style/a:lnRef/@idx"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/@w">
                  <xsl:value-of select="//a:fmtScheme/a:lnStyleLst/a:ln[position()=$lnstyleid]/@w div 12700"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'0.75'"/>
                </xsl:otherwise>
              </xsl:choose>
            </图:线粗细_805C>
          </xsl:if>
        </图:线_8057>
      </xsl:if>
  </xsl:template>
  <xsl:template match="a:bodyPr">
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
      <!--修改属性：是否大小适应文字 李杨2011-12-09-->
      <!--<xsl:if test="concat(name(a:spAutofit),name(a:normAutofit))!=''">-->
      <!--<xsl:if test ="$as='a:spAutofit'">-->
      <xsl:choose >
        <xsl:when test ="./a:spAutoFit or ./a:normAutofit">
          <xsl:value-of select="'true'"/>
        </xsl:when>
        <xsl:when test ="./a:spAutoFit and not(./a:normAutofit)">
          <xsl:value-of select="'true'"/>
        </xsl:when>
        <xsl:when test ="not(./a:spAutoFit) and ./a:normAutofit">
          <xsl:value-of select ="'true'"/>
        </xsl:when>
        <xsl:otherwise >
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
      <!--<xsl:if test ="./a:spAutoFit or ./a:normAutofit">
        <xsl:value-of select="'true'"/>
      </xsl:if>
      <xsl:if test ="./a:spAutoFit and not(./a:normAutofit)">
        <xsl:value-of select="'true'"/>
      </xsl:if>
      <xsl:if test ="not(./a:spAutoFit) and ./a:normAutofit">
        <xsl:value-of select="'true'"/>
      </xsl:if>
      --><!--<xsl:if test="concat(name(a:spAutofit),name(a:normAutofit))=''">--><!--
      <xsl:if test ="not(./a:spAutofit) and not(./a:normAutofit)">
        <xsl:value-of select="'false'"/>
      </xsl:if>-->
    </xsl:attribute>
    <xsl:attribute name ="是否文字随对象旋转_8049">
      <xsl:value-of select ="'true'"/>
    </xsl:attribute>
    <xsl:attribute name ="是否文字绕排外部对象_8068">
      <xsl:value-of select ="'true'"/>
    </xsl:attribute>
    
    <!--2014-3-6, add, Qihy 修复BUG2986-文本框_预定义图形-内部边距，内部边距功能失效 start-->
    <图:边距_803D>
      <xsl:if test="@lIns">
        <xsl:attribute name="左_C608">
          <xsl:value-of select="./@lIns div 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@tIns">
        <xsl:attribute name="上_C609">
          <xsl:value-of select="./@tIns div 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@rIns">
        <xsl:attribute name="右_C60A">
          <xsl:value-of select="./@rIns div 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@bIns">
        <xsl:attribute name="下_C60B">
          <xsl:value-of select="./@bIns div 12700"/>
        </xsl:attribute>
      </xsl:if>
    </图:边距_803D>
    <!--end-->
    
    <图:对齐_803E>
      <xsl:attribute name="水平对齐_421D">
        <xsl:choose>
          <xsl:when test="./@vert='eaVert' or ./@vert='vert'">
            <xsl:value-of select="'right'"/>
          </xsl:when>
          <xsl:when test="./@anchorCtr='1'">
            <xsl:value-of select="'center'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'left'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="./@anchor">
        <xsl:attribute name="文字对齐_421E">
          <xsl:choose>
            <xsl:when test="./@anchor='b'">
              <xsl:value-of select="'bottom'"/>
            </xsl:when>
            <xsl:when test="./@anchor='ctr' or ./@anchor='dist'">
              <!--<xsl:value-of select="'middle'"/>-->
              <xsl:value-of select ="'center'"/>
            </xsl:when>
            <xsl:when test="./@anchor='t'">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:when test="./@anchor='just'">
              <!--<xsl:value-of select="'justify'"/>-->
              <xsl:value-of select ="base"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </图:对齐_803E>
    <xsl:if test="./@vert">
      <!--修改文字排列方向  李杨2011-12-12-->
      <xsl:element name="图:文字排列方向_8042">
        <xsl:choose>
          <xsl:when test="./@vert='horz'">
            <xsl:value-of select="'t2b-l2r-0e-0w'"/>
          </xsl:when>
          <xsl:when test ="./@vert='vert'">
            <xsl:value-of select="'r2l-t2b-90e-90w'"/>
          </xsl:when>
          <xsl:when test="./@vert='eaVert' or ./@vert='wordArtVertRtl'">
            <xsl:value-of select="'r2l-t2b-0e-90w'"/>
          </xsl:when>
          <xsl:when test="./@vert='vert270'">
            <xsl:value-of select="'l2r-b2t-270e-270w'"/>
          </xsl:when>
          <xsl:when test ="./@vert='wordArtVert'">
            <!-- update xuzhenwei 2012-11-21 解决OOXML到UOF垂直文本框变成横排文本的BUG start -->
            <!-- <xsl:value-of select ="'r2l-t2b-0e-90w'"/> -->
            <xsl:value-of select ="'l2r-t2b-0e-90w'"/>
            <!-- update xuzhenwei 2012-11-21 解决OOXML到UOF垂直文本框变成横排文本的BUG end -->
          </xsl:when>
        </xsl:choose>
      </xsl:element>
    </xsl:if>
  </xsl:template>
  <xsl:template match="a:p">
    <!-- add by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致  -->
    <xsl:param name="positi"/>
    <!-- end -->
    <!--修改文本框中的文本内容 李杨2011-12-12-->
    <!--<图:内容_8043>-->
      <字:段落_416B>
        <xsl:if test="a:pPr">
          <xsl:apply-templates select="a:pPr"/>
        </xsl:if>
        <!--段落属性-->
        <xsl:if test="a:r">
          <xsl:apply-templates select="a:r" mode="ss">
            <xsl:with-param name="position" select="$positi"/>
          </xsl:apply-templates>
        </xsl:if>
        <!--句-->
      </字:段落_416B>
    <!--</图:内容_8043>-->
  </xsl:template>
  <xsl:template match="a:r" mode="ss">
    <!-- add by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致  -->
    <xsl:param name="position"/>
    <!-- end -->
    <字:句_419D>
      <xsl:if test="a:rPr">
        <xsl:apply-templates select="a:rPr">
          <xsl:with-param name="oneCellposition" select="$position"/>
        </xsl:apply-templates>
      </xsl:if>
      <!--句属性-->
      <xsl:if test="a:t">
        <xsl:apply-templates select="a:t"/>
      </xsl:if>
      
      		<!--20121127,gaoyuwei，解决OOXML到UOF文本框中中英文之间换行符丢失的问题BUG start-->
		<xsl:if test ="./following-sibling::a:br">
			<!--<xsl:for-each select="../a:br">-->
			<字:换行符_415F/>
		</xsl:if>
		<!--end-->
      
      <!--文本-->
    </字:句_419D>
  </xsl:template>
  <xsl:template match="a:t">
    <字:文本串_415B>
      <xsl:value-of select="."/>
    </字:文本串_415B>
  </xsl:template>
  <xsl:template match="a:pPr">
    <字:段落属性_419B>
      <xsl:if test="a:rPr">
        <xsl:apply-templates select="a:rPr"/>
      </xsl:if>
      <xsl:call-template name="pPr-body"/>
      <!--段落对齐-->
    </字:段落属性_419B>
  </xsl:template>
  <xsl:template match="a:rPr">
    <!-- add by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致  -->
    <xsl:param name="oneCellposition"/>
    <!-- end -->
    <字:句属性_4158>
      <!-- 203130318 add by xuzhenwei BUG_2758:功能测试 文件转换失败 start -->
      <xsl:if test="a:ln/a:solidFill or a:latin or a:ea or @sz or a:solidFill or preceding::xdr:style/a:fontRef">
        <!-- end -->
        <字:字体_4128>
          					
			<!--20121217,gaoyuwei，解决OOXML到UOF文本框中艺术字大小不正确BUG（实际是字体丢失） start-->
			<!--判断是文本框字体-->
			<xsl:if test="ancestor::xdr:txBody">
			<!--如果不存在a:latin or a:ea，则引用OOXML默认的字体calibri,中文默认宋体-->
				<xsl:if test="not(a:latin or a:ea  or a:cs)">
					<xsl:attribute name="西文字体引用_4129">
						<xsl:value-of select="'font_00002'"/>
					</xsl:attribute>
					<xsl:attribute name="中文字体引用_412A">
						<xsl:value-of select="'font_00001'"/>
					</xsl:attribute>
				</xsl:if>
			<!--如果存在a:latin or a:ea，则引用OOXML中对应的字体-->
				<xsl:if test="a:latin or a:ea  or a:cs">
		
					<!--定义变量-->
					<xsl:variable name="sp_id">
						<xsl:if test="ancestor::xdr:twoCellAnchor[xdr:sp]">
							<xsl:value-of select="ancestor::xdr:sp/xdr:nvSpPr/xdr:cNvPr/@id"/>
						</xsl:if>
					</xsl:variable>
          <!-- update by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致。此处代码删除掉  -->
          <!--<xsl:variable name="node_pos">
						<xsl:value-of select="position()"/>
					</xsl:variable>-->
          <!-- end -->
					<xsl:variable name="zitiyy">
						<xsl:value-of select="concat($sp_id,$oneCellposition)"/>
					</xsl:variable>
          
          <!--2014-3-29, update by Qihy, 文本框中字体不正确，start-->
					<!--<xsl:if test="a:latin or a:ea  or a:cs">
						<xsl:attribute name="西文字体引用_4129">
							<xsl:value-of select="concat($zitiyy,'latin')"/>
						</xsl:attribute>
						<xsl:attribute name="中文字体引用_412A">
							<xsl:value-of select="concat($zitiyy,'ea')"/>
						</xsl:attribute>
					</xsl:if>-->
          <xsl:if test="a:latin">
            <xsl:attribute name="西文字体引用_4129">
              <xsl:value-of select="concat($zitiyy,'latin')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="a:ea  or a:cs">
            <xsl:attribute name="中文字体引用_412A">
              <xsl:value-of select="concat($zitiyy,'ea')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
			</xsl:if>
      <xsl:if test="@lang and @lang ='en-US'">
        <xsl:attribute name="是否西文绘制_412C">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@lang and @lang = 'zh-CN'">
        <xsl:attribute name="是否西文绘制_412C">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <!--2014-3-29, end-->

          <!--end-->
          
          
          <xsl:if test="@sz">
            <xsl:attribute name="字号_412D">
              <xsl:value-of select="@sz div 100"/>
            </xsl:attribute>
          </xsl:if>
          <!--yx,add,2010.4.28-->
          <xsl:attribute name="颜色_412F">
            <xsl:choose>

              <!--2014-3-10, update by Qihy, 修复BUG3075 集成测试ooxml-uof（20102013），部分艺术字转换后丢失以及字体颜色转换不正确（UOF不支持渐变，渐变效果丢失） start-->
              <xsl:when test ="a:solidFill/a:srgbClr/@val = 'FFFFFF' or a:ln/a:solidFill/a:srgbClr/@val = 'FFFFFF'">
                <xsl:value-of select="'#000000'"/>
              </xsl:when>
              <xsl:when test ="a:solidFill/a:srgbClr">
                <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
              </xsl:when>
              <xsl:when test ="a:ln/a:solidFill/a:srgbClr and a:ln/a:solidFill/a:srgbClr">
                <xsl:value-of select ="concat('#',a:ln/a:solidFill/a:srgbClr/@val)"/>
              </xsl:when>
              <xsl:when test ="a:solidFill/a:schemeClr or a:ln/a:solidFill/a:schemeClr or a:gradFill/a:gsLst/a:gs/a:schemeClr or ./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:schemeClr">
                <xsl:variable name ="id">
                  <xsl:choose>
                    <xsl:when test ="a:solidFill/a:schemeClr">
                      <xsl:value-of select ="a:solidFill/a:schemeClr/@val"/>
                    </xsl:when>
                    <xsl:when test ="a:ln/a:solidFill/a:schemeClr">
                      <xsl:value-of select ="a:ln/a:solidFill/a:schemeClr/@val"/>
                    </xsl:when>
                    <xsl:when test ="a:gradFill/a:gsLst/a:gs/a:schemeClr">
                      <xsl:value-of select ="a:gradFill/a:gsLst/a:gs/a:schemeClr/@val"/>
                    </xsl:when>
                    <xsl:when test ="./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:schemeClr">
                      <xsl:value-of select ="./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:schemeClr/@val"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="$id='tx1'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  <xsl:when test="$id='dk1'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  <xsl:when test="$id='lt1'">
                    <xsl:value-of select="'#FFFFFF'"/>
                  </xsl:when>
                  <xsl:when test="$id='dk2'">
                    <xsl:value-of select="'#1F497D'"/>
                  </xsl:when>
                  <xsl:when test="$id='lt2'">
                    <xsl:value-of select="'#EEECE1'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent1'">
                    <xsl:value-of select="'#4F81BD'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent2'">
                    <xsl:value-of select="'#C0504D'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent3'">
                    <xsl:value-of select="'#9BBB59'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent4'">
                    <xsl:value-of select="'#8064A2'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent5'">
                    <xsl:value-of select="'#4BACC6'"/>
                  </xsl:when>
                  <xsl:when test="$id='accent6'">
                    <xsl:value-of select="'#F79646'"/>
                  </xsl:when>
                  <xsl:when test="$id='hlink'">
                    <xsl:value-of select="'#0000FF'"/>
                  </xsl:when>
                  <xsl:when test="$id='folHlink'">
                    <xsl:value-of select="'#800080'"/>
                  </xsl:when>
                  
                  <!--2014-5-5, add by Qihy, bug3302,文本框中字体颜色不正确， start-->
                  <xsl:when test="$id='bg1'">
                    <xsl:value-of select="'#FFFFFF'"/>
                  </xsl:when>
                  <!--2014-5-5 end-->
                  
                  <xsl:otherwise>
                    <xsl:value-of select="'auto'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:when test ="not(a:solidFill) and a:gradFill/a:gsLst/a:gs[position()=2]/a:srgbClr">
                <xsl:value-of select ="concat('#',a:gradFill/a:gsLst/a:gs[position()=2]/a:srgbClr/@val)"/>
              </xsl:when>
              <xsl:when test ="./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:srgbClr">
                <xsl:value-of select ="concat('#',preceding::xdr:style/a:fontRef/a:srgbClr/@val)"/>
              </xsl:when>
              <!--2014-3-10 end-->
              
              <!-- 203130318 add by xuzhenwei BUG_2758:功能测试 文件转换失败 start --><!--
              <xsl:when test="a:solidFill or a:ln/a:solidFill">
                <xsl:choose>
                  
                  <xsl:when test="a:ln/a:solidFill/a:srgbClr = 'FFFFFF'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  <xsl:when test="a:ln/a:solidFill/a:srgbClr">
                    <xsl:value-of select="concat('#',a:ln/a:solidFill/a:srgbClr/@val)"/>
                  </xsl:when>
                  --><!-- end --><!--
                  <xsl:when test="a:solidFill/a:srgbClr = 'FFFFFF'">
                    <xsl:value-of select="'#000000'"/>
                  </xsl:when>
                  
                  <xsl:when test="a:solidFill/a:srgbClr">
                    <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
                  </xsl:when>
                  <xsl:when test="a:solidFill/a:schemeClr">
                    <xsl:variable name="id">
                      <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="$id='dk1'">
                        <xsl:value-of select="'#000000'"/>
                      </xsl:when>
                      <xsl:when test="$id='lt1'">
                        <xsl:value-of select="'#FFFFFF'"/>
                      </xsl:when>
                      <xsl:when test="$id='dk2'">
                        <xsl:value-of select="'#1F497D'"/>
                      </xsl:when>
                      <xsl:when test="$id='lt2'">
                        <xsl:value-of select="'#EEECE1'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent1'">
                        <xsl:value-of select="'#4F81BD'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent2'">
                        <xsl:value-of select="'#C0504D'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent3'">
                        <xsl:value-of select="'#9BBB59'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent4'">
                        <xsl:value-of select="'#8064A2'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent5'">
                        <xsl:value-of select="'#4BACC6'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent6'">
                        <xsl:value-of select="'#F79646'"/>
                      </xsl:when>
                      <xsl:when test="$id='hlink'">
                        <xsl:value-of select="'#0000FF'"/>
                      </xsl:when>
                      <xsl:when test="$id='folHlink'">
                        <xsl:value-of select="'#800080'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'auto'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'#000000'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:choose>
                  <xsl:when test="./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:schemeClr">
                    <xsl:variable name="id">
                      <xsl:value-of select="./ancestor::xdr:txBody[1]/preceding-sibling::xdr:style/a:fontRef/a:schemeClr/@val"/>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="$id='dk1'">
                        <xsl:value-of select="'#000000'"/>
                      </xsl:when>
                      --><!--yx,add,20105.19--><!--
                      <xsl:when test="$id='tx1'">
                        <xsl:value-of select="'#000000'"/>
                      </xsl:when>
                      <xsl:when test="$id='lt1'">
                        <xsl:value-of select="'#FFFFFF'"/>
                      </xsl:when>
                      <xsl:when test="$id='dk2'">
                        <xsl:value-of select="'#1F497D'"/>
                      </xsl:when>
                      <xsl:when test="$id='lt2'">
                        <xsl:value-of select="'#EEECE1'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent1'">
                        <xsl:value-of select="'#4F81BD'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent2'">
                        <xsl:value-of select="'#C0504D'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent3'">
                        <xsl:value-of select="'#9BBB59'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent4'">
                        <xsl:value-of select="'#8064A2'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent5'">
                        <xsl:value-of select="'#4BACC6'"/>
                      </xsl:when>
                      <xsl:when test="$id='accent6'">
                        <xsl:value-of select="'#F79646'"/>
                      </xsl:when>
                      <xsl:when test="$id='hlink'">
                        <xsl:value-of select="'#0000FF'"/>
                      </xsl:when>
                      <xsl:when test="$id='folHlink'">
                        <xsl:value-of select="'#800080'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'auto'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="preceding::xdr:style/a:fontRef/a:srgbClr">
                    <xsl:value-of select="concat('#',preceding::xdr:style/a:fontRef/a:srgbClr/@val)"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:otherwise>-->
            </xsl:choose>
          </xsl:attribute>
        </字:字体_4128>
      </xsl:if>
      <xsl:if test="@b">
        <字:是否粗体_4130>
            <xsl:choose>
              <xsl:when test="(@b='on')or(@b='1')or(@b='true')">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
      <xsl:if test="@i">
        <字:是否斜体_4131>
            <xsl:choose>
              <xsl:when test="(@i='on')or(@i='1')or(@i='true')">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
        </字:是否斜体_4131>
      </xsl:if>
      <字:删除线_4135>
          <xsl:choose>
            <xsl:when test="@strike='sngStrike'">single</xsl:when>
            <xsl:when test="@strike='dblStrike'">double</xsl:when>
            <xsl:when test="not(@strike) or (@strike='noStrike')">none</xsl:when>
          </xsl:choose>
      </字:删除线_4135>
      <字:下划线_4136>
        <!--下划线类型，虚实 需修改  李杨2011-11-30-->
        <xsl:if test="@u">
          <xsl:choose>
            <xsl:when test="@u='dottedHeavy'">
              <xsl:attribute name="线型_4137">dotted-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dashHeavy'">
              <xsl:attribute name="线型_4137">dashed-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dashLong'">
              <xsl:attribute name="线型_4137">dash-long</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dashLongHeavy'">
              <xsl:attribute name="线型_4137">dash-long-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dotDash'">
              <xsl:attribute name="线型_4137">dot-dash</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dotDashHeavy'">
              <xsl:attribute name="线型_4137">dash-dot-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dotDotDash'">
              <xsl:attribute name="线型_4137">dot-dot-dash</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dotDotDashHeavy'">
              <xsl:attribute name="线型_4137">dash-dot-dot-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='wavyHeavy'">
              <xsl:attribute name="线型_4137">wave-heavy</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='wavyDbl'">
              <xsl:attribute name="线型_4137">wavy-double</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='dbl'">
              <xsl:attribute name="线型_4137">double</xsl:attribute>
            </xsl:when>
            <!--heavy无对应 转为single类型-->
            <xsl:when test="@u='sng'">
              <xsl:attribute name="线型_4137">single</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='heavy'">
              <xsl:attribute name="线型_4137">single</xsl:attribute>
            </xsl:when>
            <xsl:when test="@u='words'">
              <xsl:attribute name="线型_4137">single</xsl:attribute>
              <xsl:attribute name="是否字下划线_4139">true</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="线型_4137">
                <xsl:value-of select="@u"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <!--oox中 表示类型的@u属性不一定出现，不出现表示类型为none-->
        <xsl:if test="not(@u)">
          <xsl:attribute name="线型_4137">none</xsl:attribute>
        </xsl:if>
        <!--4.18-->
        <xsl:if test="a:uFill/a:solidFill/a:srgbClr">
          <xsl:attribute name="颜色_412F">
            <xsl:value-of select="concat('#', a:uFill/a:solidFill/a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:if>
      </字:下划线_4136>
      <xsl:if test="@baseline">
        <字:上下标类型_4143>
            <xsl:if test="@baseline = '0'">
              <xsl:value-of select="'none'"/>
            </xsl:if>
            <xsl:if test="@baseline='30000'">
              <xsl:value-of select="'sup'"/>
            </xsl:if>
            <xsl:if test="@baseline='-25000'">
              <xsl:value-of select="'sub'"/>
            </xsl:if>
        </字:上下标类型_4143>
      </xsl:if>
      <字:字符间距_4145>
        <xsl:if test="@spc">
          <xsl:value-of select="@spc div 100"/>
        </xsl:if>
        <xsl:if test="not(@spc)">
          <xsl:value-of select="'0'"/>
        </xsl:if>
      </字:字符间距_4145>
      <字:调整字间距_4146>
        <xsl:if test="@kern">
          <xsl:value-of select="format-number(@kern div 100,'0.0')"/>
        </xsl:if>
        <xsl:if test="not(@kern)">
          <xsl:value-of select="'0'"/>
        </xsl:if>
      </字:调整字间距_4146>
    </字:句属性_4158>
  </xsl:template>
  <xsl:template name="pPr-body">
    <xsl:if test="@algn|@fontAlgn">
      <字:对齐_417D>
        <xsl:if test="@algn">
          <xsl:attribute name="水平对齐_421D">
            <xsl:choose>
              <xsl:when test="@algn='l'">left</xsl:when>
              <xsl:when test="@algn='r'">right</xsl:when>
              <xsl:when test="@algn='ctr'">center</xsl:when>
              <xsl:when test="@algn='just' or @algn='justLow'">justified</xsl:when>
              <xsl:when test="@algn='dist'or @algn='thaiDist'">distributed</xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@fontAlgn">
          <xsl:attribute name="文字对齐_421E">
            <xsl:choose>
              <xsl:when test="@fontAlgn='auto'">auto</xsl:when>
              <xsl:when test="@fontAlgn='t'">top</xsl:when>
              <xsl:when test="@fontAlgn='ctr'">center</xsl:when>
              <xsl:when test="@fontAlgn='base'">base</xsl:when>
              <xsl:when test="@fontAlgn='b'">bottom</xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
      </字:对齐_417D>
    </xsl:if>
    <!--缩进-->
    <xsl:if test="@indent|@marL|@marR">
      <字:缩进_411D>
        <xsl:choose>
          <!--hwj4.18-->
          <xsl:when test="@marL">
            <字:左_410E>
              <字:绝对_4107>
                <xsl:attribute name="值_410F">
                  <xsl:value-of select="@marL div 12700"/>
                </xsl:attribute>
              </字:绝对_4107>
            </字:左_410E>
          </xsl:when>
          <xsl:when test="@marR">
            <字:右_4110>
              <字:绝对_4107>
                <xsl:attribute name="值_410F">
                  <xsl:value-of select="@marR div 12700"/>
                </xsl:attribute>
              </字:绝对_4107>
            </字:右_4110>
          </xsl:when>
          <xsl:when test="@indent">
            <字:首行_4111>
              <字:绝对_4107>
                <xsl:attribute name="值_410F">
                  <xsl:value-of select="@indent div 12700"/>
                </xsl:attribute>
              </字:绝对_4107>
            </字:首行_4111>
          </xsl:when>
        </xsl:choose>
      </字:缩进_411D>
    </xsl:if>
    <!--行距-->
    <xsl:if test="a:lnSpc">
      <字:行距_417E>
        <xsl:if test="a:lnSpc/a:spcPct">
          <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
          <xsl:attribute name="值_4108">
            <xsl:value-of select="a:lnSpc/a:spcPct/@val div 100000"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="a:lnSpc/a:spcPts">
          <xsl:attribute name="类型_417F">fixed</xsl:attribute>
          <xsl:attribute name="值_4108">
            <xsl:value-of select="a:lnSpc/a:spcPts/@val div 100"/>
          </xsl:attribute>
        </xsl:if>
        <!--最小值、行间距 无对应-->
      </字:行距_417E>
    </xsl:if>
    <!--段间距-->
    <xsl:if test="a:spcBef|a:spcAft">
      <字:段间距_4180>
        <xsl:if test="a:spcBef">
          <字:段前距_4181>
            <xsl:if test="a:spcBef/a:spcPct">
              <字:相对值_4184>
                  <xsl:value-of select="a:spcBef/a:spcPct/@val div 100000"/>
              </字:相对值_4184>
            </xsl:if>
            <xsl:if test="a:spcBef/a:spcPts">
              <字:绝对值_4183>
                  <xsl:value-of select="a:spcBef/a:spcPts/@val div 100"/>
              </字:绝对值_4183>
            </xsl:if>
            <!--自动 无对应-->
          </字:段前距_4181>
        </xsl:if>
        <xsl:if test="a:spcAft">
          <字:段后距_4185>
            <!-- 20130315 update by xuzhenwei 修正smartArt断后距 start -->
            <xsl:if test="a:spcAft/a:spcPct">
              <!-- end -->
              <字:相对值_4184>
                  <xsl:value-of select="a:spcAft/a:spcPct/@val div 100000"/>
              </字:相对值_4184>
            </xsl:if>
            <xsl:if test="a:spcAft/a:spcPts">
              <字:绝对值_4183>
                  <xsl:value-of select="a:spcAft/a:spcPts/@val div 100"/>
              </字:绝对值_4183>
            </xsl:if>
            <!--自动 无对应-->
          </字:段后距_4185>
        </xsl:if>
      </字:段间距_4180>
    </xsl:if>
    <!--hwj-->
    <!--编号引用-->
    <xsl:if test="a:buChar|a:buAutoNum|a:buBlip">
      <字:自动编号信息_4186>
        <xsl:attribute name="编号引用_4187">
          <xsl:if test="@id">
            <xsl:value-of select="concat('bn_',@id)"/>
          </xsl:if>
          <!--p/pPr/bu...的情况 -->
          <xsl:if test="not(@id)">
            <xsl:value-of select="concat('bn_',translate(../@id,':',''))"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:attribute name="编号级别_4188">1</xsl:attribute>
      </字:自动编号信息_4186>
    </xsl:if>
    <xsl:if test="a:buNone">
      <字:自动编号信息_4186 编号引用_4187="bn_0" 编号级别_4188="0"/>
    </xsl:if>
    <!--允许单词断字，即西文在单词中间换行-->
    <xsl:if test="@eaLnBrk='0' or @eaLnBrk='false'">
      <字:是否允许单词断字_4194>false</字:是否允许单词断字_4194>
    </xsl:if>
    <xsl:if test="@eaLnBrk='1' or @eaLnBrk='true' or not(@eaLnBrk)">
      <字:是否允许单词断字_4194>true</字:是否允许单词断字_4194>
    </xsl:if>
    <!--行首尾标点控制-->
    <xsl:if test="@hangingPunct='0' or @hangingPunct='false' or not(@hangingPunct)">
      <!--标点在下一行显示==允许标点溢出边界为false-->
      <字:是否行首尾标点控制_4195>false</字:是否行首尾标点控制_4195>
    </xsl:if>
    <xsl:if test="@hangingPunct='1' or @hangingPunct='true'">
      <!--标点在本行显示==允许标点溢出边界-->
      <字:是否行首尾标点控制_4195>true</字:是否行首尾标点控制_4195>
    </xsl:if>
    <!--中文习惯首尾字符-->
    <xsl:if test="@latinLnBrk='1' or @latinLnBrk='true' or not(@latinLnBrk)">
      <字:是否采用中文习惯首尾字符_4197>true</字:是否采用中文习惯首尾字符_4197>
    </xsl:if>
    <xsl:if test="@latinLnBrk = '0' or @latinLnBrk='false'">
      <字:是否采用中文习惯首尾字符_4197>false</字:是否采用中文习惯首尾字符_4197>
    </xsl:if>
  </xsl:template>
  
  <!-- update xuzhenwei 20121228 解决新增功能点smartArt 将smartArt转成组合图形 -->
  <xsl:template match="xdr:sp|dsp:sp">
    <!--图形-->
    <xsl:param name="s_name"/>
    <xsl:param name="s_number"/>
    <xsl:param name="drawingId"/>
    <xsl:param name="twoCellAnchorPos"/>
    <!-- add by xuzhenwei BUG_2910 第二轮回归 oo-uof 功能 文本框字体不正确，该变量用来控制与style.xsl中<式样:字体声明_990D>计算字体名称的方式保持一致  -->
    <xsl:param name="posi"/>
    <!-- end -->
    <图:图形_8062>
      <xsl:variable name="ppid">
        <xsl:choose>
          <xsl:when test="xdr:nvSpPr">
            <xsl:value-of select="xdr:nvSpPr/xdr:cNvPr/@id"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat(substring-before($drawingId,'.xml'),$s_number,$s_number)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$ppid)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <xsl:choose>
          <xsl:when test="name(.)='dsp:sp'">
            <xsl:value-of select="'1'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:number count="xdr:twoCellAnchor" level="single"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <图:预定义图形_8018>
        <xsl:apply-templates select="xdr:spPr|dsp:spPr" mode="first"/>
        <图:属性_801D>
          <!-- 20130515 add by xuzhenwei 解决 Alt Text功能丢失 start -->
          <xsl:if test="xdr:nvSpPr/xdr:cNvPr/@descr">
            <图:Web文字_804F>
              <xsl:value-of select="xdr:nvSpPr/xdr:cNvPr/@descr"/>
            </图:Web文字_804F>
          </xsl:if>
          <!-- end -->
          <xsl:apply-templates select="xdr:spPr|dsp:spPr" mode="second"/>
        </图:属性_801D>
      </图:预定义图形_8018>

      <!--2014-3-4 Qihy 添加对象-控制点的转换实现 start -->
      <xsl:choose>
        <!--椭圆形标注，圆角矩形标注，矩形标注，云形标注-->
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='wedgeEllipseCallout' or ./xdr:spPr/a:prstGeom/@prst='wedgeRoundRectCallout' or ./xdr:spPr/a:prstGeom/@prst='wedgeRectCallout' or ./xdr:spPr/a:prstGeom/@prst='cloudCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='borderCallout1' or ./xdr:spPr/a:prstGeom/@prst='accentCallout1' or ./xdr:spPr/a:prstGeom/@prst='callout1' or ./xdr:spPr/a:prstGeom/@prst='accentBorderCallout1') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='borderCallout2' or ./xdr:spPr/a:prstGeom/@prst='accentCallout2' or ./xdr:spPr/a:prstGeom/@prst='callout2' or ./xdr:spPr/a:prstGeom/@prst='accentBorderCallout2') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='borderCallout3' or ./xdr:spPr/a:prstGeom/@prst='accentCallout3' or ./xdr:spPr/a:prstGeom/@prst='callout3' or ./xdr:spPr/a:prstGeom/@prst='accentBorderCallout3') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='star4' or ./xdr:spPr/a:prstGeom/@prst='star8' or ./xdr:spPr/a:prstGeom/@prst='star16' or ./xdr:spPr/a:prstGeom/@prst='star24' or ./xdr:spPr/a:prstGeom/@prst='star32' or ./xdr:spPr/a:prstGeom/@prst='star6' or ./xdr:spPr/a:prstGeom/@prst='star7' or ./xdr:spPr/a:prstGeom/@prst='star10' or ./xdr:spPr/a:prstGeom/@prst='star12') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='ribbon2') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='ribbon') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='ellipseRibbon2') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='ellipseRibbon') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='verticalScroll' or ./xdr:spPr/a:prstGeom/@prst='horizontalScroll') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='wave' or ./xdr:spPr/a:prstGeom/@prst='doubleWave') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='rightArrow' or ./xdr:spPr/a:prstGeom/@prst='stripedRightArrow' or ./xdr:spPr/a:prstGeom/@prst='notchedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='leftArrow' or ./xdr:spPr/a:prstGeom/@prst='leftRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='upArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='upDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='downArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='quadArrow' or ./xdr:spPr/a:prstGeom/@prst='leftRightUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='bentArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='leftUpArrow' or ./xdr:spPr/a:prstGeom/@prst='bentUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='curvedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='curvedLeftArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='curvedUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='curvedDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test ="(./xdr:spPr/a:prstGeom/@prst='homePlate' or ./xdr:spPr/a:prstGeom/@prst='chevron') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - $x * 21600 * $height div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--右箭头标注-->
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='rightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='leftArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='upArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='downArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='leftRightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='quadArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='triangle' or ./xdr:spPr/a:prstGeom/@prst='octagon' or ./xdr:spPr/a:prstGeom/@prst='plus' or ./xdr:spPr/a:prstGeom/@prst='plaque' or ./xdr:spPr/a:prstGeom/@prst='cube' or ./xdr:spPr/a:prstGeom/@prst='bevel' or ./xdr:spPr/a:prstGeom/@prst='sun' or ./xdr:spPr/a:prstGeom/@prst='moon' or ./xdr:spPr/a:prstGeom/@prst='bracketPair' or ./xdr:spPr/a:prstGeom/@prst='bracePair') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='parallelogram' or ./xdr:spPr/a:prstGeom/@prst='trapezoid' or ./xdr:spPr/a:prstGeom/@prst='hexagon' or ./xdr:spPr/a:prstGeom/@prst='donut' or ./xdr:spPr/a:prstGeom/@prst='noSmoking') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x * 21600 * $height div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--圆柱形-->
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='can') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='foldedCorner') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='smileyFace') and ./*/a:prstGeom/a:avLst/a:gd">
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='leftBracket' or ./xdr:spPr/a:prstGeom/@prst='rightBracket') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$y * 21600 * $width div $height div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左大括号，右大括号-->
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='leftBrace' or ./xdr:spPr/a:prstGeom/@prst='rightBrace') and ./*/a:prstGeom/a:avLst/a:gd">
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
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./xdr:spPr/a:xfrm/a:ext/@cy"/>
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
        <xsl:when test="(./xdr:spPr/a:prstGeom/@prst='bentConnector3' or ./xdr:spPr/a:prstGeom/@prst='curvedConnector3') and ./*/a:prstGeom/a:avLst/a:gd">
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
      </xsl:choose>
      <!--2014-3-4 Qihy 添加对象-控制点的转换实现 end -->
      
      <xsl:if test="xdr:sp/xdr:spPr/a:xfrm[@rot]|dsp:sp/dsp:spPr/a:xfrm[@rot]">
        <图:翻转_803A><!--uof:attrList="方向" 图:方向="x">-->
          x
          <!--<xsl:variable name="rot" select="xdr:sp/xdr:spPr/a:xfrm/@rot"/>
          <xsl:value-of select="$rot div 60000"/>-->
        </图:翻转_803A>
      </xsl:if>
      <!--组合位置-->
      <xsl:if test="parent::xdr:grpSp">
        <图:组合位置_803B>
          <xsl:call-template name="groupLocation"/>
        </图:组合位置_803B>
      </xsl:if>
      <!-- start xuzhenwei 新增smartArt功能点：控制组合图形位置 2013-01-08 -->
      <xsl:if test="ancestor::dsp:drawing">
        <图:组合位置_803B>
          <xsl:call-template name="groupLocationSmart">
            <xsl:with-param name="twoCellAnchorPosition" select="$twoCellAnchorPos"/>
          </xsl:call-template>
        </图:组合位置_803B>
      </xsl:if>
      <!-- end -->
      <!-- 20130315 update by xuzhenwei 考虑smartArt中内容为空的时候，不能生成文本内容 start -->
      <xsl:if test="xdr:txBody/a:bodyPr|dsp:txBody/a:bodyPr">
        <图:文本_803C>
          <xsl:apply-templates select="xdr:txBody/a:bodyPr|dsp:txBody/a:bodyPr"/>
          <图:内容_8043>
            <xsl:apply-templates select="xdr:txBody/a:p|dsp:txBody/a:p">
              <xsl:with-param name="positi" select="$posi"/>
            </xsl:apply-templates>
          </图:内容_8043>
        </图:文本_803C>
      </xsl:if>
      <!-- end -->
    </图:图形_8062>
  </xsl:template>
  <xsl:template match="xdr:cxnSp">
    <xsl:param name="s_name"/>
    <图:图形_8062>
      <xsl:variable name="ppid">
        <xsl:if test="xdr:nvSpPr">
          <xsl:value-of select="xdr:nvSpPr/xdr:cNvPr/@id"/>
        </xsl:if>
        <xsl:if test="xdr:nvCxnSpPr">
          <xsl:value-of select="xdr:nvCxnSpPr/xdr:cNvPr/@id"/>
        </xsl:if>
      </xsl:variable>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$ppid)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <xsl:number count="xdr:twoCellAnchor" level="single"/>
      </xsl:attribute>
      <图:预定义图形_8018>
        <xsl:apply-templates select="xdr:spPr" mode="first"/>
        <图:属性_801D>
          <xsl:apply-templates select="xdr:spPr" mode="second"/>
        </图:属性_801D>
      </图:预定义图形_8018>
      
      
      
      
      <xsl:if test="xdr:cxnSp/xdr:spPr/a:xfrm[@rot]">
        <图:翻转_803A>x</图:翻转_803A>
        <!--uof:attrList="方向" 图:方向="x">
          <xsl:variable name="rot" select="xdr:sp/xdr:spPr/a:xfrm/@rot"/>
          <xsl:value-of select="$rot div 60000"/>-->
      </xsl:if>
      <xsl:if test="//xdr:grpSpPr">
        <图:组合位置_803B>
          <xsl:call-template name="groupLocation"/>
        </图:组合位置_803B>
      </xsl:if>
      <xsl:if test="xdr:sp/xdr:txBody">
        <图:文本_803C>
          <xsl:apply-templates select="xdr:txBody/a:bodyPr"/>
          <图:内容_8043>
            <xsl:apply-templates select="xdr:txBody/a:p"/>
          </图:内容_8043>
        </图:文本_803C>
      </xsl:if>
    </图:图形_8062>
  </xsl:template>

  <xsl:template match="xdr:grpSp">
    <xsl:param name="s_name"/>
    <xsl:for-each select="xdr:cxnSp|xdr:grpSp|xdr:pic|xdr:sp">
      <xsl:apply-templates select=".">
        <xsl:with-param name="s_name" select="$s_name"/>
      </xsl:apply-templates>
    </xsl:for-each>
    <图:图形_8062>
      <xsl:variable name="agrid" select="xdr:nvGrpSpPr/xdr:cNvPr/@id"/>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$agrid)"/>
      </xsl:attribute>
      <xsl:variable name="agr">
        <xsl:number count="xdr:grpSp" level="single"/>
      </xsl:variable>
      <xsl:attribute name="层次_8063">
        <xsl:value-of select="$agr - 1"/>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <!-- update by xuzhenwei 调整一下输出内容组合列表的顺序，以便于uof到OO方向进行处理 start -->
        <xsl:for-each select="xdr:pic[(preceding::xdr:grpSp)]">
          <xsl:variable name="id">
            <xsl:value-of select="xdr:nvPicPr/xdr:cNvPr/@id"/>
          </xsl:variable>
          <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
        </xsl:for-each>
        <xsl:for-each select="xdr:pic[not(preceding::xdr:grpSp)]">
          <xsl:variable name="id">
            <xsl:value-of select="xdr:nvPicPr/xdr:cNvPr/@id"/>
          </xsl:variable>
          <xsl:if test="position()!=last()">
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
          </xsl:if>
          <xsl:if test="position()=last()">
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
          </xsl:if>
        </xsl:for-each>
        <xsl:for-each select="xdr:grpSp|xdr:sp|xdr:cxnSp">
          <xsl:variable name="id">
            <xsl:choose>
              <xsl:when test="xdr:nvSpPr">
                <xsl:value-of select="xdr:nvSpPr/xdr:cNvPr/@id"/>
              </xsl:when>
              <xsl:when test="xdr:nvCxnSpPr">
                <xsl:value-of select="xdr:nvCxnSpPr/xdr:cNvPr/@id"/>
              </xsl:when>
              <xsl:when test="xdr:nvGrpSpPr">
                <xsl:value-of select="xdr:nvGrpSpPr/xdr:cNvPr/@id"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:if test="position()!=last()">
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
          </xsl:if>
          <xsl:if test="position()=last()">
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
          </xsl:if>
        </xsl:for-each>
        <!-- end -->
        <xsl:if test="xdr:grpSp/xdr:pic">
          <xsl:for-each select="xdr:grpSp/xdr:pic">
            <xsl:variable name="id">
              <xsl:value-of select="xdr:nvPicPr/xdr:cNvPr/@id"/>
            </xsl:variable>
            <xsl:if test="position()!=last()">
              <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:attribute>
      <图:预定义图形_8018>
        <xsl:apply-templates select="xdr:grpSpPr" mode="first"/>
        <图:属性_801D>
          <xsl:apply-templates select="xdr:grpSpPr" mode="second"/>
        </图:属性_801D>
      </图:预定义图形_8018>
      <!--组合位置-->
      <xsl:if test="parent::xdr:twoCellAnchor">
        <图:组合位置_803B>
          <xsl:call-template name="groupLocation"/>
        </图:组合位置_803B>
      </xsl:if>
    </图:图形_8062>
  </xsl:template>
  <xsl:template name="groupLocation">

    <!--<图:组合位置_803B>-->
      <xsl:attribute name="x_C606">
        <xsl:value-of select=".//a:xfrm/a:off/@x div 12700"/>
      </xsl:attribute>
      <xsl:attribute name="y_C607">
        <xsl:value-of select=".//a:xfrm/a:off/@y div 12700"/>
      </xsl:attribute>
    <!--</图:组合位置_803B>-->

  </xsl:template>
  <!-- add by xuzhenwei 20121231 新增功能点smartArt start 20130329 再修正 -->
  <xsl:template name="groupLocationSmart">
    <xsl:param name="twoCellAnchorPosition"/>
    <!--<图:组合位置_803B>-->
    <xsl:variable name="fc">
      <xsl:value-of select="ancestor::ws:Drawings/xdr:wsDr/xdr:twoCellAnchor[position()=$twoCellAnchorPosition]/xdr:from/xdr:col"/>
    </xsl:variable>
    <xsl:variable name="fr">
      <xsl:value-of select="ancestor::ws:Drawings/xdr:wsDr/xdr:twoCellAnchor[position()=$twoCellAnchorPosition]/xdr:from/xdr:row"/>
    </xsl:variable>
    <xsl:variable name="fco">
      <xsl:value-of select="ancestor::ws:Drawings/xdr:wsDr/xdr:twoCellAnchor[position()=$twoCellAnchorPosition]/xdr:from/xdr:colOff"/>
    </xsl:variable>
    <xsl:variable name="fro">
      <xsl:value-of select="ancestor::ws:Drawings/xdr:wsDr/xdr:twoCellAnchor[position()=$twoCellAnchorPosition]/xdr:from/xdr:rowOff"/>
    </xsl:variable>
    <xsl:variable name="xDistance" select="floor($fc * 54 + $fco * 28.3 div 360000)"/>
    <xsl:variable name="yDistance" select ="floor($fr * 13.5 + $fro * 28.3 div 360000)"/>
    <xsl:attribute name="x_C606">
      <xsl:value-of select=".//a:xfrm/a:off/@x div 12700 + $xDistance"/>
    </xsl:attribute>
    <xsl:attribute name="y_C607">
      <xsl:value-of select=".//a:xfrm/a:off/@y div 12700 + $yDistance"/>
    </xsl:attribute>
    <!--</图:组合位置_803B>-->
  </xsl:template>
  <xsl:template  name="dspdrawing">
    <xsl:param name="s_name"/>
    <xsl:param name="dsmRelid"/>
    <!-- 根据dsmRelid参数，查找对应的drawingX.xml文件 -->
    <xsl:variable name="drawingId">
      <xsl:for-each select="parent::xdr:wsDr/pr:Relationships//pr:Relationship[@Type='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing']">
        <xsl:if test="@Id=$dsmRelid">
          <xsl:variable name="target" select="@Target"/>
          <xsl:value-of select="substring-after($target,'../diagrams/')"/>
        </xsl:if>
      </xsl:for-each>
    </xsl:variable>
    <!-- 2014-04-13 update by lingfeng 3076 集成测试ooxml->uof（2010/2013），智能标记转换后丢失 start -->
    <xsl:variable name="twoCellAnchor" select="substring-after($drawingId,'drawing')"/>
    <xsl:variable name="twoCellAnchorPos" select="substring-before($twoCellAnchor,'.xml')"/>
    <!-- 2014-04-13 update by lingfeng 3076 集成测试ooxml->uof（2010/2013），智能标记转换后丢失 end -->
    <xsl:for-each select="ancestor::xdr:wsDr/dsp:drawing[@id=$drawingId]//dsp:sp">
      <xsl:apply-templates select=".">
        <xsl:with-param name="s_name" select="$s_name"/>
        <xsl:with-param name="drawingId" select="$drawingId"/>
        <xsl:with-param name="s_number" select="position()"/>
        <xsl:with-param name="twoCellAnchorPos" select="$twoCellAnchorPos"/>
      </xsl:apply-templates>
    </xsl:for-each>
    <图:图形_8062>
      <xsl:variable name="drawingname" select="substring-before($drawingId,'.xml')"/>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$drawingname)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <xsl:value-of select="'0'"/>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <xsl:for-each select="parent::xdr:wsDr/dsp:drawing[@id=$drawingId]//dsp:sp">
          <xsl:variable name="s_number" select="position()"/>
          <xsl:value-of select="concat($s_name,'_OBJ0000',$drawingname,$s_number,$s_number,' ')"/>
         </xsl:for-each> 
      </xsl:attribute>
     <图:预定义图形_8018>
          <图:类别_8019>11</图:类别_8019>
          <图:名称_801A>Rectangle</图:名称_801A>
          <图:生成软件_801B>Microsoft Office</图:生成软件_801B>
          <图:属性_801D>
            <图:是否打印对象_804E>true</图:是否打印对象_804E>
            <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
            <图:大小_8060>
              <xsl:variable name="fc">
                <xsl:value-of select="./xdr:from/xdr:col"/>
              </xsl:variable>
              <xsl:variable name="fr">
                <xsl:value-of select="./xdr:from/xdr:row"/>
              </xsl:variable>
              <xsl:variable name="tc">
                <xsl:value-of select="./xdr:to/xdr:col"/>
              </xsl:variable>
              <xsl:variable name="tr">
                <xsl:value-of select="./xdr:to/xdr:row"/>
              </xsl:variable>
              <xsl:variable name="fco">
                <xsl:value-of select="./xdr:from/xdr:colOff"/>
              </xsl:variable>
              <xsl:variable name="fro">
                <xsl:value-of select="./xdr:from/xdr:rowOff"/>
              </xsl:variable>
              <xsl:variable name="tco">
                <xsl:value-of select="./xdr:to/xdr:colOff"/>
              </xsl:variable>
              <xsl:variable name="tro">
                <xsl:value-of select="./xdr:to/xdr:rowOff"/>
              </xsl:variable>
              <xsl:variable name="c">
                <xsl:value-of select="$tc - $fc"/>
              </xsl:variable>
              <xsl:variable name="r">
                <xsl:value-of select="$tr - $fr"/>
              </xsl:variable>
              <xsl:variable name="co">
                <xsl:value-of select="$tco - $fco"/>
              </xsl:variable>
              <xsl:variable name="ro">
                <xsl:value-of select="$tro - $fro"/>
              </xsl:variable>
              <xsl:variable name="wit">
                <xsl:value-of select="$co"/>
              </xsl:variable>
              <xsl:variable name="hit">
                <xsl:value-of select="$ro"/>
              </xsl:variable>
              <xsl:attribute name="长_C604">
                <xsl:choose>
                  <xsl:when test="contains($hit,'-')">
                    <xsl:value-of select="translate($hit,'-','')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$hit"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="宽_C605">
                <xsl:choose>
                  <xsl:when test="contains($wit,'-')">
                    <xsl:value-of select="translate($wit,'-','')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$wit"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </图:大小_8060>
          </图:属性_801D>
        </图:预定义图形_8018>
    </图:图形_8062>
  </xsl:template>
  <!--end -->
  
  <xsl:template match="xdr:grpSp" mode="single">
    <xsl:param name="s_name"/>
    <!--组合图形-->
    <图:图形_8062>
      <xsl:variable name="agrid" select="xdr:nvGrpSpPr/xdr:cNvPr/@id"/>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$agrid)"/>
      </xsl:attribute>
      <xsl:variable name="agr">
        <xsl:number count="xdr:grpSp" level="single"/>
      </xsl:variable>
      <xsl:attribute name="层次_8063">
        <xsl:value-of select="$agr - 1"/>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <xsl:for-each select="xdr:sp">
          <xsl:variable name="id">
            <xsl:if test="xdr:nvSpPr">
              <xsl:value-of select="xdr:nvSpPr/xdr:cNvPr/@id"/>
            </xsl:if>
            <xsl:if test="xdr:nvCxnSpPr">
              <xsl:value-of select="xdr:nvCxnSpPr/xdr:cNvPr/@id"/>
            </xsl:if>
          </xsl:variable>
          <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
        </xsl:for-each>
        <xsl:if test="xdr:pic">
          <xsl:for-each select="xdr:pic">
            <xsl:variable name="id">
              <xsl:value-of select="xdr:nvPicPr/xdr:cNvPr/@id"/>
            </xsl:variable>
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id,' ')"/>
          </xsl:for-each>
        </xsl:if>
        <xsl:if test="xdr:grpSp">
          <xsl:for-each select="xdr:grpSp">
            <xsl:variable name="id" select="xdr:nvGrpSpPr/xdr:cNvPr/@id"/>
            <xsl:value-of select="concat($s_name,'_OBJ0000',$id)"/>
          </xsl:for-each>
        </xsl:if>
      </xsl:attribute>
      <!--图:组合列表="Sheet1_OBJ1 Sheet_OBJ2 "-->
      <图:预定义图形_8018>
        <xsl:apply-templates select="xdr:grpSpPr" mode="first"/>
        <图:属性_801D>
          <xsl:apply-templates select="xdr:grpSpPr" mode="second"/>
        </图:属性_801D>
      </图:预定义图形_8018>
    </图:图形_8062>
    <!--新增加的12月24日夏艳霞-->
  </xsl:template>
  <xsl:template match="xdr:grpSpPr" mode="first">
    <图:类别_8019>
      <xsl:call-template name="ShapeClass"/>
    </图:类别_8019>
    <图:名称_801A>
      <xsl:call-template name="ShapeName"/>
    </图:名称_801A>
    <图:生成软件_801B>Microsoft Office</图:生成软件_801B>
  </xsl:template>
  <xsl:template match="xdr:grpSpPr" mode="second">
    <!--填充-->
    <xsl:if test="not(a:noFill)">
      <xsl:if test="a:solidFill/a:srgbClr">
        <图:填充_804C>
          <图:颜色_8004>
            <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
          </图:颜色_8004>
        </图:填充_804C>
      </xsl:if>
      <xsl:if test="a:solidFill/a:schemeClr">
        <xsl:variable name="id">
          <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
        </xsl:variable>
        <图:填充_804C>
          <图:颜色_8004>
            <xsl:choose>
              <xsl:when test="$id='dk1'">
                <xsl:value-of select="'#000000'"/>
              </xsl:when>
              <xsl:when test="$id='lt1'">
                <xsl:value-of select="'#FFFFFF'"/>
              </xsl:when>
              <xsl:when test="$id='dk2'">
                <xsl:value-of select="'#1F497D'"/>
              </xsl:when>
              <xsl:when test="$id='lt2'">
                <xsl:value-of select="'#EEECE1'"/>
              </xsl:when>
              <xsl:when test="$id='accent1'">
                <xsl:value-of select="'#4F81BD'"/>
              </xsl:when>
              <xsl:when test="$id='accent2'">
                <xsl:value-of select="'#C0504D'"/>
              </xsl:when>
              <xsl:when test="$id='accent3'">
                <xsl:value-of select="'#9BBB59'"/>
              </xsl:when>
              <xsl:when test="$id='accent4'">
                <xsl:value-of select="'#8064A2'"/>
              </xsl:when>
              <xsl:when test="$id='accent5'">
                <xsl:value-of select="'#4BACC6'"/>
              </xsl:when>
              <xsl:when test="$id='accent6'">
                <xsl:value-of select="'#F79646'"/>
              </xsl:when>
              <xsl:when test="$id='hlink'">
                <xsl:value-of select="'#0000FF'"/>
              </xsl:when>
              <xsl:when test="$id='folHlink'">
                <xsl:value-of select="'#800080'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'auto'"/>
              </xsl:otherwise>
            </xsl:choose>
          </图:颜色_8004>
        </图:填充_804C>
      </xsl:if>
      <xsl:if test="a:gradFill">
        <图:填充_804C>
          <xsl:apply-templates select="a:gradFill"/>
        </图:填充_804C>
      </xsl:if>

      <!--<xsl:if test="not(a:solidFill) and not(a:gradFill)">
        <图:填充 uof:locID="g0012">
          <图:颜色_8004>auto</图:颜色_8004>
        </图:填充>
      </xsl:if>-->
    </xsl:if>
    <xsl:if test="a:xfrm/@rot">
      <!--4.18-->
      <图:旋转角度_804D>
        <xsl:value-of select="360 - (21600000-a:xfrm/@rot) div 60000"/>
      </图:旋转角度_804D>
    </xsl:if>
    <xsl:if test="a:scene3d/a:camera/a:rot">
      <图:旋转角度_804D>
        <xsl:value-of select="(21600000-a:scene3d/a:camera/a:rot/@rev) div 60000"/>
      </图:旋转角度_804D>
    </xsl:if>
    <图:是否打印对象_804E>true</图:是否打印对象_804E>
    
    <xsl:apply-templates select="a:ln"/>
    <xsl:if test="ancestor::p:sp/p:nvSpPr/p:cNvPr/@descr">
      <图:Web文字_804F>
        <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:cNvPr/@descr"/>
      </图:Web文字_804F>
    </xsl:if>
    <xsl:if test="ancestor::p:cxnSp/p:nvCxnSpPr/p:cNvPr/@descr">
      <图:Web文字_804F>
        <xsl:value-of select="ancestor::p:cxnSp/p:nvCxnSpPr/p:cNvPr/@descr"/>
      </图:Web文字_804F>
    </xsl:if>
    <xsl:if test="a:solidFill/*/a:alpha">
      <图:透明度_8050>
        <xsl:variable name="tm">

          <!--2014-3-13, update by Qihy, 修复透明度取值为NaN， start-->
          <!--<xsl:value-of select="substring(a:solidFill/*/a:alpha/@val,1,string-length(a:solidFill/*/a:alpha/@val)-3)"/>
        </xsl:variable>
        <xsl:value-of select="100 - $tm"/>-->
          <xsl:value-of select="a:solidFill/*/a:alpha/@val"/>
        </xsl:variable>
        <xsl:value-of select="number(100000-$tm) div 1000"/>
        <!--2014-3-13 end-->
        
      </图:透明度_8050>
    </xsl:if>
    <!--阴影  李杨2011-12-09-->
    <xsl:if test="a:effectLst/a:outerShdw">
      <!--外部阴影-->
      <xsl:apply-templates select="a:outerShdw" mode="shadow"/>
    </xsl:if>
    <xsl:if test="a:effectLst/a:innerShdw">
      <!--内部阴影-->
      <xsl:apply-templates select="a:innerShdw" mode="shadow"/>
    </xsl:if>
    <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
    <xsl:if test="a:xfrm/a:ext">
      <图:大小_8060>
        <xsl:attribute name ="长_C604">
          <xsl:value-of select="a:xfrm/a:ext/@cy div 12700"/>
        </xsl:attribute>
        <xsl:attribute name ="宽_C605">
          <xsl:value-of select="a:xfrm/a:ext/@cx div 12700"/>
        </xsl:attribute>
      </图:大小_8060>
    </xsl:if>
    <xsl:if test="a:scene3d">
      <图:三维效果_8061>
        <xsl:apply-templates select="a:scene3d"/>
      </图:三维效果_8061>
    </xsl:if>
    
  </xsl:template>
 

  <!--添加“xdr:graphicFrame”模板，用于图表  李杨2011-12-26-->
  <xsl:template match="xdr:graphicFrame">
    <!-- 20130426 update by xuzhenwei 计算实际的图表大小 start -->
    <!-- 图表大小的实际宽度 -->
    <xsl:param name="grapWidth"/>
    <!--图形-->
    <xsl:param name="s_name"/>
    <图:图形_8062>
      <xsl:variable name="ppid">
        <xsl:if test="xdr:nvGraphicFramePr/xdr:cNvPr">
          <xsl:value-of select="xdr:nvGraphicFramePr/xdr:cNvPr/@id"/>
        </xsl:if>
      </xsl:variable>
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="concat($s_name,'_OBJ0000',$ppid)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <xsl:number count="xdr:twoCellAnchor" level="single"/>
      </xsl:attribute>
      <图:预定义图形_8018>
        <图:类别_8019>11</图:类别_8019>
        <图:名称_801A>Rectangle</图:名称_801A>
        <图:生成软件_801B>Yozo Office</图:生成软件_801B>
        <图:属性_801D>
          <xsl:if test="a:xfrm/@rot">
            <图:旋转角度_804D>
              <xsl:value-of select="360-(21600000-a:xfrm/@rot) div 60000"/>
            </图:旋转角度_804D>
          </xsl:if>
          <xsl:if test ="not(a:xfrm/@rot)">
            <图:旋转角度_804D>0.0</图:旋转角度_804D>
          </xsl:if>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
          <图:缩放是否锁定纵横比_8055>false</图:缩放是否锁定纵横比_8055>
          <图:线_8057>
            <图:线类型_8059 线型_805A="single" 虚实_805B="solid"/>
            <图:线粗细_805C>0.75</图:线粗细_805C>
          </图:线_8057>
            <!-- 20140415 update by lingfeng 调整表格的长度 start -->
          <图:大小_8060>
              <xsl:variable name="fr">
                  <xsl:value-of select="../xdr:from/xdr:row"/>
              </xsl:variable>
              <xsl:variable name="tr">
                  <xsl:value-of select="../xdr:to/xdr:row"/>
              </xsl:variable>
              <xsl:variable name="fro">
                  <xsl:value-of select="../xdr:from/xdr:rowOff"/>
              </xsl:variable>
              <xsl:variable name="tro">
                  <xsl:value-of select="../xdr:to/xdr:rowOff"/>
              </xsl:variable>
              <xsl:variable name="r">
                  <xsl:value-of select="$tr - $fr"/>
              </xsl:variable>
              <xsl:variable name="ro">
                  <xsl:value-of select="$tro - $fro"/>
              </xsl:variable>
              <xsl:variable name="hit">
                  <xsl:value-of select="floor($r * 13.5 + $ro * 28.3 div 360000)"/>
              </xsl:variable>
              <xsl:attribute name="长_C604">
                  <xsl:choose>
                      <xsl:when test="contains($hit,'-')">
                          <xsl:value-of select="translate($hit,'-','')"/>
                      </xsl:when>
                      <xsl:otherwise>
                          <xsl:value-of select="$hit"/>
                      </xsl:otherwise>
                  </xsl:choose>
              </xsl:attribute>
            <xsl:attribute name="宽_C605">
              <xsl:value-of select="$grapWidth"/>
            </xsl:attribute>
              <!--end-->
          </图:大小_8060>
          <!-- end -->
        </图:属性_801D>
      </图:预定义图形_8018>
      
      <图:图表数据引用_8065>
        <xsl:variable name ="sheetNo">
          <xsl:value-of select ="ancestor::ws:spreadsheet/@sheetName"/>
        </xsl:variable>
        <xsl:variable name ="id">
          <xsl:value-of select ="a:graphic/a:graphicData/c:chart/@r:id"/>
        </xsl:variable>
        <xsl:value-of select ="concat($sheetNo,'_chart_',$id)"/>
      </图:图表数据引用_8065>
    </图:图形_8062>
  </xsl:template>

  <!--外部阴影  李杨2011-12-09-->
  <xsl:template match="a:outerShdw" mode="shadow">
    <图:阴影_8051 是否显示阴影_C61C="true" 类型_C61D="single">

      <xsl:choose>
        <xsl:when test="a:srgbClr">
          <xsl:attribute name="颜色_C61E">
            <xsl:if test ="a:srgbClr/@val">
              <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
            </xsl:if>
            <xsl:if test ="not(a:srgbClr/@val)">
              <xsl:value-of select ="'auto'"/>
            </xsl:if>
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
          <!--<xsl:value-of select="round((//a:alpha/@val div (1000 * 100))*256)"/>-->

          <xsl:when test="//a:alpha/@val &lt; 80000">
            <xsl:attribute name="透明度_C61F">50</xsl:attribute>
          </xsl:when>
          <xsl:when test="//a:alpha/@val &gt; 80000">
            <xsl:attribute name="透明度_C61F">100</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>
      <!-- update by xuzhenwei BUG_2473:转换后阴影效果显示不正确 2013-01-21 start -->
      <uof:偏移量_C61B>
        <xsl:choose>
          <xsl:when test="./@dir">
            <xsl:choose>
              <xsl:when test="./@dir &lt; '5400000'">
                <xsl:attribute name ="x_C606">
                  <xsl:value-of select ="'6.0'"/>
                </xsl:attribute>
                <xsl:attribute name ="y_C607">
                  <xsl:value-of select ="'6.0'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="./@dir &gt; '5400000'">
                <xsl:attribute name ="x_C606">
                  <xsl:value-of select ="'-6.0'"/>
                </xsl:attribute>
                <xsl:attribute name ="y_C607">
                  <xsl:value-of select ="'6.0'"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name ="x_C606">
                  <xsl:value-of select ="'2.0'"/>
                </xsl:attribute>
                <xsl:attribute name ="y_C607">
                  <xsl:value-of select ="'2.0'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="x">
              <xsl:attribute name="x_C606">
                <xsl:value-of select="round(x)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(x)">
              <xsl:attribute name ="x_C606">
                <xsl:value-of select ="'-6.0'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="y">
              <xsl:attribute name="y_C607">
                <xsl:value-of select="round(y)"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="not(y)">
              <xsl:attribute name ="y_C607">
                <xsl:value-of select ="'6.0'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </uof:偏移量_C61B>
      <!-- end -->
    </图:阴影_8051>
  </xsl:template>
  <!--内部阴影  李杨2011-12-09-->
  <xsl:template match="a:innerShdw" mode="shadow">
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
          <!--<xsl:value-of select="round((//a:alpha/@val div (1000 * 100))*256)"/>-->

          <xsl:when test="//a:alpha/@val &lt; 80000">
            <xsl:attribute name="透明度_C61F">50</xsl:attribute>
          </xsl:when>
          <xsl:when test="//a:alpha/@val &gt; 80000">
            <xsl:attribute name="透明度_C61F">100</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:if>

      <uof:偏移量_C61B>
        <xsl:if test="(@dir div 60000) = 0  or not(@dir)">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 45">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 90">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 135">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 180">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 225">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 270">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 315">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
        </xsl:if>
      </uof:偏移量_C61B>

    </图:阴影_8051>
  </xsl:template>

  <!--三维效果 李杨2011-12-01-->
  <xsl:template match="a:scene3d">
    <xsl:if test="not(../a:sp3d/@extrusionH)">
      <!--无深度-->
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
              
              <!--2014-6-8, 删除凌峰添加代码，其增加代码导致测试用例PredefinedShape_SmartArt.xlsx打开需要恢复， start-->
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
              <!--2014-6-8 end-->
              
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
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>324.6</uof:y方向_C637>
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
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>35.4</uof:y方向_C637>
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
                <uof:x方向_C636>306.5</uof:x方向_C636>
                <uof:y方向_C637>301.3</uof:y方向_C637>
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
                <uof:y方向_C637>309.6</uof:y方向_C637>
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
                <uof:x方向_C636>43.9</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
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
                <uof:x方向_C636>316.1</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
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
                <uof:x方向_C636>34.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
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
                <uof:x方向_C636>325.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
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
    <xsl:if test="../a:sp3d">
      <xsl:if test="../a:sp3d/a:extrusionClr/a:srgbClr">
        <uof:颜色_C63F>
          <xsl:variable name ="color">
            <xsl:value-of select="../a:sp3d/a:extrusionClr/a:srgbClr/@val"/>
          </xsl:variable>
          <xsl:value-of select ="concat('#',$color)"/>
        </uof:颜色_C63F>
      </xsl:if>
      <xsl:if test="../a:sp3d/@extrusionH">
        <uof:深度_C63B>
          <xsl:value-of select="../a:sp3d/@extrusionH div 12700"/>
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
    <xsl:if test="not(../a:sp3d)">
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
  </xsl:template>
</xsl:stylesheet>
