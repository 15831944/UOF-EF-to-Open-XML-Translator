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
        xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
        xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph">
  <!--六种颜色模式的子元素模板区-->
  <!--改变红色区-->
	<!--11.11.10 图案填充的图形引用属性 没有添加 李娟-->
  <xsl:template match="a:red">
    <xsl:param name="rgbColor"/>
    <xsl:variable name="red">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($red,substring($rgbColor,3,4))"/>
  </xsl:template>
  <xsl:template match="a:redMod">
    <xsl:param name="rgbColor"/>
    <!--取出原红色的十进制值-->
    <xsl:variable name="redVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,1,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="redMod">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * $redVal"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($redMod,substring($rgbColor,3,4))"/>
  </xsl:template>
  <xsl:template match="a:redOff">
    <xsl:param name="rgbColor"/>
    <!--取出原红色的十进制值-->
    <xsl:variable name="redVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,1,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="redOff">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$redVal + @val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($redOff,substring($rgbColor,3,4))"/>
  </xsl:template>
  <!--改变绿色区-->
  <xsl:template match="a:green">
    <xsl:param name="rgbColor"/>
    <xsl:variable name="green">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,2),$green,substring($rgbColor,5,2))"/>
  </xsl:template>
  <xsl:template match="a:greenMod">
    <xsl:param name="rgbColor"/>
    <!--取出原绿色的十进制值-->
    <xsl:variable name="greenVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,3,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="greenMod">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * $greenVal"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,2),$greenMod,substring($rgbColor,5,2))"/>
  </xsl:template>
  <xsl:template match="a:greenOff">
    <xsl:param name="rgbColor"/>
    <!--取出原绿色的十进制值-->
    <xsl:variable name="greenVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,3,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="greenOff">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$greenVal + @val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,2),$greenOff,substring($rgbColor,5,2))"/>
  </xsl:template>
  <!--改变蓝色区-->
  <xsl:template match="a:blue">
    <xsl:param name="rgbColor"/>
    <xsl:variable name="blue">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,4),$blue)"/>
  </xsl:template>
  <xsl:template match="a:blueMod">
    <xsl:param name="rgbColor"/>
    <!--取出原蓝色的十进制值-->
    <xsl:variable name="blueVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,5,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="blueMod">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="@val div 100000 * $blueVal"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,4),$blueMod)"/>
  </xsl:template>
  <xsl:template match="a:blueOff">
    <xsl:param name="rgbColor"/>
    <!--取出原蓝色的十进制值-->
    <xsl:variable name="blueVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,5,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算结果值-->
    <xsl:variable name="blueOff">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$blueVal + @val div 100000 * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat(substring($rgbColor,1,4),$blueOff)"/>
  </xsl:template>
  <!--shade模板-->
  <xsl:template match="a:shade">
    <xsl:param name="rgbColor"/>
    <xsl:variable name="shade">
      <xsl:value-of select="@val div 100000"/>
    </xsl:variable>
    <!--计算红色值-->
    <xsl:variable name="redOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,1,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="redRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$shade * $redOVal"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算绿色值-->
    <xsl:variable name="greenOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,3,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$shade * $greenOVal"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算蓝色值-->
    <xsl:variable name="blueOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,5,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$shade * $blueOVal"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($redRVal,$greenRVal,$blueRVal)"/>
  </xsl:template>
  <!--tint模板-->
  <xsl:template match="a:tint">
    <xsl:param name="rgbColor"/>
    <xsl:variable name="tint">
      <xsl:value-of select="@val div 100000"/>
    </xsl:variable>
    <!--计算红色值-->
    <xsl:variable name="redOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,1,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="redRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$tint * $redOVal + (1 - $tint) * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算绿色值-->
    <xsl:variable name="greenOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,3,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$tint * $greenOVal + (1 - $tint) * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <!--计算蓝色值-->
    <xsl:variable name="blueOVal">
      <xsl:call-template name="hexToDecimal">
        <xsl:with-param name="hex" select="substring($rgbColor,5,2)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueRVal">
      <xsl:call-template name="decimalToHex">
        <xsl:with-param name="result" select="$tint * $blueOVal + (1 - $tint) * 255"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($redRVal,$greenRVal,$blueRVal)"/>
  </xsl:template>
  <!--alpha区-->
  <xsl:template match="a:alpha">
    <xsl:param name="rgbColor"/>
    <!--把透明度值传到transparence模板中-->
    <xsl:call-template name="transparence">
      <xsl:with-param name="trans" select="100 - @val div 1000"/>
    </xsl:call-template>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--alphaMod区，暂不处理-->
  <xsl:template match="a:alphaMod">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--alphaOff区，暂不处理-->
  <xsl:template match="a:alphaOff">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--补足区，暂不处理-->
  <xsl:template match="a:comp">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--gamma区，暂不处理-->
  <xsl:template match="a:gamma">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--gray区，暂不处理-->
  <xsl:template match="a:gray">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--inv区，暂不处理-->
  <xsl:template match="a:inv">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--invGamma区，暂不处理-->
  <xsl:template match="a:invGamma">
    <xsl:param name="rgbColor"/>
    <xsl:value-of select="$rgbColor"/>
  </xsl:template>
  <!--保存透明度模板-->
  <xsl:template name="transparence">
    <xsl:param name="trans"/>
  </xsl:template>
  <!--颜色计算模板，统一调度所有颜色计算的子结点-->
  <xsl:template name="colorChild">
    <xsl:param name="recy">
      <xsl:value-of select="count(child::node())"/>
    </xsl:param>
    <xsl:param name="position" select="1"/>
    <xsl:param name="rgbColor"/>
    <xsl:variable name="currClr">
      <xsl:choose>
        <xsl:when test="$recy &gt; 0">
			
          <xsl:apply-templates select="descendant::*[$position]">
            <xsl:with-param name="rgbColor" select="$rgbColor"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
			
			<xsl:value-of select="$rgbColor"/>
		</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$recy  &gt; 0">
		 
        <xsl:call-template name="colorChild">
          <xsl:with-param name="rgbColor" select="$currClr"/>
          <xsl:with-param name="recy" select="$recy - 1"/>
          <xsl:with-param name="position" select="$position +1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
		  
		  <xsl:value-of select="concat('#',$currClr)"/>
	  </xsl:otherwise>
    </xsl:choose>
	  
  </xsl:template>
  <!--所有的填充模式-->
  <!--xsl:template name="colorChoice">
    <xsl:call-template name="colorChild">
      <xsl:with-param name="rgbColor">
        <xsl:call-template name="colorInitial"/>
      </xsl:with-param>
    </xsl:call-template>
	</xsl:template-->

  <xsl:template name="colorChoice">
    <xsl:param name="phClr"/>
    <!--liuyangyang phClr-->
    <xsl:choose>
      <xsl:when test="a:srgbClr">
        <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
        <!--<xsl:value-of select="@val"/>-->
        <!--<xsl:apply-templates select="a:srgbClr"/>-->
      </xsl:when>
      <xsl:when test="a:hslClr">
        <xsl:apply-templates select="a:hslClr"/>
      </xsl:when>
      <xsl:when test="a:prstClr">
        <xsl:apply-templates select="a:prstClr"/>
      </xsl:when>
      <xsl:when test="a:schemeClr">
        <xsl:apply-templates select="a:schemeClr">
          <xsl:with-param name="phClr" select="$phClr"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="a:scrgbClr">
        <xsl:apply-templates select="a:scrgbClr"/>
      </xsl:when>
      <xsl:when test="a:sysClr">
        <xsl:apply-templates select="a:sysClr"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--以下六种模式均转换成RGB模式-->
  <!--srgb模式-->
  <xsl:template match="a:srgbClr">
    <xsl:call-template name="colorChild">
      <xsl:with-param name="rgbColor" select="@val"/>
    </xsl:call-template>
  </xsl:template>
  <!--hsl模式-->
  <xsl:template match="a:hslClr">
    <xsl:variable name="orignalHue" select="@hue"/>
    <xsl:variable name="orignalSat" select="@sat"/>
    <xsl:variable name="orignalLum" select="@lum"/>
    <!--以下三个变量为需传递的参数值-->
    <xsl:variable name="passHue">
      <xsl:choose>
        <xsl:when test="a:hue">
          <xsl:value-of select="./a:hue/@val"/>
        </xsl:when>
        <xsl:when test="a:hueMod">
          <xsl:value-of select="./a:hueMod/@val div 100000 * $orignalHue"/>
        </xsl:when>
        <xsl:when test="a:hueOff">
          <xsl:value-of select="./a:hueOff/@val + $orignalHue"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$orignalHue"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="passSat">
      <xsl:choose>
        <xsl:when test="a:sat">
          <xsl:value-of select="./a:sat/@val"/>
        </xsl:when>
        <xsl:when test="a:satMod">
          <xsl:value-of select="./a:satMod/@val div 100000 * $orignalSat"/>
        </xsl:when>
        <xsl:when test="a:satOff">
          <xsl:value-of select="./a:satOff/@val + $orignalSat"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$orignalSat"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="passLum">
      <xsl:choose>
        <xsl:when test="a:lum">
          <xsl:value-of select="./a:lum/@val"/>
        </xsl:when>
        <xsl:when test="a:lumMod">
          <xsl:value-of select="./a:lumMod/@val div 100000 * $orignalLum"/>
        </xsl:when>
        <xsl:when test="a:lumOff">
          <xsl:value-of select="./a:lumOff/@val + $orignalLum"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$orignalLum"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!--调用转换模板-->
    <xsl:call-template name="HSLtoRGB">
      <xsl:with-param name="hue" select="$passHue"/>
      <xsl:with-param name="sat" select="$passSat"/>
      <xsl:with-param name="lum" select="$passLum"/>
    </xsl:call-template>
  </xsl:template>
  <!--scrgbClr模式-->
  <xsl:template match="a:scrgbClr"/>
  <!--sysClr模式-->
  <xsl:template match="a:sysClr">
    <xsl:value-of select="concat('#', ./@lastClr)"/>
  </xsl:template>
  <!--prstClr模式-->
  <xsl:template match="a:prstClr">
    <xsl:choose>
      <xsl:when test="@val='aliceBlue'">#f0f8ff</xsl:when>
      <xsl:when test="@val='antiqueWhite'">#faebd7</xsl:when>
      <xsl:when test="@val='aqua'">#00ffff</xsl:when>
      <xsl:when test="@val='aquamarine'">#7fffd4</xsl:when>
      <xsl:when test="@val='azure'">#f0ffff</xsl:when>
      <xsl:when test="@val='beige'">#f5f5dc</xsl:when>
      <xsl:when test="@val='bisque'">#ffe4c4</xsl:when>
      <xsl:when test="@val='black'">#000000</xsl:when>
      <xsl:when test="@val='blanchedAlmond'">#ffebcd</xsl:when>
      <xsl:when test="@val='blue'">#0000ff</xsl:when>
      <xsl:when test="@val='blueViolet'">#8a2be2</xsl:when>
      <xsl:when test="@val='brown'">#a52a2a</xsl:when>
      <xsl:when test="@val='burlyWood'">#deb887</xsl:when>
      <xsl:when test="@val='cadetBlue'">#5f9ea0</xsl:when>
      <xsl:when test="@val='chartreuse'">#7fff00</xsl:when>
      <xsl:when test="@val='chocolate'">#d2691e</xsl:when>
      <xsl:when test="@val='coral'">#ff7f50</xsl:when>
      <xsl:when test="@val='cornflowerBlue'">#6495ed</xsl:when>
      <xsl:when test="@val='cornsilk'">#fff8dc</xsl:when>
      <xsl:when test="@val='crimson'">#dc143c</xsl:when>
      <xsl:when test="@val='cyan'">#00ffff</xsl:when>
      <xsl:when test="@val='deepPink'">#ff1493</xsl:when>
      <xsl:when test="@val='deepSkyBlue'">#00bfff</xsl:when>
      <xsl:when test="@val='dimGray'">#696969</xsl:when>
      <xsl:when test="@val='dkBlue'">#00008b</xsl:when>
      <xsl:when test="@val='dkCyan'">#008b8b</xsl:when>
      <xsl:when test="@val='dkGoldenrod'">#b8860b</xsl:when>
      <xsl:when test="@val='dkGray'">#a9a9a9</xsl:when>
      <xsl:when test="@val='dkGreen'">#006400</xsl:when>
      <xsl:when test="@val='dkKhaki'">#bdb767</xsl:when>
      <xsl:when test="@val='dkMagenta'">#8b008b</xsl:when>
      <xsl:when test="@val='dkOliveGreen'">#556b2f</xsl:when>
      <xsl:when test="@val='dkOrange'">#ff8c00</xsl:when>
      <xsl:when test="@val='dkOrchild'">#9932cc</xsl:when>
      <xsl:when test="@val='dkRed'">#8b0000</xsl:when>
      <xsl:when test="@val='dkSalmon'">#e9967a</xsl:when>
      <xsl:when test="@val='dkSeaGreen'">#8fbc8b</xsl:when>
      <xsl:when test="@val='dkSlateBlue'">#483d8b</xsl:when>
      <xsl:when test="@val='dkSlateGray'">#2f4f4f</xsl:when>
      <xsl:when test="@val='dkTurquoise'">#00cece</xsl:when>
      <xsl:when test="@val='dkViolet'">#9400d3</xsl:when>
      <xsl:when test="@val='dodgerBlue'">#1e90ff</xsl:when>
      <xsl:when test="@val='firebrick'">#b22222</xsl:when>
      <xsl:when test="@val='floralWhite'">#fffaf0</xsl:when>
      <xsl:when test="@val='forestGreen'">#228b22</xsl:when>
      <xsl:when test="@val='fuchsia'">#ff00ff</xsl:when>
      <xsl:when test="@val='gainsboro'">#dcdcdc</xsl:when>
      <xsl:when test="@val='ghostWhite'">#f8f8ff</xsl:when>
      <xsl:when test="@val='gold'">#ffd700</xsl:when>
      <xsl:when test="@val='goldenrod'">#daa520</xsl:when>
      <xsl:when test="@val='gray'">#808080</xsl:when>
      <xsl:when test="@val='green'">#008000</xsl:when>
      <xsl:when test="@val='greenYellow'">#adff2f</xsl:when>
      <xsl:when test="@val='honeydew'">#f0fff0</xsl:when>
      <xsl:when test="@val='hotPink'">#ff69b4</xsl:when>
      <xsl:when test="@val='indianRed'">#cd5c5c</xsl:when>
      <xsl:when test="@val='indigo'">#4b0082</xsl:when>
      <xsl:when test="@val='ivory'">#fffff0</xsl:when>
      <xsl:when test="@val='khaki'">#f0e68c</xsl:when>
      <xsl:when test="@val='lavender'">#e6e6fa</xsl:when>
      <xsl:when test="@val='lavenderBlush'">#fff0f5</xsl:when>
      <xsl:when test="@val='lawnGreen'">#7cfc00</xsl:when>
      <xsl:when test="@val='lemonChiffon'">#fffacd</xsl:when>
      <xsl:when test="@val='lime'">#00ff00</xsl:when>
      <xsl:when test="@val='limeGreen'">#32cd32</xsl:when>
      <xsl:when test="@val='linen'">#faf0e6</xsl:when>
      <xsl:when test="@val='ltBlue'">#add8e6</xsl:when>
      <xsl:when test="@val='ltCoral'">#f08080</xsl:when>
      <xsl:when test="@val='ltCyan'">#e0ffff</xsl:when>
      <xsl:when test="@val='ltGoldenrodYellow'">#fafa78</xsl:when>
      <xsl:when test="@val='ltGray'">#d3d3d3</xsl:when>
      <xsl:when test="@val='ltGreen'">#90ee90</xsl:when>
      <xsl:when test="@val='ltPink'">#ffb6c1</xsl:when>
      <xsl:when test="@val='ltSalmon'">#ffa07a</xsl:when>
      <xsl:when test="@val='ltSeaGreen'">#20b2aa</xsl:when>
      <xsl:when test="@val='ltSkyBlue'">#87cefa</xsl:when>
      <xsl:when test="@val='ltSlateGray'">#778899</xsl:when>
      <xsl:when test="@val='ltSteelBlue'">#b0c4de</xsl:when>
      <xsl:when test="@val='ltYellow'">#ffffe0</xsl:when>
      <xsl:when test="@val='magenta'">#ff00ff</xsl:when>
      <xsl:when test="@val='maroon'">#800000</xsl:when>
      <xsl:when test="@val='medAquamarine'">#66cdaa</xsl:when>
      <xsl:when test="@val='medBlue'">#0000cd</xsl:when>
      <xsl:when test="@val='medOrchid'">#ba55d3</xsl:when>
      <xsl:when test="@val='medPurple'">#9370db</xsl:when>
      <xsl:when test="@val='medSeaGreen'">#3cb371</xsl:when>
      <xsl:when test="@val='medSlateBlue'">#7b68ee</xsl:when>
      <xsl:when test="@val='medSpringGreen'">#00fa9a</xsl:when>
      <xsl:when test="@val='medTurquoise'">#48d1cc</xsl:when>
      <xsl:when test="@val='medVioletRed'">#c71585</xsl:when>
      <xsl:when test="@val='midnightBlue'">#191970</xsl:when>
      <xsl:when test="@val='mintCream'">#f5fffa</xsl:when>
      <xsl:when test="@val='mistyRose'">#ffe4e1</xsl:when>
      <xsl:when test="@val='moccasin'">#ffe4b5</xsl:when>
      <xsl:when test="@val='navajoWhite'">#ffdead</xsl:when>
      <xsl:when test="@val='navy'">#000080</xsl:when>
      <xsl:when test="@val='oldLace'">#fdf5e6</xsl:when>
      <xsl:when test="@val='olive'">#808000</xsl:when>
      <xsl:when test="@val='oliveDrab'">#6b8e23</xsl:when>
      <xsl:when test="@val='orange'">#ffa500</xsl:when>
      <xsl:when test="@val='orangeRed'">#ff4500</xsl:when>
      <xsl:when test="@val='orchid'">#da70d6</xsl:when>
      <xsl:when test="@val='paleGoldenrod'">#eee8aa</xsl:when>
      <xsl:when test="@val='paleGreen'">#98fb98</xsl:when>
      <xsl:when test="@val='paleTurquoise'">#afeeee</xsl:when>
      <xsl:when test="@val='paleVioletRed'">#db7093</xsl:when>
      <xsl:when test="@val='papayaWhip'">#ffefd5</xsl:when>
      <xsl:when test="@val='peachPuff'">#ffdab9</xsl:when>
      <xsl:when test="@val='peru'">#cd853f</xsl:when>
      <xsl:when test="@val='pink'">#ffc0cb</xsl:when>
      <xsl:when test="@val='plum'">#dda0dd</xsl:when>
      <xsl:when test="@val='powderBlue'">#b0e0e6</xsl:when>
      <xsl:when test="@val='purple'">#800080</xsl:when>
      <xsl:when test="@val='red'">#ff0000</xsl:when>
      <xsl:when test="@val='rosyBrown'">#bc8f8f</xsl:when>
      <xsl:when test="@val='royalBlue'">#4169e1</xsl:when>
      <xsl:when test="@val='saddleBrown'">#8b4513</xsl:when>
      <xsl:when test="@val='salmon'">#fa8072</xsl:when>
      <xsl:when test="@val='sandyBrown'">#f4a460</xsl:when>
      <xsl:when test="@val='seaGreen'">#2e8b57</xsl:when>
      <xsl:when test="@val='seaShell'">#fff5ee</xsl:when>
      <xsl:when test="@val='sienna'">#a0522d</xsl:when>
      <xsl:when test="@val='silver'">#c0c0c0</xsl:when>
      <xsl:when test="@val='skyBlue'">#87ceeb</xsl:when>
      <xsl:when test="@val='slateBlue'">#6a5acd</xsl:when>
      <xsl:when test="@val='slateGray'">#708090</xsl:when>
      <xsl:when test="@val='snow'">#fffafa</xsl:when>
      <xsl:when test="@val='springGreen'">#00ff7f</xsl:when>
      <xsl:when test="@val='steelBlue'">#4682b4</xsl:when>
      <xsl:when test="@val='tan'">#d2b48c</xsl:when>
      <xsl:when test="@val='teal'">#008080</xsl:when>
      <xsl:when test="@val='thistle'">#d8bfd8</xsl:when>
      <xsl:when test="@val='tomato'">#ff6347</xsl:when>
      <xsl:when test="@val='turquoise'">#40e0d0</xsl:when>
      <xsl:when test="@val='violet'">#ee82ee</xsl:when>
      <xsl:when test="@val='wheat'">#f5deb3</xsl:when>
      <xsl:when test="@val='white'">#ffffff</xsl:when>
      <xsl:when test="@val='whiteSmoke'">#f5f5f5</xsl:when>
      <xsl:when test="@val='yellow'">#ffff00</xsl:when>
      <xsl:when test="@val='yellowGreen'">#9acd32</xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--2015-03-06,liuyangyang，修改主题色转换模版-->
  <!--2014-01-13, tangjiang,修改OOXML到UOF文字表式样的转换,文本框渐变填充 start -->
  <xsl:template match="a:schemeClr">
    <xsl:param name="bgClr"/>
    <xsl:param name="phClr" select="'###'"/>
    <xsl:value-of select="'#'"/>
    <xsl:variable name="colorAttri" select="child::*"/>
    <xsl:variable name="colorAttriCount" select="count(child::*)"/>
    <xsl:variable name="sldLayoutID">
      <xsl:choose>
        <xsl:when test="ancestor::dsp:drawing/@refBy!=''">
          <xsl:variable name="slideID" select="ancestor::dsp:drawing/@refBy"/>
          <xsl:value-of select="//p:sld[@id=$slideID]/@sldLayoutID"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="ancestor::p:sld/@sldLayoutID"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="sldLayoutRel" select="concat($sldLayoutID,'.rels')"/>
    <xsl:variable name="sldMasterID" select="ancestor::p:sldMaster/@id"/>
    <xsl:variable name="slideMasterID">
      <xsl:choose>
        <xsl:when test="$sldLayoutID!=''">
          <xsl:value-of select="substring-after(//rel:Relationships[@id=$sldLayoutRel]//rel:Relationship[contains(@Type,'slideMaster')]/@Target,'../slideMasters/')"/>
        </xsl:when>
        <xsl:when test="$sldMasterID">
          <xsl:value-of select="$sldMasterID"/>
        </xsl:when>
        <xsl:otherwise>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="themeID" select="//a:theme[@refBy=$slideMasterID]/@id"/>

    <xsl:variable name="schemeColor">
      <xsl:choose>
        <xsl:when test="@val = 'dk1'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr"/>
          <!--<xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'lt1'">
          <!--<xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr"/>-->
          <!--修改字体lt1取值不正确 2013-03-19 liqiuling -->
          <xsl:choose>
          <xsl:when test="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt1/a:srgbClr">
            <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt1/a:srgbClr/@val"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr"/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:when>
        <xsl:when test="@val = 'dk2'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val"/>
          <!--<xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'lt2'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent1'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent2'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent3'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent4'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val"/>
          <!--<xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent5'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'accent6'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
          <!--<xsl:choose>
          <xsl:when test="./a:tint/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:tint/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:alpha/@val and not(ancestor::dsp:sp) and not(ancestor::p:spPr/@bwMode='gray')">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="trans" select="./a:alpha/@val"/>
            </xsl:call-template>
          </xsl:when>-->
          <!--2014.10.29，liuyangyang添加对主题色的hsl调节-->
          <!--<xsl:when test="./a:hue/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'hue'"/>
              <xsl:with-param name="attri" select="./a:hue/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:hueMod/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'hueMod'"/>
              <xsl:with-param name="attri" select="./a:hueMod/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:hueOff/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'hueOff'"/>
              <xsl:with-param name="attri" select="./a:hueOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:sat/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'sat'"/>
              <xsl:with-param name="attri" select="./a:sat/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:satMod/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'satMod'"/>
              <xsl:with-param name="attri" select="./a:satMod/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:satOff/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'satOff'"/>
              <xsl:with-param name="attri" select="./a:satOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lum/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'lum'"/>
              <xsl:with-param name="attri" select="./a:lum/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumMod/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'lumMod'"/>
              <xsl:with-param name="attri" select="./a:lumMod/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumOff/@val and not(ancestor::dsp:sp)">
            <xsl:call-template name="transByHSL">
              <xsl:with-param name="rgb" select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val"/>
              <xsl:with-param name="pattern" select="'lumOff'"/>
              <xsl:with-param name="attri" select="./a:lumOff/@val"/>
            </xsl:call-template>
          </xsl:when>-->
          <!--end-->
          <!--<xsl:otherwise>
            <xsl:value-of select="concat('#',//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'hlink'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:hlink/a:srgbClr/@val"/>
        </xsl:when>
        <xsl:when test="@val = 'folHlink'">
          <xsl:value-of select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:clrScheme/a:folHlink/a:srgbClr/@val"/>
        </xsl:when>
        <!--修改bug图形背景色转换不正确 增加phClr的情形  liqiuling 2013-03-08-->
        <xsl:when test="@val = 'bg1'">
          <xsl:value-of select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg1,'#')"/>
          <!-- <xsl:choose>-->
          <!-- 亮度由lumOff和lumMod两个控制，若lumOff存在说明是变亮，若只有lumMod存在则变暗 -->
          <!--<xsl:when test="./a:lumOff/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg1,'#')"/>
              <xsl:with-param name="trans" select="./a:lumOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumMod/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg1,'#')"/>
              <xsl:with-param name="trans" select="concat(substring-before( ./a:lumMod/@val, '%') - 100, '%')"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//p:presentation/p:sldMaster/p:clrMap/@bg1"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'phClr'">
          <xsl:choose>
            <xsl:when test="$phClr = '###'">
              <xsl:value-of select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg1,'#')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring-after($phClr,'#')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="@val = 'bg2'">
          <xsl:value-of select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg2,'#')"/>
          <!--<xsl:choose>
          <xsl:when test="./a:lumOff/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg2,'#')"/>
              <xsl:with-param name="trans" select="./a:lumOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumMod/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@bg2,'#')"/>
              <xsl:with-param name="trans" select="concat(substring-before( ./a:lumMod/@val, '%') - 100, '%')"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//p:presentation/p:sldMaster/p:clrMap/@bg2"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'tx1'">
          <xsl:value-of select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx1,'#')"/>
          <!-- <xsl:choose>
          <xsl:when test="./a:lumOff/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx1,'#')"/>
              <xsl:with-param name="trans" select="./a:lumOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumMod/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx1,'#')"/>
              <xsl:with-param name="trans" select="concat(substring-before( ./a:lumMod/@val, '%') - 100, '%')"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//p:presentation/p:sldMaster/p:clrMap/@tx1"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:when test="@val = 'tx2'">
          <xsl:value-of select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx2,'#')"/>
          <!--  <xsl:choose>
          <xsl:when test="./a:lumOff/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx2,'#')"/>
              <xsl:with-param name="trans" select="./a:lumOff/@val"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="./a:lumMod/@val">
            <xsl:call-template name="schemeClrTransparency">
              <xsl:with-param name="schemeClr" select="substring-after(//p:presentation/p:sldMaster/p:clrMap/@tx2,'#')"/>
              <xsl:with-param name="trans" select="concat(substring-before( ./a:lumMod/@val, '%') - 100, '%')"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//p:presentation/p:sldMaster/p:clrMap/@tx2"/>
          </xsl:otherwise>
        </xsl:choose>-->
        </xsl:when>
        <xsl:otherwise>
          <!--xsl:value-of select="'#4F81BD'"/-->
          <xsl:value-of select="'#000000'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:call-template name="calClrFromScheme">
      <xsl:with-param name="colorAttri" select="$colorAttri"/>
      <xsl:with-param name="colorAttriCount" select="$colorAttriCount"/>
      <xsl:with-param name="schemeColor" select="$schemeColor"/>
      <xsl:with-param name="bgClr" select="$bgClr"/>
    </xsl:call-template>
  </xsl:template>
  <!-- end 2014-01-13, tangjiang -->
  <!-- end 2015-03-06,liuyangyang-->
  <!--2015-03-06,liuyangyang 添加主题色转换节点的递归计算-->
  <xsl:template name="calClrFromScheme">
    <xsl:param name="colorAttri"/>
    <xsl:param name="colorAttriCount"/>
    <xsl:param name="schemeColor"/>
    <xsl:param name="bgClr"/>
    <xsl:choose>
      <xsl:when test ="$colorAttriCount = 0">
        <xsl:value-of select="$schemeColor"/>
      </xsl:when>
      <xsl:when test="$colorAttriCount = 1">
        <xsl:call-template name="transByHSL">
          <xsl:with-param name="rgb" select="$schemeColor"/>
          <xsl:with-param name="attri" select="$colorAttri[1]/@val"/>
          <xsl:with-param name="pattern">
            <xsl:apply-templates select="$colorAttri[1]" mode="getPattern"/>
          </xsl:with-param>
          <xsl:with-param name="bgClr" select="$bgClr"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="$colorAttriCount &gt; 1">
        <xsl:call-template name="transByHSL">
          <xsl:with-param name="rgb">
            <xsl:call-template name="calClrFromScheme">
              <xsl:with-param name="colorAttri" select="$colorAttri"/>
              <xsl:with-param name="colorAttriCount" select="$colorAttriCount - 1"/>
              <xsl:with-param name="schemeColor" select="$schemeColor"/>
              <xsl:with-param name="bgClr" select="$bgClr"/>
            </xsl:call-template>
          </xsl:with-param>
          <xsl:with-param name="attri" select="$colorAttri[$colorAttriCount]/@val"/>
          <xsl:with-param name="pattern">
            <!-- <xsl:value-of select="name($colorAttri[$colorAttriCount])"/>-->
            <xsl:apply-templates select="$colorAttri[$colorAttriCount]" mode="getPattern"/>
          </xsl:with-param>
          <xsl:with-param name="bgClr" select="$bgClr"/>
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="a:sat" mode="getPattern">
    <xsl:value-of select="'sat'"/>
  </xsl:template>
  <xsl:template match="a:satOff" mode="getPattern">
    <xsl:value-of select="'satOff'"/>
  </xsl:template>
  <xsl:template match="a:satMod" mode="getPattern">
    <xsl:value-of select="'satMod'"/>
  </xsl:template>
  <xsl:template match="a:hue" mode="getPattern">
    <xsl:value-of select="'hue'"/>
  </xsl:template>
  <xsl:template match="a:hueOff" mode="getPattern">
    <xsl:value-of select="'hueOff'"/>
  </xsl:template>
  <xsl:template match="a:hueMod" mode="getPattern">
    <xsl:value-of select="'hueMod'"/>
  </xsl:template>
  <xsl:template match="a:lum" mode="getPattern">
    <xsl:value-of select="'lum'"/>
  </xsl:template>
  <xsl:template match="a:lumOff" mode="getPattern">
    <xsl:value-of select="'lumOff'"/>
  </xsl:template>
  <xsl:template match="a:lumMod | a:shade" mode="getPattern">
    <xsl:value-of select="'lumMod'"/>
  </xsl:template>
  <xsl:template match="a:tint" mode="getPattern">
    <xsl:value-of select="'tint'"/>
  </xsl:template>
  <xsl:template match="*" mode="getPattern">
    <xsl:value-of select ="'other'"/>
  </xsl:template>
  <!-- end 2015-03-06,liuyangyang-->
  <!--2014-01-13, tangjiang,添加OOXML到UOF文字表式样的转换,文本框渐变填充 start -->
  <xsl:template name="schemeClrTransparency">
    <xsl:param name="trans" select="'100%'"/>
    <xsl:param name="schemeClr" select="'FFFFFF'"/>
    <xsl:variable name="redH">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,1,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="redL">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,2,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenH">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,3,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenL">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,4,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueH">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,5,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueL">
      <xsl:call-template name="hexToDec">
        <xsl:with-param name="HexCode"  select="substring($schemeClr,6,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="rgbMin" select="'0'"/>
    <xsl:variable name="rgbMax" select="'255'"/>
    <xsl:variable name="redColorNum" select="($redH * 16 + $redL) + $rgbMax * (substring($trans,1,string-length($trans)-1) div 100)"/>
    <xsl:variable name="greenColorNum" select="($greenH * 16 + $greenL) +  $rgbMax * (substring($trans,1,string-length($trans)-1) div 100)"/>
    <xsl:variable name="blueColorNum" select="($blueH * 16 + $blueL) +  $rgbMax * (substring($trans,1,string-length($trans)-1) div 100)"/>
    <xsl:variable name="redColorNum_r">
      <xsl:choose>
        <xsl:when test="$redColorNum &lt; $rgbMin">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:when test="$redColorNum &gt; $rgbMax">
          <xsl:value-of select="'250'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$redColorNum"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="greenColorNum_r">
      <xsl:choose>
        <xsl:when test="$greenColorNum &lt; $rgbMin">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:when test="$greenColorNum &gt; $rgbMax">
          <xsl:value-of select="'250'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$greenColorNum"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="blueColorNum_r">
      <xsl:choose>
        <xsl:when test="$blueColorNum &lt; $rgbMin">
          <xsl:value-of select="'0'"/>
        </xsl:when>
        <xsl:when test="$blueColorNum &gt; $rgbMax">
          <xsl:value-of select="'250'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$blueColorNum"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="redColorH">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($redColorNum_r div 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="redColorL">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($redColorNum_r mod 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenColorH">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($greenColorNum_r div 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="greenColorL">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($greenColorNum_r mod 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueColorH">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($blueColorNum_r div 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="blueColorL">
      <xsl:call-template name="decToHex">
        <xsl:with-param name="DecCode" select="floor($blueColorNum_r mod 16)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat('#',$redColorH,$redColorL,$greenColorH,$greenColorL,$blueColorH,$blueColorL)"/>
  </xsl:template>
  <xsl:template name="hexToDec" >
    <xsl:param name="HexCode" select="'0'"/>
    <xsl:choose>
      <xsl:when test=" $HexCode = 'A' or $HexCode = 'a' ">
        <xsl:value-of select="'10'"/>
      </xsl:when>
      <xsl:when test=" $HexCode = 'B' or $HexCode = 'b' ">
        <xsl:value-of select="'11'"/>
      </xsl:when>
      <xsl:when test=" $HexCode = 'C' or $HexCode = 'c' ">
        <xsl:value-of select="'12'"/>
      </xsl:when>
      <xsl:when test=" $HexCode = 'D' or $HexCode = 'd' ">
        <xsl:value-of select="'13'"/>
      </xsl:when>
      <xsl:when test=" $HexCode = 'E' or $HexCode = 'e' ">
        <xsl:value-of select="'14'"/>
      </xsl:when>
      <xsl:when test=" $HexCode = 'F' or $HexCode = 'f' ">
        <xsl:value-of select="'15'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select=" $HexCode "/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="decToHex" >
    <xsl:param name="DecCode" select="'0'"/>
    <xsl:choose>
      <xsl:when test=" $DecCode = 10">
        <xsl:value-of select="'A'"/>
      </xsl:when>
      <xsl:when test=" $DecCode = 11">
        <xsl:value-of select="'B'"/>
      </xsl:when>
      <xsl:when test=" $DecCode = 12">
        <xsl:value-of select="'C'"/>
      </xsl:when>
      <xsl:when test=" $DecCode = 13 ">
        <xsl:value-of select="'D'"/>
      </xsl:when>
      <xsl:when test=" $DecCode = 14 ">
        <xsl:value-of select="'E'"/>
      </xsl:when>
      <xsl:when test=" $DecCode = 15 ">
        <xsl:value-of select="'F'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select=" $DecCode "/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- end 2014-01-13, tangjiang -->

  <!--20090413schemeClr模式-->
  <!--xsl:template match="a:schemeClr">
		<xsl:call-template name="findSchemeClrPosition">
			<xsl:with-param name="schemeClrVal" select="@val"/>
		</xsl:call-template>
	</xsl:template-->
  <!--schemeClr颜色定义在theme.xml中,此模板为找出定义在哪个theme文件中-->
  <xsl:template name="findSchemeClrPosition">
    <xsl:param name="schemeClrVal"/>
    <!--寻找定义该shceme颜色的theme节点-->
    <xsl:variable name="sldLayoutPosition" select="ancestor::p:sld/@sldLayoutID"/>

    <xsl:apply-templates select="ancestor::p:sld" mode="findTheme">
      <xsl:with-param name="passsldLayoutID" select="$sldLayoutPosition"/>
      <xsl:with-param name="passSchemeClrVal" select="$schemeClrVal"/>
    </xsl:apply-templates>

  </xsl:template>

  <xsl:template match="p:sld" mode="findTheme">
    <xsl:param name="passsldLayoutID"/>
    <xsl:param name="passSchemeClrVal"/>
    <xsl:apply-templates select="../rel:Relationships" mode="findTheme">
      <xsl:with-param name="sldLayoutID" select="$passsldLayoutID"/>
      <xsl:with-param name="passClrVal" select="$passSchemeClrVal"/>
    </xsl:apply-templates>
  </xsl:template>
  <xsl:template match="rel:Relationships" mode="findTheme">
    <xsl:param name="sldLayoutID"/>
    <xsl:param name="passClrVal"/>
    <xsl:for-each select="rel:Relationship">
      <xsl:if test="contains(@Target,$sldLayoutID)">
        <xsl:if test="../following::a:theme">
          <xsl:apply-templates select="../following-sibling::a:theme[1]/a:themeElements/a:clrScheme">
            <xsl:with-param name="schemeClrValInTheme" select="$passClrVal"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="a:clrScheme">
    <xsl:param name="schemeClrValInTheme"/>
    <xsl:choose>
      <xsl:when test="$schemeClrValInTheme='accent1'">
        <xsl:value-of select="concat('#',./a:accent1/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='accent2'">
        <xsl:value-of select="concat('#',./a:accent2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='accent3'">
        <xsl:value-of select="concat('#',./a:accent3/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='accent4'">
        <xsl:value-of select="concat('#',./a:accent4/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='accent5'">
        <xsl:value-of select="concat('#',./a:accent5/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='accent6'">
        <xsl:value-of select="concat('#',./a:accent6/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='bg1'">
        <xsl:value-of select="concat('#',./a:dk1/a:sysClr/@lastVal)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='bg2'">
        <xsl:value-of select="concat('#',./a:dk2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='dk1'">
        <xsl:value-of select="concat('#',./a:dk1/a:sysClr/@lastVal)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='dk2'">
        <xsl:value-of select="concat('#',./a:dk2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='folHlink'">
        <xsl:value-of select="concat('#',./a:folHlink/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='hlink'">
        <xsl:value-of select="concat('#',./a:hlink/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='lt1'">
        <xsl:value-of select="concat('#',./a:lt1/a:sysClr/@lastVal)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='lt2'">
        <xsl:value-of select="concat('#',./a:lt2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$schemeClrValInTheme='phClr'">
      </xsl:when>
    </xsl:choose>

  </xsl:template>





  <!--渐变填充-->
  <!--对a:gsLst/a:gs的处理-->
  <!--输出翻转信息-->
  <xsl:template match="a:gradFill" mode="flip">
    <xsl:value-of select="@flip"/>
  </xsl:template>
  <!--输出是否跟随旋转信息-->
  <!--10.02.01马有旭<xsl:template match="a:gradFill" mode="rotWithShape">
    <xsl:choose>
      <xsl:when test="@rotWithShape='0'">
        <xsl:value-of select="false"/>
      </xsl:when>
      <xsl:when test="@rotWithShape='1'">
        <xsl:value-of select="true"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>-->
  <!--边界-->
  <xsl:template name="borderLine">
    <xsl:variable name="sborderLine">
      <xsl:apply-templates select="a:gs" mode="first"/>
    </xsl:variable>
    <xsl:variable name="eborderLine">
      <xsl:apply-templates select="a:gs" mode="last"/>
    </xsl:variable>
    <xsl:value-of select="$eborderLine div 1000-$sborderLine div 1000"/>
  </xsl:template>
  <!--第一个a:gs的属性信息-->
  <xsl:template match="a:gs" mode="first">
    <xsl:if test="position()=1">
      <xsl:value-of select="@pos"/>
    </xsl:if>
  </xsl:template>
  <!--最后一个a:gs的属性信息-->
  <xsl:template match="a:gs" mode="last">
    <xsl:if test="position()=last()">
      <xsl:value-of select="@pos"/>
    </xsl:if>
  </xsl:template>
  <!--起始色-->
  <xsl:template match="a:gs" mode="firstClr">
    <xsl:if test="position()=1">
      <xsl:call-template name="colorChoice"/>
    </xsl:if>
  </xsl:template>
  <!--终止色-->
  <xsl:template match="a:gs" mode="lastClr">
    <xsl:if test="position()=last()">
      <xsl:call-template name="colorChoice"/>
    </xsl:if>
  </xsl:template>
  <!--透明度,为100%-其值-->
  <xsl:template match="a:gs" mode="alpha">
    <xsl:if test="//a:alpha">
      <xsl:value-of select="100- //a:alpha/@val div 1000"/>
    </xsl:if>
  </xsl:template>
  <!--对a:lin的处理-->

  <xsl:template match="a:lin" mode="scaled"/>
  <xsl:template match="a:gsLst">
    <!--bsla:b,borderLine;s:startClr,l:lastClr,a:alpha-->
    <xsl:param name="bsla"/>
    <xsl:choose>
      <!--边界-->
      <xsl:when test="$bsla='1'">
        <xsl:call-template name="borderLine"/>
      </xsl:when>
      <!--起始色-->
      <xsl:when test="$bsla='2'">
        <xsl:apply-templates select="a:gs" mode="firstClr"/>
      </xsl:when>
      <!--终止色-->
      <xsl:when test="$bsla='3'">
        <xsl:apply-templates select="a:gs" mode="lastClr"/>
      </xsl:when>
      <!--透明度-->
      <xsl:when test="$bsla='4'">
        <xsl:apply-templates select="a:gs" mode="alpha"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--uof的渐变填充-->
	  <xsl:template match="a:gradFill">
      <xsl:param name="phClr"/>
      <!--liuyangyang phClr-->
	  <图:渐变_800D>
    <!--<图:渐变 uof:locID="g0037" uof:attrList="起始色 终止色 种子类型 起始浓度 终止浓度 渐变方向 边界 种子X位置 种子Y位置 类型">-->
      <xsl:variable name="angle">
        <xsl:choose>
          <!-- 10.02.01 马有旭 修改-->
          <xsl:when test="a:lin">
            
            <!--2014-03-09, pengxin, 修改OOXML到UOF母版-渐变颜色方向错误 start -->
            <xsl:value-of select="(360 - (round(a:lin/@ang div 60000 div 45)*45+90)) mod 360"/>
            <!-- <xsl:value-of select="(360 - round(a:lin/@ang div 60000 div 45)*45+90) mod 360"/>
            <xsl:value-of select="round(((360-(a:lin/@ang div 60000+90)) mod 360) div 45)*45"/>-->
            <!-- end 2014-03-09, pengxin -->
            
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="种子类型_8010">
        <xsl:choose>
          <xsl:when test="a:path">
            <xsl:choose>
              <xsl:when test="a:path/@path='rect'">square</xsl:when>
              <xsl:when test="a:path/@path='circle'">oval</xsl:when>
              <!--10.01.31马有旭 修改<xsl:when test="a:path/@path='shape'">linear</xsl:when>-->
              <xsl:when test="a:path/@path='shape'">square</xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>linear</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="起始浓度_8011">1.0</xsl:attribute>
      <xsl:attribute name="终止浓度_8012">1.0</xsl:attribute>
      <xsl:attribute name="渐变方向_8013">
        <xsl:value-of select="$angle"/>
      </xsl:attribute>
      <xsl:choose>

        <!--2014-03-09, pengxin, 修改OOXML到UOF母版-渐变颜色方向错误，互操作渐变反向错误 start -->
        <xsl:when test="$angle='0' or $angle='135' or $angle='225' or $angle='270'or $angle='90'">
        <!-- <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">-->
        <!-- end 2014-03-09, pengxin 修改OOXML到UOF母版-渐变颜色方向错误，互操作渐变反向错误-->
          
         <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice">
                  <xsl:with-param name="phClr" select="$phClr"/>
                  <!--liuyangyang phClr-->
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice">
                  <xsl:with-param name="phClr" select="$phClr"/>
                  <!--liuyangyang phClr-->
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice">
                  <xsl:with-param name="phClr" select="$phClr"/>
                  <!--liuyangyang phClr-->
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice">
                  <xsl:with-param name="phClr" select="$phClr"/>
                  <!--liuyangyang phClr-->
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>

      <!--<xsl:choose>
        <xsl:when test="$angle &lt; 180">
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="图:起始色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="图:终止色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="图:终止色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="图:起始色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>-->

      <xsl:attribute name="边界_8014">50</xsl:attribute>
        <!--10.01.31马有旭<xsl:choose>
          <xsl:when test="not(a:path)">
            <xsl:variable name="startClr">
              <xsl:for-each select="a:gsLst/a:gs">
                <xsl:sort select="@pos" data-type="number"/>
                <xsl:if test="position()=1">
                  <xsl:value-of select="@pos"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:variable>
            <xsl:variable name="endClr">
              <xsl:for-each select="a:gsLst/a:gs">
                <xsl:sort select="@pos" data-type="number"/>
                <xsl:if test="position()=last()">
                  <xsl:value-of select="@pos"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:variable>
            <xsl:value-of select="($startClr+$endClr) div 2 div 1000"/>
          </xsl:when>
          <xsl:when test="a:path/@path='rect'">
            <xsl:value-of select="5"/>
          </xsl:when>
        </xsl:choose>

      </xsl:attribute>-->
      <!--<xsl:apply-templates select="a:gsLst">
          <xsl:with-param name="bsla" select="1"/>
        </xsl:apply-templates>-->
      <xsl:choose>
        <xsl:when test="not(a:path)">
          <xsl:attribute name="种子X位置_8015">100</xsl:attribute>
          <xsl:attribute name="种子Y位置_8016">100</xsl:attribute>
        </xsl:when>
        <xsl:when test="a:path/@path='rect'">
          <xsl:choose>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>

	  </图:渐变_800D>
  </xsl:template>
  <!--HSL的主转换模板-->
  <xsl:template name="HSLtoRGB">
    <xsl:param name="H"/>
    <xsl:param name="S"/>
    <xsl:param name="L"/>
    <xsl:choose>
      <xsl:when test="$S='0'">
        <xsl:variable name="tmpcolor">
          <xsl:call-template name="decimalToHex">
            <xsl:with-param name="result" select="$L*255"/>
          </xsl:call-template>
        </xsl:variable>
        <!--#-->
        <xsl:value-of select="$tmpcolor"/>
        <xsl:value-of select="$tmpcolor"/>
        <xsl:value-of select="$tmpcolor"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="temp2">
          <xsl:choose>
            <xsl:when test="$L &lt; 0.5">
              <xsl:value-of select="$L * (1+$S)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$L + $S - $L * $S"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="temp1">
          <xsl:value-of select="2 * $L - $temp2"/>
        </xsl:variable>
        <!--计算R值-->
        <xsl:variable name="rValue">
          <xsl:call-template name="RGB">
            <xsl:with-param name="temp1" select="$temp1"/>
            <xsl:with-param name="temp2" select="$temp2"/>
            <xsl:with-param name="hue" select="$H + 1 div 3"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="decimalToHex">
          <xsl:with-param name="result" select="$rValue * 255"/>
        </xsl:call-template>
        <!--计算G值-->
        <xsl:variable name="gValue">
          <xsl:call-template name="RGB">
            <xsl:with-param name="temp1" select="$temp1"/>
            <xsl:with-param name="temp2" select="$temp2"/>
            <xsl:with-param name="hue" select="$H"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="decimalToHex">
          <xsl:with-param name="result" select="$gValue * 255"/>
        </xsl:call-template>
        <!--计算B值-->
        <xsl:variable name="bValue">
          <xsl:call-template name="RGB">
            <xsl:with-param name="temp1" select="$temp1"/>
            <xsl:with-param name="temp2" select="$temp2"/>
            <xsl:with-param name="hue" select="$H - 1 div 3"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="decimalToHex">
          <xsl:with-param name="result" select="$bValue * 255"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--RGB模板，计算RGB的值，结果为小数-->
  <xsl:template name="RGB">
    <xsl:param name="temp1"/>
    <xsl:param name="temp2"/>
    <xsl:param name="hue"/>
    <!--用于转换hue出现小于零的情况-->
    <xsl:param name="color">
      <xsl:choose>
        <xsl:when test="$hue &lt; 0">
          <xsl:value-of select="$hue + 1"/>
        </xsl:when>
        <xsl:when test="$hue &gt; 1">
          <xsl:value-of select="$hue - 1"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$hue"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:param>
    <xsl:choose>
      <xsl:when test="6 * $color &lt; 1">
        <xsl:value-of select="$temp1 + ($temp2 - $temp1)*6*$color"/>
      </xsl:when>
      <xsl:when test="2 * $color &lt; 1">
        <xsl:value-of select="$temp2"/>
      </xsl:when>
      <xsl:when test="3 * $color &lt; 2">
        <xsl:value-of select="$temp1 + ($temp2 - $temp1) * (2 div 3 - $color) * 6"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$temp1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--把用RGB表示的数(0-255)转换成两位两位的十六进制数-->
  <xsl:template name="decimalToHex">
    <xsl:param name="result"/>
    <xsl:param name="deci">
      <xsl:choose>
        <xsl:when test="$result &lt; 0">0</xsl:when>
        <xsl:when test="$result &gt; 255">255</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="round($result)"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:param>
    <xsl:variable name="strHex" select="'0123456789ABCDEF'"/>
    <xsl:value-of select="substring($strHex,floor($deci div 16)+1,1)"/>
    <xsl:value-of select="substring($strHex,($deci mod 16)+1,1)"/>
  </xsl:template>
  <!--从十六进制转成十进制，只转两位的十六进制-->
  <xsl:template name="hexToDecimal">
    <!--从外界传进的两位十六进制值-->
    <xsl:param name="hex"/>
    <!--十位数-->
    <xsl:variable name="num1">
      <xsl:call-template name="hexAl">
        <xsl:with-param name="al" select="substring($hex,1,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <!--个位数-->
    <xsl:variable name="num2">
      <xsl:call-template name="hexAl">
        <xsl:with-param name="al" select="substring($hex,2,1)"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="$num1 * 16 + $num2"/>
  </xsl:template>
  <xsl:template name="hexAl">
    <xsl:param name="al"/>
    <xsl:choose>
      <xsl:when test="$al = 'a' or $al = 'A'">10</xsl:when>
      <xsl:when test="$al = 'b' or $al = 'B'">11</xsl:when>
      <xsl:when test="$al = 'c' or $al = 'C'">12</xsl:when>
      <xsl:when test="$al = 'd' or $al = 'D'">13</xsl:when>
      <xsl:when test="$al = 'e' or $al = 'E'">14</xsl:when>
      <xsl:when test="$al = 'f' or $al = 'F'">15</xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$al"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--图案填充-->
  <xsl:template name="patternFill">
    <xsl:param name="pa"/>
    <xsl:choose>
      <xsl:when test="$pa = '1'">
        <xsl:choose>
          <!--2011-3-20罗文甜修改图案类型对应-->
          <xsl:when test="@prst ='pct5'">ptn001</xsl:when>
          <xsl:when test="@prst ='pct50'">ptn007</xsl:when>
          <xsl:when test="@prst ='ltDnDiag'">ptn013</xsl:when>
          <xsl:when test="@prst ='ltVert'">ptn019</xsl:when>
          <xsl:when test="@prst = 'dashDnDiag'">ptn025</xsl:when>
          <xsl:when test="@prst ='zigZag'">ptn031</xsl:when>
          <xsl:when test="@prst ='divot'">ptn037</xsl:when>
          <xsl:when test="@prst ='smGrid'">ptn043</xsl:when>
          <xsl:when test="@prst ='pct10'">ptn002</xsl:when>
          <xsl:when test="@prst ='pct60'">ptn008</xsl:when>
          <xsl:when test="@prst ='ltUpDiag'">ptn014</xsl:when>
          <xsl:when test="@prst ='ltHorz'">ptn020</xsl:when>
          <xsl:when test="@prst = 'dashUpDiag'">ptn026</xsl:when>
          <xsl:when test="@prst ='wave'">ptn032</xsl:when>
          <xsl:when test="@prst ='dotGrid'">ptn038</xsl:when>
          <xsl:when test="@prst ='lgGrid'">ptn044</xsl:when>
          <xsl:when test="@prst ='pct20'">ptn003</xsl:when>
          <xsl:when test="@prst ='pct70'">ptn009</xsl:when>
          <xsl:when test="@prst ='dkDnDiag'">ptn015</xsl:when>
          <xsl:when test="@prst ='narVert'">ptn021</xsl:when>
          <xsl:when test="@prst = 'dashHorz'">ptn027</xsl:when>
          <xsl:when test="@prst ='diagBrick'">ptn033</xsl:when>
          <xsl:when test="@prst ='dotDmnd'">ptn039</xsl:when>
          <xsl:when test="@prst ='smCheck'">ptn045</xsl:when>          
          <xsl:when test="@prst ='pct25'">ptn004</xsl:when>          
          <xsl:when test="@prst ='pct75'">ptn010</xsl:when>          
          <xsl:when test="@prst ='dkUpDiag'">ptn016</xsl:when>          
          <xsl:when test="@prst ='narHorz'">ptn022</xsl:when>
          <xsl:when test="@prst = 'dashVert'">ptn028</xsl:when>      
          <xsl:when test="@prst ='horzBrick'">ptn034</xsl:when>          
          <xsl:when test="@prst ='shingle'">ptn040</xsl:when>          
          <xsl:when test="@prst ='lgCheck'">ptn046</xsl:when>          
          <xsl:when test="@prst ='pct30'">ptn005</xsl:when>          
          <xsl:when test="@prst ='pct80'">ptn011</xsl:when>          
          <xsl:when test="@prst ='wdDnDiag'">ptn017</xsl:when>          
          <xsl:when test="@prst ='dkVert'">ptn023</xsl:when>          
          <xsl:when test="@prst ='smConfetti'">ptn029</xsl:when>          
          <xsl:when test="@prst ='weave'">ptn035</xsl:when> 
          <xsl:when test="@prst ='trellis'">ptn041</xsl:when>          
          <xsl:when test="@prst ='openDmnd'">ptn047</xsl:when>          
          <xsl:when test="@prst ='pct40'">ptn006</xsl:when>          
          <xsl:when test="@prst ='pct90'">ptn012</xsl:when>
          <xsl:when test="@prst ='wdUpDiag'">ptn018</xsl:when>
          <xsl:when test="@prst ='dkHorz'">ptn024</xsl:when>          
          <xsl:when test="@prst ='lgConfetti'">ptn030</xsl:when>
          <xsl:when test="@prst ='plaid'">ptn036</xsl:when>          
          <xsl:when test="@prst ='sphere'">ptn042</xsl:when>          
          <xsl:when test="@prst ='solidDmnd'">ptn048</xsl:when>
          
          <!--xsl:when test="@prst = 'cross'">ptn043</xsl:when>
          <xsl:when test="@prst ='vert'">ptn029</xsl:when>
          <xsl:when test="@prst ='diagCross'">ptn047</xsl:when> 
          <xsl:when test="@prst ='dnDiag'">ptn005</xsl:when> 
          <xsl:when test="@prst ='ltUpDiag'">ptn011</xsl:when>
          <xsl:when test="@prst ='upDiag'">ptn013</xsl:when>  
          <xsl:when test="@prst ='horz'">ptn021</xsl:when-->
          
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$pa = '2'">
        <xsl:apply-templates select="a:fgClr"/>
      </xsl:when>
      <xsl:when test="$pa = '3'">
        <xsl:apply-templates select="a:bgClr"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--fgClr前景色-->
  <xsl:template match="a:fgClr">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>
  <!--bgClr背景色-->
  <xsl:template match="a:bgClr">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>
  <!--uof的图案填充-->
  <xsl:template match="a:pattFill">
	  <图:图案_800A>
      <xsl:attribute name="类型_8008">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="1"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="前景色_800B">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="2"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="背景色_800C">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="3"/>
        </xsl:call-template>
      </xsl:attribute>
		  <!--<xsl:attribute name="图形引用_8007"></xsl:attribute> 李娟 11.11.10-->
    </图:图案_800A>
  </xsl:template>
  <!--uof的纯色填充-->
  <xsl:template match="a:solidFill">
    <xsl:param name="phClr"/>
    <图:颜色_8004>
      <xsl:call-template name="colorChoice">
        <xsl:with-param name="phClr" select="$phClr"/>
        <!--liuyangyang phClr-->
      </xsl:call-template>
    </图:颜色_8004>
  </xsl:template>
	
	<!--线条颜色,只需得到颜色值-->
	<xsl:template match="a:solidFill" mode="ln">
		<图:线颜色_8058>
		<xsl:call-template name="colorChoice"/>
		</图:线颜色_8058>
	</xsl:template>
	<xsl:template match="a:solidFill" mode="tabln">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
  

  <!--组填充-->
  <xsl:template match="a:grpFill">
  </xsl:template>

  <!--无填充-->
  <xsl:template match="a:noFill">
  </xsl:template>





  <!--xsl:template match="/">
		<lllll>
			<xsl:apply-templates select="p:presentation/p:sldMaster//a:blipFill"/>
		</lllll>
	</xsl:template-->


  <!--图片填充-->

  <xsl:template match="a:blipFill">
	  <图:图片_8005>
      <xsl:attribute name="位置_8006">
        <xsl:choose>
          <xsl:when test="a:stretch">
            <xsl:value-of select="'stretch'"/>
          </xsl:when>
          <xsl:when test="a:tile">
            <xsl:value-of select="'tile'"/>
          </xsl:when>
          <!--2010.04.12 myx add -->
          <xsl:when test="a:srcRect">tile</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'center'"/>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:attribute>
      <xsl:attribute name="图形引用_8007">
        <xsl:call-template name="findPicName"/>
      </xsl:attribute>
    </图:图片_8005>
  </xsl:template>
  <xsl:template name="findPicName">
    <xsl:variable name="findPictureName">
      <xsl:value-of select=".//a:blip/@r:embed"/>
    </xsl:variable>
    <xsl:call-template name="findRelatedRelationships">
      <xsl:with-param name="id">
        <xsl:value-of select="$findPictureName"/>
      </xsl:with-param>
    </xsl:call-template>

  </xsl:template>
  <xsl:template name="findRelatedRelationships">
    <xsl:param name="id"/>


    <!--增加smartart图片填充的功能 liqiuling 2013-03-22 start-->
     <!--09.12.21 马有旭 修改 -->
    <xsl:choose>
      <xsl:when test="ancestor::p:cSld | ancestor::p:txStyles">
        <xsl:for-each select="ancestor::p:cSld | ancestor::p:txStyles">
          <xsl:apply-templates select="../following-sibling::rel:Relationships[1]">
            <xsl:with-param name="targetID">
              <xsl:value-of select="$id"/>
            </xsl:with-param>
          </xsl:apply-templates>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="ancestor::dsp:drawing">
        <xsl:variable name="dspid" select="ancestor::dsp:drawing/@id"/>
        <xsl:for-each select="//dsp:drawingRelationship">
          <xsl:if test="contains(@id,$dspid)">
            <xsl:call-template name="dspRelationships">
              <xsl:with-param name="targetID">
                <xsl:value-of select="$id"/>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <!--liuyangyang 2015-04-09 添加主题背景的图片引用-->
      <xsl:when test="ancestor::a:theme">
        <xsl:value-of select="concat(substring-before(ancestor::a:theme/@id,'.xml'),$id)"/>
      </xsl:when>
    </xsl:choose>
    
  </xsl:template>
  

  <xsl:template match="rel:Relationships">
    <xsl:param name="targetID"/>
    <xsl:for-each select="rel:Relationship">
      <xsl:if test="./@Id = $targetID">
        <xsl:variable name="slideName">
          <xsl:value-of select="substring-before(ancestor::rel:Relationships/@id, '.xml.rels')"/>
        </xsl:variable>
        <xsl:value-of select="concat($slideName, ./@Id)"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="dspRelationships">
    <xsl:param name="targetID"/>
    <xsl:for-each select="rel:Relationship">
      <xsl:if test="./@Id = $targetID">
        <xsl:variable name="slideName">
          <xsl:value-of select="substring-before(ancestor::rel:Relationships/@id, '.xml.rels')"/>
        </xsl:variable>
        <xsl:value-of select="concat($slideName, ./@Id)"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <!--增加smartart图片填充的功能 liqiuling 2013-03-22 end--> 
  
  <!--六种填充在这里统一控制-->
  <xsl:template name="sixFill">
    <xsl:choose>
      <xsl:when test="a:gradFill">
        <xsl:apply-templates select="a:gradFill"/>
      </xsl:when>
      <xsl:when test="a:pattFill">
        <xsl:apply-templates select="a:pattFill"/>
      </xsl:when>
      <xsl:when test="a:solidFill">
        <xsl:apply-templates select="a:solidFill"/>
      </xsl:when>
      <xsl:when test="a:grpFill">
        <xsl:apply-templates select="a:grpFill"/>
      </xsl:when>
      <xsl:when test="a:noFill">
        <xsl:apply-templates select="a:noFill"/>
      </xsl:when>
      <xsl:when test="a:blipFill">
        <xsl:apply-templates select="a:blipFill"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>


  <!--背景填充-->
  <xsl:template match="p:bg">
    <xsl:choose>
      <xsl:when test="p:bgPr">
        <xsl:apply-templates select="p:bgPr"/>
      </xsl:when>
      <xsl:when test="p:bgRef">
        <xsl:apply-templates select="p:bgRef"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="p:bgPr">
    <xsl:call-template name="sixFill"/>
  </xsl:template>
  <!--母版背景 2012.11.16 李秋玲-->
  <xsl:template match="a:fillStyleLst">
    <xsl:call-template name="sixFill"/>
  </xsl:template>
  <xsl:template match="a:bgFillStyleLst">
    <xsl:call-template name="sixFill"/>
  </xsl:template>

  <xsl:template match="p:bgRef">
    <!--liuyangyang 2015-04-09 修改背景填充错误 start-->
    <xsl:choose>
      <xsl:when test="number(@idx) &lt; 1004 and number(@idx) &gt; 1000">
        <xsl:variable name="masterID" select="concat(/p:presentation/p:sldMaster/@id,'.rels')"/>
        <xsl:call-template name="getThemeBgFill">
          <xsl:with-param name="idx" select="@idx"/>
          <xsl:with-param name="themeID">
            <xsl:value-of select="substring-after(//rel:Relationships[@id=$masterID]/rel:Relationship[contains(@Type,'theme')]/@Target,'../theme/')"/>
          </xsl:with-param>
          <xsl:with-param name="phClr">
            <xsl:apply-templates select="./a:schemeClr"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="a:schemeClr">
          <图:颜色_8004>
            <xsl:variable name="schemeClrName">
              <xsl:value-of select="./a:schemeClr/@val"/>
            </xsl:variable>
            <!--母版背景 2012.11.16 李秋玲
      <xsl:variable name="schemeClr">-->
            <xsl:choose>
              <xsl:when test="$schemeClrName='bg1'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@bg1"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='bg2'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@bg2"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='tx1'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@tx1"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='tx2'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@tx2"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent1'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent1"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent2'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent2"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent3'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent3"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent4'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent4"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent5'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent5"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='accent6'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@accent6"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='hlink'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@hlink"/>
              </xsl:when>
              <xsl:when test="$schemeClrName='folHlink'">
                <xsl:value-of select="/p:presentation/p:sldMaster/p:clrMap/@folHlink"/>
              </xsl:when>
            </xsl:choose>
            <!--  </xsl:variable>
      母版背景  李秋玲 增加  2012.11.16
      
      <xsl:choose>
      <xsl:when test="not($schemeClr='#FFFFFF') or @idx=0 or @idx=1000">
        <xsl:value-of select="$schemeClr"/>
      </xsl:when>
       <xsl:otherwise>
           <xsl:variable name="idxx" select="@idx"/>
         <xsl:variable name="ID" select="ancestor::p:sldMaster/@id"/>
           <xsl:choose>
          <xsl:when test="$idxx &gt; 0 and $idxx &lt;1000" >
            <xsl:apply-templates select="//a:theme[@refBy=$ID]/a:themeElements/a:fmtScheme/a:fillStyleLst/*[position() = $idxx ]"/>
           </xsl:when>
           <xsl:when test="$idxx &gt;1000">
                <xsl:apply-templates select="//a:theme[@refBy=$ID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst/*[position()=($idxx - 1000)]"/>
           </xsl:when>
             <xsl:otherwise>
               <xsl:value-of select="$schemeClr"/>
             </xsl:otherwise>
           </xsl:choose>
        
          
          
          
        </xsl:otherwise>
          
        </xsl:choose>-->
          </图:颜色_8004>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--end  liuyangyang 2015-04-09 修改背景填充错误-->
  </xsl:template>
  <xsl:template name="getThemeBgFill">
    <!--从主题中获取背景填充-->
    <xsl:param name="idx"/>
    <xsl:param name="themeID"/>
    <xsl:param name="phClr"/>
    <xsl:for-each select="//a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst">
      <xsl:variable name="bgFill" select="child::*"/>
      <xsl:choose>
        <xsl:when test="$idx='1001'">
            <xsl:apply-templates select="$bgFill[1]">
              <xsl:with-param name="phClr" select="$phClr"/>
            </xsl:apply-templates>
        </xsl:when>
        <xsl:when test="$idx='1002'">
            <xsl:apply-templates select="$bgFill[2]">
              <xsl:with-param name="phClr" select="$phClr"/>
            </xsl:apply-templates>
        </xsl:when>
        <xsl:when test="$idx='1003'">
            <xsl:apply-templates select="$bgFill[3]">
            <xsl:with-param name="phClr" select="$phClr"/>
            </xsl:apply-templates>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="backgroundFill">
	  <演:背景_6B2C>
      <!--xsl:apply-templates select=".//p:bg[1]"/-->
      <xsl:apply-templates select="./p:cSld/p:bg"/>
		  <!--<xsl:apply-templates select="../p:sldMaster/p:cSld/p:bg"/>-->
    </演:背景_6B2C>
  </xsl:template>

  <!--项目编号的图片>
	<xsl:template name="bulletBlip">
		<xsl:if test="a:buBlip">
			<xsl:apply-templates select="a:buBlip"/>
		</xsl:if>
	</xsl:template-->
  <!--xsl:template match="a:buBlip" mode="autoNumbering">
		<字:图片符号引用 uof:locID="t0164" uof:attrList="宽度 高度">
			<xsl:call-template name="findRelatedRelationships"><xsl:with-param name="id"><xsl:value-of select=".//a:blip/@r:embed"/></xsl:with-param></xsl:call-template>
		</字:图片符号引用>
		
	</xsl:template-->
  <!--
  <xsl:template match="a:buBlip"  mode="pictureText">
    <字:自动编号信息 uof:locID="t0059" uof:attrList="编号引用 编号级别 重新编号 起始编号">
      <xsl:attribute name="字:编号引用">
        <xsl:call-template name="findRelatedRelationships">
          <xsl:with-param name="id">
            
            <xsl:if test="@id">
              <xsl:value-of select="concat('bn_',@id)"/>
            </xsl:if>
            <!-p/pPr/bu...的情况 -->
  <!--<xsl:if test="not(@id)">
              <xsl:value-of select="concat('bn_',translate(../@id,':',''))"/>
            </xsl:if>
            
            <!-xsl:value-of select=".//a:blip/@r:embed"/-->
  <!--</xsl:with-param>
        </xsl:call-template>
      </xsl:attribute>
      <!-编号级别默认为1-->
  <!--       <xsl:attribute name="字:编号级别">1</xsl:attribute>
    </字:自动编号信息>    
  </xsl:template>
  -->
  <!--测试项目编号是图片的情况
	<xsl:template match="/">
		<xsl:for-each select="p:presentation/p:sld">
			
				<xsl:apply-templates select="//..//a:pPr/a:buBlip"></xsl:apply-templates>
			
			</xsl:for-each>
		
	</xsl:template>
-->
  <!--2015.03.06,liuyangyang添加对透明度的计算模版-->
  <xsl:template name="transByAlpha">
    <xsl:param name="rgb"/>
    <xsl:param name="alpha"/>
    <xsl:param name="bgcolor"/>
    <xsl:choose>
      <xsl:when test="not ($bgcolor)">
        <xsl:value-of select="$rgb"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="tDecR">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($rgb,1,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="tDecG">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($rgb,3,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="tDecB">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($rgb,5,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="bDecR">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($bgcolor,2,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="bDecG">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($bgcolor,4,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="bDecB">
          <xsl:call-template name="hexToDecimal">
            <xsl:with-param name="hex" select="substring($bgcolor,6,2)"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="R">
          <xsl:call-template name="decimalToHex">
            <xsl:with-param name="result" select="$tDecR * (1 - $alpha) + $bDecR * $alpha"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="G">
          <xsl:call-template name="decimalToHex">
            <xsl:with-param name="result" select="$tDecG * (1 - $alpha) + $bDecG * $alpha"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="B">
          <xsl:call-template name="decimalToHex">
            <xsl:with-param name="result" select="$tDecB * (1 - $alpha) + $bDecB * $alpha"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="concat($R,$G,$B)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--luyangyang添加对颜色透明度的计算模版-->
  <!--2014.03.06,liuyangyang添加对rgb颜色调节hsl分量的计算模版-->
  <xsl:template name="transByHSL">
    <xsl:param name="rgb"/>
    <xsl:param name="pattern"/>
    <xsl:param name="attri"/>
    <xsl:param name="bgClr"/>
    <xsl:variable name="value">
      <xsl:choose>
        <xsl:when test="contains($attri , '%')">
          <xsl:value-of select="substring($attri,1,string-length($attri) - 1) div 100"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$attri div 100000"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$value = 0">
        <xsl:value-of select="$rgb"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$pattern = 'other'">
            <xsl:call-template name="transByAlpha">
              <xsl:with-param name="rgb" select="$rgb"/>
              <xsl:with-param name="alpha" select="$value"/>
              <xsl:with-param name="bgcolor" select="$bgClr"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="tDecR">
              <xsl:call-template name="hexToDecimal">
                <xsl:with-param name="hex" select="substring($rgb,1,2)"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="tDecG">
              <xsl:call-template name="hexToDecimal">
                <xsl:with-param name="hex" select="substring($rgb,3,2)"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="tDecB">
              <xsl:call-template name="hexToDecimal">
                <xsl:with-param name="hex" select="substring($rgb,5,2)"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="DecR">
              <xsl:value-of select="$tDecR div 255"/>
            </xsl:variable>
            <xsl:variable name="DecG">
              <xsl:value-of select="$tDecG div 255"/>
            </xsl:variable>
            <xsl:variable name="DecB">
              <xsl:value-of select="$tDecB div 255"/>
            </xsl:variable>
            <xsl:variable name="maxClrValue">
              <xsl:choose>
                <xsl:when test="$DecR &gt; $DecG and $DecR &gt; $DecB">
                  <xsl:value-of select="$DecR"/>
                </xsl:when>
                <xsl:when test="$DecG &gt; $DecR and $DecG &gt; $DecB">
                  <xsl:value-of select="$DecG"/>
                </xsl:when>
                <xsl:when test="$DecB &gt; $DecG and $DecB &gt; $DecR">
                  <xsl:value-of select="$DecB"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$DecR"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="minClrValue">
              <xsl:choose>
                <xsl:when test="$DecR &lt; $DecG and $DecR &lt; $DecB">
                  <xsl:value-of select="$DecR"/>
                </xsl:when>
                <xsl:when test="$DecG &lt; $DecR and $DecG &lt; $DecB">
                  <xsl:value-of select="$DecG"/>
                </xsl:when>
                <xsl:when test="$DecB &lt; $DecG and $DecB &lt; $DecR">
                  <xsl:value-of select="$DecB"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$DecR"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="L">
              <xsl:value-of select="($maxClrValue+$minClrValue) div 2"/>
            </xsl:variable>
            <xsl:variable name="max_min">
              <xsl:value-of select="($maxClrValue - $minClrValue)"/>
            </xsl:variable>
            <xsl:variable name="S">
              <xsl:choose>
                <xsl:when test="$L = 0 or $L = 1">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:when test="$L &lt; 0.5">
                  <xsl:value-of select="$max_min div ($maxClrValue+$minClrValue)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$max_min div (2.0 - $maxClrValue - $minClrValue)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="tempH">
              <xsl:choose>
                <xsl:when test="$max_min=0">
                  <xsl:value-of select="'0'"/>
                </xsl:when>
                <xsl:when test="$maxClrValue=$DecR">
                  <xsl:value-of select="($DecG - $DecB) div $max_min"/>
                </xsl:when>
                <xsl:when test="$maxClrValue=$DecG">
                  <xsl:value-of select="2.0 + ($DecB - $DecR) div $max_min"/>
                </xsl:when>
                <xsl:when test="$maxClrValue=$DecB">
                  <xsl:value-of select="4.0 + ($DecR - $DecG) div $max_min"/>
                </xsl:when>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="H">
              <xsl:choose>
                <xsl:when test="$tempH &lt; 0">
                  <xsl:value-of select="$tempH div 6 + 1"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$tempH div 6"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$pattern='hue'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$value"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='hueMod'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$value * $H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='hueOff'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$value + $H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='sat'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$value"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='satMod'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$value * $S"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='satOff'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$value + $S"/>
                  <xsl:with-param name="L" select="$L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='lum'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$value"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='lumMod'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$value * $L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='shade'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$value * $L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='lumOff'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="$value + $L"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="$pattern='tint'">
                <xsl:call-template name="HSLtoRGB">
                  <xsl:with-param name="H" select="$H"/>
                  <xsl:with-param name="S" select="$S"/>
                  <xsl:with-param name="L" select="1 - $value + $L * $value"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$rgb"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--end-->
</xsl:stylesheet>
