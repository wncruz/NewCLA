<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >

  <xsl:template match="/">
    <HTML>
      <BODY>
        <TABLE>
          <xsl:for-each select="customers/customer">
            <xsl:sort select="state" order="descending"/>
            <xsl:sort select="name"/>
            <TR>
              <TD><xsl:value-of select="name" /></TD>
              <TD><xsl:value-of select="address" /></TD>
              <TD><xsl:value-of select="phone" /></TD>
            </TR>
          </xsl:for-each>
        </TABLE>
      </BODY>
    </HTML>
  </xsl:template>

</xsl:stylesheet>


<?xml version="1.0"?>
<?xml-stylesheet type="text/xsl" href="attribute.xsl"?>
<investment>
   <type>stock</type>
   <name>Microsoft</name>
   <price type="high">100</price>
   <price type="low">94</price>
</investment>


<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" indent="yes"/>

<xsl:template match="investment">
   <xsl:element name="{type}">
      <xsl:attribute name="name" >
         <xsl:value-of select="name"/>
      </xsl:attribute>
      <xsl:for-each select="price">
      <xsl:attribute name="{@type}" >
         <xsl:value-of select="."/>
      </xsl:attribute>
      </xsl:for-each>
   </xsl:element>
</xsl:template>

</xsl:stylesheet>
