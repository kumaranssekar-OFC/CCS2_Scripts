<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">

<xsl:template match="SW_Overview">
  <fo:root xmlns:fo="http://www.w3.org/1999/XSL/Format">
    <fo:layout-master-set>
      <fo:simple-page-master page-height="210mm" page-width="297mm"
                             margin="5mm 5mm 5mm 5mm" master-name="PageMaster_landscape">
       <fo:region-body margin="5mm 5mm 5mm 5mm"/>
      </fo:simple-page-master>
      <fo:simple-page-master page-height="297mm" page-width="210mm"
                             margin="5mm 5mm 5mm 5mm" master-name="PageMaster">
       <fo:region-body margin="5mm 5mm 5mm 5mm"/>
      </fo:simple-page-master>
    </fo:layout-master-set>
    <fo:page-sequence master-reference="PageMaster">
      <fo:flow flow-name="xsl-region-body" >
        <fo:block>
          <xsl:apply-templates select="DocInfo"/>
        </fo:block>
        <fo:block font-size="24pt" font-weight="bold" text-align="center" line-height="60pt">
          SW Overview for Production
        </fo:block>
        <fo:block font-size="16pt" color="red" text-align="center">
          <xsl:for-each select="BoschSPLNrSet/BoschSPLNr">
            <xsl:sort select="." order="ascending"></xsl:sort>
            <xsl:value-of select="."/><fo:block/>
          </xsl:for-each>
        </fo:block>
        <fo:block font-size="16pt" text-align="center" line-height="60pt">for</fo:block>
        <fo:block font-size="16pt" text-align="center">
          <xsl:for-each select="BoschPrjName">
            <xsl:value-of select="."/> devices<fo:block/>
          </xsl:for-each>
        </fo:block>
        <fo:block font-size="16pt" color="red" text-align="center">
          <xsl:for-each select="BoschPrjNameNr">
            <xsl:sort select="." order="ascending"></xsl:sort>
            <xsl:value-of select="."/><fo:block/>
          </xsl:for-each>
        </fo:block>
        <fo:block>
          <xsl:apply-templates select="ProjectInfo"/>
        </fo:block>
      </fo:flow>
    </fo:page-sequence>
    <fo:page-sequence master-reference="PageMaster_landscape">
      <fo:flow flow-name="xsl-region-body" >
        <fo:block hyphenate="true" language="en">
          <xsl:apply-templates select="SW_Versions"/>
          <xsl:apply-templates select="Set_Versions"/>
          <xsl:apply-templates select="Set_Versions_Partnumbers"/>
          <xsl:apply-templates select="Set_Versions_Products"/>
        </fo:block>
          <xsl:apply-templates select="eCSD"/>
      </fo:flow>
    </fo:page-sequence>
  </fo:root>
</xsl:template>

<xsl:template match="eCSD">
  <fo:block font-size="16pt" text-align="left" line-height="60pt">eCSD Settings:</fo:block>
  <fo:block>
    <fo:external-graphic src="eCSD_V01.00.png"/>
  </fo:block>
</xsl:template>

<xsl:template match="DocInfo"> 
   <fo:table table-layout="fixed" border-color="black" border-style="solid" border-width="0px">
   <fo:table-column column-width="19cm" border-color="black" border-style="solid" border-width="0px"/>
   <fo:table-body>
    <fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          Version: <xsl:value-of select="@Doc_Version"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
    <fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          Date: <xsl:value-of select="@Doc_Date"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
    <fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          State: <xsl:value-of select="@Doc_State"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
    <fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          <xsl:value-of select="@Doc_Access"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
    <fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          <xsl:value-of select="@Doc_SetDef"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
	<fo:table-row>
      <fo:table-cell>
        <fo:block font-size="9pt" text-align="right">
          <xsl:value-of select="@Doc_Scope"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
   </fo:table>
</xsl:template>

<xsl:template match="ProjectInfo"> 
   <fo:table table-layout="fixed" width="190mm" border-color="black" border-style="solid" border-width="0px">
   <fo:table-column column-width="2cm" border-color="black" border-style="solid" border-width="1px"/>
   <fo:table-column column-width="4cm" border-color="black" border-style="solid" border-width="1px"/>
   <fo:table-column column-width="1cm" border-color="black" border-style="solid" border-width="0px"/>
   <fo:table-column column-width="2cm" border-color="black" border-style="solid" border-width="1px"/>
   <fo:table-column column-width="4cm" border-color="black" border-style="solid" border-width="1px"/>
   <fo:table-body>
    <fo:table-row>
      <fo:table-cell padding="3pt">
        <fo:block font-weight="bold" font-size="12pt" text-align="start">
          Overall SW Version:
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="3pt">
        <fo:block color="red" font-size="12pt" text-align="start">
          <xsl:value-of select="@SW_Version"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell>
      </fo:table-cell>
      <fo:table-cell padding="3pt">
        <fo:block font-weight="bold" font-size="12pt" text-align="start">
          SW-ID:
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="3pt">
        <fo:block color="red" font-size="12pt" text-align="start">
          <xsl:value-of select="@SW_ID"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
    <fo:table-row>
      <fo:table-cell padding="3pt">
        <fo:block font-weight="bold" font-size="12pt" text-align="start">
          SW-PM:
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="3pt">
        <fo:block font-size="12pt" text-align="start">
          <xsl:value-of select="@SW_PM"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell>
      </fo:table-cell>
      <fo:table-cell padding="3pt">
        <fo:block font-weight="bold" font-size="12pt" text-align="start">
          Change Number:
         </fo:block>
      </fo:table-cell>
      <fo:table-cell>
        <fo:block font-size="12pt" text-align="start">
          <xsl:value-of select="@CHNG_NR"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
      <xsl:apply-templates/>
   </fo:table>
</xsl:template>

<xsl:template match="SW_Versions"> 
   <fo:table table-layout="fixed" border-color="black" border-style="solid" border-width="1px" break-after="page">
   <fo:table-column column-width="0.8cm"/>
   <fo:table-column column-width="2.2cm"/>
   <fo:table-column column-width="3.3cm"/>
   <fo:table-column column-width="3.7cm"/>
   <fo:table-column column-width="7.5cm"/>
   <fo:table-column column-width="3.6cm"/>
   <fo:table-column column-width="3.5cm"/>
   <fo:table-column column-width="1.8cm"/>
   <fo:table-column column-width="1.5cm"/>
   <fo:table-body>
    <fo:table-row keep-together.within.page="always">
      <fo:table-cell number-columns-spanned="9" padding="2pt" background-color="#B6B6B4" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="center">
          SW Versions
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
      <xsl:apply-templates/>
   </fo:table>
</xsl:template>

<xsl:template match="SW_Table_Header">
  <fo:table-body>
    <fo:table-row background-color="#B6B6B4" keep-together.within.page="always">
      <fo:table-cell number-columns-spanned="2" padding="2pt" background-color="#B6B6B4"  border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="7pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col7"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col8"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col9"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="SW_Table_Header1">
  <fo:table-body>
    <fo:table-row keep-together.within.page="always">
      <fo:table-cell number-columns-spanned="9" padding="2pt" background-color="#B6B6B4" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Product_Info">
  <fo:table-body keep-together.within.page="1">
    <fo:table-row keep-together.within.page="always">
      <fo:table-cell number-columns-spanned="2" padding="2pt"  border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="9pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="7pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col7"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col8"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col9"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Product_Info_Col8|Product_Info_SplitCol1">
  <fo:table-body keep-together.within.page="1">
    <fo:table-row keep-together.within.page="always">
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="9pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="9pt" text-align="start">
          <xsl:value-of select="@Col2"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="7pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col7"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col8"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col9"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="SW_Comments">
  <fo:table-body>
    <fo:table-row>
      <fo:table-cell number-columns-spanned="9" padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          Comments:
          <xsl:apply-templates/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Comment">
  <fo:block>
    <xsl:apply-templates/>
  </fo:block>
</xsl:template>

<xsl:template match="Set_Versions">
   <!-- <fo:block break-before="page"></fo:block> not implemented in fop 0.20.5 -->
   <fo:table table-layout="fixed" border-color="black" border-style="solid" border-width="1px">
   <fo:table-column column-width="3.2cm"/>
   <fo:table-column column-width="2.8cm"/>
   <fo:table-column column-width="3cm"/>
   <fo:table-column column-width="3.7cm"/>
   <fo:table-column column-width="2cm"/>
   <fo:table-column column-width="2cm"/>
   <fo:table-column column-width="2.4cm"/>
   <fo:table-column column-width="1.5cm"/>
   <fo:table-column column-width="1.1cm"/>
   <fo:table-column column-width="1.1cm"/>
   <fo:table-column column-width="1.2cm"/>
   <fo:table-column column-width="2.6cm"/>
   <fo:table-body keep-together="always">
    <fo:table-row>
      <fo:table-cell number-columns-spanned="12" padding="2pt" background-color="#B6B6B4" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="center">
          Set Information
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
      <xsl:apply-templates/>
   </fo:table>
</xsl:template>

<xsl:template match="Set_Versions_Partnumbers">
   <!-- <fo:block break-before="page"></fo:block> not implemented in fop 0.20.5 -->
  <fo:table table-layout="fixed" border-color="black" border-style="solid" border-width="1px">
  <fo:table-column column-width="5cm"/>
  <fo:table-column column-width="7cm"/>
  <fo:table-column column-width="7cm"/>
 
  <fo:table-body keep-together="always">
    <fo:table-row>
      <fo:table-cell number-columns-spanned="4" padding="2pt" background-color="#B6B6B4" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="center">
          Set Information (Partnumber Info)
        </fo:block>
      </fo:table-cell>
    </fo:table-row>
  </fo:table-body>
  <fo:table-body keep-together="always" hyphenate="false">
    <fo:table-row background-color="#B6B6B4">
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          SW Set (Container)
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          SAP Productnumber(s)
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          SPL Partnumber
         </fo:block>
      </fo:table-cell>
       
      </fo:table-row>
  </fo:table-body>
    <xsl:apply-templates/>
  </fo:table>
</xsl:template>

<xsl:template match="Set_Versions_Products">
  <fo:table table-layout="fixed" border-color="black" border-style="solid" border-width="1px">
  <fo:table-column column-width="3cm"/>
  <fo:table-column column-width="3.7cm"/>
  <fo:table-column column-width="3.7cm"/>
  <fo:table-column column-width="1.7cm"/>
  <fo:table-column column-width="3.5cm"/>
  <fo:table-column column-width="1.2cm"/>
  >
  <fo:table-body>
    <fo:table-row>
      <fo:table-cell number-columns-spanned="13" padding="2pt" background-color="#B6B6B4" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="center">
          Set Information (Variant Info)
        </fo:block>
      </fo:table-cell>
    </fo:table-row>
  </fo:table-body>
  <fo:table-body keep-together="always" hyphenate="false">
    <fo:table-row background-color="#B6B6B4">
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          SW Set (Container)
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          CTS
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          TEST_MANAGER
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
           UBLOX
          </fo:block>
        </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
           SXM
          </fo:block>
        </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
           DTV
          </fo:block>
        </fo:table-cell>
      </fo:table-row>
  </fo:table-body>
    <xsl:apply-templates/>
  </fo:table>
</xsl:template>

<xsl:template match="Set_Table_Header">
  <fo:table-body keep-together="always" hyphenate="false">
    <fo:table-row background-color="#B6B6B4">
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col2"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="10pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
     
    
      
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Set_Info">
  <fo:table-body keep-together="always">
    <fo:table-row>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col2"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
     
    
     
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Set_Info_Col4">
  <fo:table-body keep-together="always">
    <fo:table-row>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col2"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
    </fo:table-row>
   </fo:table-body>
</xsl:template>

<xsl:template match="Set_Info_Col13">
  <fo:table-body keep-together="always">
    <fo:table-row>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col1"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col2"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col3"/>
        </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col4"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col5"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col6"/>
         </fo:block>
      </fo:table-cell>
      <fo:table-cell padding="2pt" border-color="black" border-style="solid" border-width="1px">
        <fo:block font-size="8pt" text-align="start">
          <xsl:value-of select="@Col7"/>
         </fo:block>
      </fo:table-cell>
   
     
    </fo:table-row>
   </fo:table-body>
</xsl:template>

</xsl:stylesheet >