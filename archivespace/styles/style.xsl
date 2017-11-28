<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:HTML="http://www.w3.org/Profiles/XHTML-transitional">

<!--xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="http://www.w3.org/1999/xhtml"-->

<!-- File last modified: 2003-07-27 (dwr) -->
	<xsl:template match="/">
		<HTML>
			<HEAD><TITLE>
					<xsl:value-of select="ead/frontmatter/titlepage/titleproper"/>
				</TITLE>
				<LINK rel="stylesheet" type="text/css" name="RMCstyle" href="../styles/rmc.css" />


			</HEAD>
			<BODY>
				<xsl:apply-templates select="ead"/>
			</BODY>
		</HTML>
	</xsl:template>


<!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- remove any elements marked for internal audience only -->
	<xsl:template match="//*[@audience='internal']" priority="1"/> 

<!-- EAD -->
	<xsl:template match="ead">
		<xsl:apply-templates/>
	</xsl:template>

<!-- EADHEADER -->
	<xsl:template match="eadheader"/>

<!-- FRONTMATTER -->
	<xsl:template match="frontmatter">
		<xsl:apply-templates/>
	</xsl:template>

<!-- TITLEPAGE -->
	<xsl:template match="titlepage">
		<xsl:apply-templates/>
	</xsl:template>
	<xsl:template match="titleproper">
		<H1>
			<xsl:apply-templates/>
		</H1>
	</xsl:template>
	
<!-- NUM -->
	<xsl:template match="num">
		<H2>
			<xsl:apply-templates/>
		</H2>
	</xsl:template>

<!-- PUBLISHER -->
	<xsl:template match="publisher">
		<H4>
			<xsl:apply-templates/>
		</H4>
	</xsl:template>

<!-- TITLEPAGE/DATE -->
	<xsl:template match="titlepage/date">
		<H5>
			<P><xsl:apply-templates/></P>
		</H5>
	</xsl:template>
	
<!-- TITLEPAGE/LISTS -->
	<xsl:template match="titlepage/list">
	<CENTER><TABLE CELLPADDING="10">
		<TR VALIGN="top">
			<xsl:apply-templates />
		</TR></TABLE></CENTER>
	</xsl:template>
	
<!-- DEFITEM LIST (AUTHOR INFO/RESPOSITORY ADDRESS)-->
	<xsl:template match="titlepage/list/defitem">
		<TD>
			<xsl:apply-templates />
		</TD>
	</xsl:template>
	<xsl:template match="titlepage/list/defitem/label">
		<DIV class="smalllabel">
			<xsl:apply-templates />
		</DIV>
	</xsl:template>
	<xsl:template match="titlepage/list/defitem/item">
		<DIV class="smallitem">
			<xsl:apply-templates />
		</DIV>
	</xsl:template>
	
<!-- DIV/PREFACE -->
	<xsl:template match="div">
		<A NAME="aa1">
			<xsl:apply-templates/>
		</A>
	</xsl:template>
	
	
<!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ARCHDESC -->
	<xsl:template match="archdesc">
		<xsl:apply-templates/>
	</xsl:template>


<!-- HEAD -->
	<xsl:template match="head">
		<DIV class="H4">
			<xsl:apply-templates/>
		</DIV>
		<P></P>
	</xsl:template>


<!-- LABEL -->
	<xsl:template match="@label">
		<DIV class="heading">
			<xsl:value-of select="." />
		</DIV>
	</xsl:template>


<!-- HIGH-LEVEL DID/DESCRIPTIVE SUMMARY -->
	<xsl:template match="archdesc/did">
		<HR/>
		<A NAME="a1">
			<xsl:apply-templates/>
		</A>
		<P></P>
	</xsl:template>
	
	<xsl:template match="archdesc/did/physdesc | archdesc/did/unittitle | archdesc/did/unitid | repository | archdesc/did/abstract | origination | archdesc/did/langmaterial">
		<xsl:apply-templates select="@label"/>
		<DIV class="item">
			<xsl:apply-templates/>
		</DIV>
	</xsl:template>
	
	<xsl:template match="unittitle/unitdate | origination/persname | origination/corpname">
		<xsl:apply-templates/>
	</xsl:template>


<!-- BIOGHIST/BIOGRAPHICAL NOTE -->
	<xsl:template match="bioghist">
		<HR/>
		<A NAME="a2">
			<head></head>
			<xsl:apply-templates/>
		</A>

	</xsl:template>

	<!--This template rule formats a chronlist element.-->
	<xsl:template match="chronlist">
		<table width="100%" style="margin-left:25pt">
			<tr>
				<td width="5%"> </td>
				<td width="15%"> </td>
				<td width="80%"> </td>
			</tr>
			<xsl:apply-templates/>
		</table>
	</xsl:template>
	
	<xsl:template match="chronlist/head">
		<tr>
			<td colspan="3">
				<h4>
					<xsl:apply-templates/>
				</h4>
			</td>
		</tr>
	</xsl:template>
	
	<xsl:template match="chronlist/listhead">
		<tr>
			<td> </td>
			<td>
				<b>
					<xsl:apply-templates select="head01"/>
				</b>
			</td>
			<td>
				<b>
					<xsl:apply-templates select="head02"/>
				</b>
			</td>
		</tr>
	</xsl:template>
	
	<xsl:template match="chronitem">
		<!--Determine if there are event groups.-->
		<xsl:choose>
			<xsl:when test="eventgrp">
				<!--Put the date and first event on the first line.-->
				<tr>
					<td> </td>
					<td valign="top">
						<xsl:apply-templates select="date"/>
					</td>
					<td valign="top">
						<xsl:apply-templates select="eventgrp/event[position()=1]"/>
					</td>
				</tr>
				<!--Put each successive event on another line.-->
				<xsl:for-each select="eventgrp/event[not(position()=1)]">
					<tr>
						<td> </td>
						<td> </td>
						<td valign="top">
							<xsl:apply-templates select="."/>
						</td>
					</tr>
				</xsl:for-each>
			</xsl:when>
			<!--Put the date and event on a single line.-->
			<xsl:otherwise>
				<tr>
					<td> </td>
					<td valign="top">
						<xsl:apply-templates select="date"/>
					</td>
					<td valign="top">
						<xsl:apply-templates select="event"/>
					</td>
				</tr>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>


<!-- SCOPECONTENT/COLLECTION DESCRIPTION -->
	<xsl:template match="archdesc/scopecontent">
		<HR/>
		<A NAME="a3">
			<xsl:apply-templates/>
		</A>

		
		
<!-- CUSTODHIST, ACQIINFO / HISTORY OF OWNERSHIP, PROVENANCE  (appear after SCOPECONTENT) -->		
		<xsl:choose>
			<xsl:when test="//custodhist">
				<HR/>
				<A NAME="a4"><DIV class="H4">
					<xsl:value-of select="//custodhist/head"/>
				</DIV></A>
				<P></P>
					<xsl:apply-templates select="//custodhist/p"/>							

			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<xsl:when test="//acqinfo">
				<HR/>
				<A NAME="a5"><DIV class="H4">
					<xsl:value-of select="//acqinfo/head"/>
				</DIV></A>
				<P></P>
					<xsl:apply-templates select="//acqinfo/p"/>

			</xsl:when>
		</xsl:choose>
	</xsl:template>


<!-- CONTROLACCESS/SUBJECTS -->
	<xsl:template match="controlaccess">
		<HR/>
		<A NAME="a6">
			<xsl:apply-templates/>
		</A>

	</xsl:template>
	
	<xsl:template match="controlaccess/controlaccess">
		<DIV class="heading">
			<xsl:value-of select="head"/>
		</DIV>
			<xsl:apply-templates select="subject | persname | corpname | geogname | genreform | occupation | title"/>
		<P></P>
	</xsl:template>

	<xsl:template match="subject | controlaccess/persname | controlaccess/corpname | geogname | genreform | occupation">
		<DIV class="item">
			<xsl:value-of select="." />
		</DIV>
	</xsl:template>

	<xsl:template match="controlaccess/title">
		<DIV class="item">
			<I><xsl:value-of select="." /></I>
		</DIV>
	</xsl:template>


<!-- ADMININFO/INFORMATION FOR USERS -->
	<xsl:template match="descgrp">
	<xsl:choose>
		<xsl:when test="accessrestrict | userestrict | altformavail | processinfo | prefercite | otherfindaid">
			<HR/>
			<A NAME="a7">
				<xsl:apply-templates/>
			</A>

		</xsl:when>
	</xsl:choose>
	</xsl:template>

	<xsl:template match="descgrp/custodhist | descgrp/acqinfo"/>

	<xsl:template match="descgrp/accessrestrict | descgrp/userestrict | descgrp/altformavail | descgrp/processinfo | descgrp/prefercite | descgrp/otherfindaid">
		<DIV class="heading">
			<xsl:choose>
				<xsl:when test="@id">
					<a>
						<xsl:attribute name="name">#<xsl:value-of select="@id"/></xsl:attribute><xsl:value-of select="head"/>
					</a>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="head"/>
				</xsl:otherwise>
			</xsl:choose>
		</DIV>
		<DIV class="item">
			<xsl:apply-templates select="p"/>
		</DIV>

	</xsl:template>


<!-- RELATED MATERIAL -->
	<xsl:template match="archdesc/relatedmaterial | archdesc/*/relatedmaterial">
<hr></hr>		
		<A NAME="a8">
			<xsl:apply-templates/>
		</A>

	</xsl:template>

	<!-- NOTES -->
	<xsl:template match="archdesc/odd | archdesc/*/odd">
		<hr></hr>		
		<A NAME="a8">
			<xsl:apply-templates/>
		</A>
		
	</xsl:template>

<!--	<xsl:template match="archdesc/relatedmaterial | archdesc/*/relatedmaterial">
		<xsl:choose>
			<xsl:when test="@id">
				<hr></hr>
				<a>
					<xsl:attribute name="name">#<xsl:value-of select="@id"/></xsl:attribute><xsl:apply-templates/>
				</a>
			</xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates/>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>-->
	
	<!-- relatedmaterial lists handled below under add at end -->

	
<!-- COLLECTION ARRANGEMENT/SERIES LIST -->
	<xsl:template match="archdesc/arrangement">
		<hr/>
		<A NAME="a9">
			<xsl:apply-templates/>
		</A>
	</xsl:template>
	
	<xsl:template match="archdesc/arrangement/list | archdesc/arrangement/list/defitem/item | defitem/item/list | item/list/defitem/item | list/item/list">
		<xsl:apply-templates/>
	</xsl:template>
	
	<xsl:template match="archdesc/arrangement/list/defitem">
		<xsl:choose>
			<xsl:when test="item/list/defitem">
				<TR><TD VALIGN="top">
					<xsl:apply-templates select="label"/>
				</TD></TR>	
				<TR><TD>
					<xsl:apply-templates select="item"/>
				</TD></TR>
			</xsl:when>
			<xsl:otherwise>	
				<TR><TD VALIGN="top">
					<xsl:apply-templates select="label"/>
				</TD>
				<TD ALIGN="RIGHT">
					<xsl:apply-templates select="item"/>
				</TD></TR>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="defitem/item/list/defitem">
		<xsl:choose>
			<xsl:when test="item/list/defitem">
				<TR><TD><DIV><xsl:attribute name = "STYLE">
						margin-left: 2em;
					</xsl:attribute>
					<xsl:apply-templates select="label"/>
				</DIV></TD></TR>	
				<TR><TD>
					<xsl:apply-templates select="item"/>
				</TD></TR>
			</xsl:when>
			<xsl:when test="item/list/item">
					<xsl:apply-templates />
			</xsl:when>
			<xsl:otherwise>	
				<TR><TD><DIV class="hlabel"><xsl:attribute name = "STYLE">
						margin-left: 2em;
						</xsl:attribute>
						<xsl:apply-templates select="label"/></DIV></TD>
				<TD NOWRAP="1" ALIGN="RIGHT"><xsl:apply-templates select="item"/></TD></TR>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="item/list/item ">
		<TR><TD><DIV><xsl:attribute name = "STYLE">
				margin-left: 2em;
			</xsl:attribute>
				<xsl:apply-templates />
		</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="item/list/defitem ">
		<TR><TD><DIV class="hlabel"><xsl:attribute name = "STYLE">
				margin-left: 2em;
			</xsl:attribute>
				<xsl:apply-templates select="label"/></DIV></TD>
		<TD NOWRAP="1" ALIGN="RIGHT"><xsl:apply-templates select="item"/></TD></TR>
	</xsl:template>


<!-- DSC/CONTAINER LIST -->
	<xsl:template match="dsc">
		<hr></hr>
		<A NAME="a10">
			<xsl:apply-templates select="head"/>
		</A>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TABLE CELLSPACING="10">
					<TR>
						<TD NOWRAP="1" ALIGN="left">
							<DIV class="heading">Date</DIV>
						</TD>
						<TD>
							<DIV class="heading">Description</DIV>
						</TD>
						<TD NOWRAP="1" ALIGN="center" COLSPAN="2">
							<DIV class="heading">Container</DIV>
						</TD>
					</TR>
					<xsl:apply-templates select="c01"/>
				</TABLE>
			</xsl:when>
			<xsl:otherwise>
				<TABLE CELLSPACING="5">
					<TR>
						<TD>
							<DIV class="heading">Description</DIV>
						</TD>
						<TD NOWRAP="1" ALIGN="center" COLSPAN="2">
							<DIV class="heading">Container</DIV>
						</TD>
					</TR>
					<xsl:apply-templates select="c01"/>
				</TABLE>
			</xsl:otherwise>
		</xsl:choose>

	</xsl:template>
	
	<xsl:template match="c01|c02|c03|c04|c05|c06">
		<xsl:apply-templates/>
	</xsl:template>
	
	<xsl:template match="c01/did|c02/did|c03/did|c04/did|c05/did|c06/did">
	
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
		<xsl:choose>

			<xsl:when test="../@level='series'">
				<xsl:if test="//did/unitdate">
					<TD><xsl:apply-templates select="unitdate"/></TD>
				</xsl:if>
				<TD>
					<DIV class="serieslabel">
						<xsl:apply-templates select="unittitle"/>
					</DIV>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="center">
					<xsl:choose>
						<xsl:when test="container[@type='box']">
							Box <xsl:apply-templates select="container[@type='box']"/>
						</xsl:when>
					</xsl:choose>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="center">
					<xsl:choose>
						<xsl:when test="container[@type='folder']">
							Folder <xsl:apply-templates select="container[@type='folder']"/>
						</xsl:when>
						<xsl:when test="unitid">
							<xsl:value-of select="unitid/@label" /> <xsl:apply-templates select="unitid"/>
						</xsl:when>
					</xsl:choose>
				</TD>
					<xsl:apply-templates select="abstract"/>	 
			</xsl:when>
			
			<xsl:when test="../@level='subseries'">
				<xsl:if test="//did/unitdate">
					<TD><xsl:apply-templates select="unitdate"/></TD>
				</xsl:if>
				<TD>
					<DIV class="subserieslabel">
					<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -2em;</xsl:attribute>
						<xsl:apply-templates select="unittitle"/>
					</DIV>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="TOP">
					<xsl:choose>
						<xsl:when test="container[@type='box']">
							Box <xsl:apply-templates select="container[@type='box']"/>
						</xsl:when>
					</xsl:choose>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="TOP">
					<xsl:choose>
						<xsl:when test="container[@type='folder']">
							Folder <xsl:apply-templates select="container[@type='folder']"/>
						</xsl:when>
						<xsl:when test="unitid">
							<xsl:value-of select="unitid/@label" /> <xsl:apply-templates select="unitid"/>
						</xsl:when>
					</xsl:choose>
				</TD>
					<xsl:apply-templates select="abstract"/>	 
			</xsl:when>
			
			<xsl:otherwise>	
				<xsl:if test="//did/unitdate">
					<TD NOWRAP="1" ALIGN="LEFT" VALIGN="TOP">
						<xsl:apply-templates select="unitdate"/>
					</TD>
				</xsl:if>
				<TD>
					<DIV>
					<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -2em;</xsl:attribute>
						<xsl:apply-templates select="unittitle"/>
						<xsl:if test="unitid and container[@type='folder']">
							<BR/><xsl:value-of select="unitid/@label" /> <xsl:apply-templates select="unitid"/>
						</xsl:if>
					</DIV>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="TOP">
					<xsl:choose>
						<xsl:when test="container[@type='box']">
							Box <xsl:apply-templates select="container[@type='box']"/>
						</xsl:when>
						<xsl:when test="container[@type='map-case']">
							Mapcase <xsl:apply-templates select="container[@type='map-case']"/>
						</xsl:when>
						<xsl:when test="container[@type='reel']">
							Reel <xsl:apply-templates select="container[@type='reel']"/>
						</xsl:when>
						<xsl:when test="container[@type='file-drawer']">
							File Drawer <xsl:apply-templates select="container[@type='file-drawer']"/>
						</xsl:when>
					</xsl:choose>
				</TD>
				<TD NOWRAP="1" ALIGN="CENTER" VALIGN="TOP">
					<xsl:choose>
						<xsl:when test="container[@type='folder']">
							Folder <xsl:apply-templates select="container[@type='folder']"/>
						</xsl:when>
						<xsl:when test="unitid">
							<xsl:value-of select="unitid/@label" /> <xsl:apply-templates select="unitid"/>
						</xsl:when>
						<xsl:when test="container[@type='volume']">
							Volume <xsl:apply-templates select="container[@type='volume']"/>
						</xsl:when>
						<xsl:when test="container[@type='frame']">
							Frame <xsl:apply-templates select="container[@type='frame']"/>
						</xsl:when>
					</xsl:choose>
				</TD>
					<xsl:apply-templates select="abstract"/>
					<xsl:apply-templates select="origination"/>
					<xsl:apply-templates select="physdesc"/>
				<xsl:apply-templates select="physloc"/>
				</xsl:otherwise>

	</xsl:choose>
	</TR>
	</xsl:template>
	
	<xsl:template match="c01/scopecontent/p | c02/scopecontent/p | c03/scopecontent/p | c04/scopecontent/p | c05/scopecontent/p | c06/scopecontent/p | c07/scopecontent/p">
	
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TD/>
			</xsl:when>
		</xsl:choose>
		<TD><DIV>
		<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
			<xsl:apply-templates/>
		</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/originalsloc/p | c02/originalsloc/p | c03/originalsloc/p | c04/originalsloc/p | c05/originalsloc/p | c06/originalsloc/p | c07/originalsloc/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/accessrestrict/p | c02/accessrestrict/p | c03/accessrestrict/p | c04/accessrestrict/p | c05/accessrestrict/p | c06/accessrestrict/p | c07/accessrestrict">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/userestrict/p | c02/userestrict/p | c03/userestrict/p | c04/userestrict/p | c05/userestrict/p | c06/userestrict/p | c07/userestrict/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/altformavail/p | c02/altformavail/p | c03/altformavail/p | c04/altformavail/p | c05/altformavail/p | c06/altformavail/p | c07/altformavail/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/relatedmaterial/p | c02/relatedmaterial/p | c03/relatedmaterial/p | c04/relatedmaterial/p | c05/relatedmaterial/p | c06/relatedmaterial/p | c07/relatedmaterial/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR> 
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				Related Materials: <xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/separatedmaterial/p | c02/separatedmaterial/p | c03/separatedmaterial/p | c04/separatedmaterial/p | c05/separatedmaterial/p | c06/separatedmaterial/p | c07/separatedmaterial/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR> 
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="c01/arrangement/p | c02/arrangement/p | c03/arrangement/p | c04/arrangement/p | c05/arrangement/p | c06/arrangement/p | c07/arrangement/p">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR> 
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD><DIV>
				<xsl:attribute name = "STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; </xsl:attribute>
				<xsl:apply-templates/>
			</DIV></TD></TR>
	</xsl:template>
	
	<xsl:template match="container[@type='box'] | container[@type='map-case'] | container[@type='folder']">
		<xsl:apply-templates/>
	</xsl:template>
	
	<xsl:template match="c01/did/unitid | c02/did/unitid | c03/did/unitid | c04/did/unitid | c05/did/unitid | c06/did/unitid">
			<xsl:apply-templates/>
	</xsl:template>
	
	<xsl:template match="did/unitdate">
		<xsl:apply-templates/>
	</xsl:template>
	
	<xsl:template match="c01/did/unittitle | c02/did/unittitle | c03/did/unittitle | c04/did/unittitle | c05/did/unittitle | c06/did/unittitle">
		<xsl:choose>
			<xsl:when test="@id">
				<a>
					<xsl:attribute name="name"><xsl:value-of select="@id"/></xsl:attribute><xsl:apply-templates/>
				</a>
			</xsl:when>
			<xsl:otherwise>
				<xsl:apply-templates/>	
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	
	<xsl:template match="c01/did/abstract | c02/did/abstract | c03/did/abstract | c04/did/abstract | c05/did/abstract | c06/did/abstract">
	
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TD/>
			</xsl:when>
		</xsl:choose>
		<TD>
			<DIV>
				<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -1em;</xsl:attribute>
					<xsl:value-of select="@label"/>
					<xsl:text> </xsl:text>
					<xsl:apply-templates/>
			</DIV>
		</TD>
		</TR>
	</xsl:template>
	
	<xsl:template match="c01/did/origination | c02/did/origination | c03/did/origination | c04/did/origination | c05/did/origination | c06/did/origination">

		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TD/>
			</xsl:when>
		</xsl:choose>
		<TD>
			<DIV>
				<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -1em;</xsl:attribute>
					<xsl:value-of select="@label"/>
					<xsl:text> </xsl:text>
					<xsl:apply-templates/>
			</DIV>
		</TD>
		</TR>
	</xsl:template>

	<xsl:template match="c01/did/physdesc | c02/did/physdesc | c03/did/physdesc | c04/did/physdesc | c05/did/physdesc | c06/did/physdesc">

		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TD/>
			</xsl:when>
		</xsl:choose>
		<TD>
			<DIV>
				<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -1em;</xsl:attribute>
					<xsl:value-of select="@label"/>
					<xsl:text> </xsl:text>
					<xsl:apply-templates/>
			</DIV>
		</TD>
		</TR>
	</xsl:template>
	
	<xsl:template match="c01/did/physloc | c02/did/physloc | c03/did/physloc | c04/did/physloc | c05/did/physloc | c06/did/physloc">
		
		<xsl:variable name="indent-value">
			<xsl:call-template name="depth-of-node" />
		</xsl:variable>
		
		<TR>
			<xsl:choose>
				<xsl:when test="//did/unitdate">
					<TD/>
				</xsl:when>
			</xsl:choose>
			<TD>
				<DIV>
					<xsl:attribute name="STYLE">margin-left: <xsl:value-of select="$indent-value - 3" />em; text-indent: -1em;</xsl:attribute>
					<xsl:value-of select="@label"/>
					<xsl:text> </xsl:text>
					<xsl:apply-templates/>
				</DIV>
			</TD>
		</TR>
	</xsl:template>
	
	<xsl:template match="c02/odd">
		<TR>
		<xsl:choose>
			<xsl:when test="//did/unitdate">
				<TD/>
			</xsl:when>
		</xsl:choose>
		<TD>
			<BLOCKQUOTE>
				<xsl:apply-templates select="list"/>
			</BLOCKQUOTE>
		</TD></TR>
	</xsl:template>

<!--
 ADDITIONAL INFORMATION 
 this is for generic add elements (type='addinfo') 

	<xsl:template match="add[@type='addinfo']">
		<xsl:apply-templates />
		<HR/>
	</xsl:template>
	
	<xsl:template match="add/add">
		<DIV class="heading">
			<xsl:value-of select="head" />
		</DIV>
		<DIV class="item">
			<xsl:apply-templates select="list | p" />
		</DIV>
	</xsl:template>
		-->
	
<!-- SEPARATED MATERIAL -->

		<xsl:template match="archdesc/descgrp/separatedmaterial">
		<xsl:apply-templates />
	</xsl:template>
	
	<xsl:template match="archdesc/separatedmaterial">
		<HR/>
		<xsl:apply-templates />
	</xsl:template>

<!--INDEX-->
	<xsl:template match="archdesc/index
		| archdesc/*/index">
		<table width="100%">
			<tr>
				<td width="5%"> </td>
				<td width="45%"> </td>
				<td width="50%"> </td>
			</tr>
			<tr>
				<td colspan="3">
					<h3>
						<a name="{generate-id(head)}">
							<b>
								<xsl:apply-templates select="head"/>
							</b>
						</a>
					</h3>
				</td>
			</tr>
			<xsl:for-each select="p | note/p">
				<tr>
					<td></td>
					<td colspan="2">
						<xsl:apply-templates/>
					</td>
				</tr>
			</xsl:for-each>
			
			<!--Processes each index entry.-->
			<xsl:for-each select="indexentry">
				
				<!--Sorts each entry term.-->
				<xsl:sort select="corpname | famname | function | genreform | geogname | name | occupation | persname | subject | title"/>
				<tr>
					<td></td>
					<td>
						<xsl:apply-templates select="corpname | famname | function | genreform | geogname | name | occupation | persname | subject | title"/>
					</td>
					<!--Supplies whitespace and punctuation if there is a pointer
						group with multiple entries.-->
					
					<xsl:choose>
						<xsl:when test="ptrgrp">
							<td>
								<xsl:for-each select="ptrgrp">
									<xsl:for-each select="ref | ptr">
										<xsl:apply-templates/>
										<xsl:if test="preceding-sibling::ref or preceding-sibling::ptr">
											<xsl:text>, </xsl:text>
										</xsl:if>
									</xsl:for-each>
								</xsl:for-each>
							</td>
						</xsl:when>
						<!--If there is no pointer group, process each reference or pointer.-->
						<xsl:otherwise>
							<td>
								<xsl:for-each select="ref | ptr">
									<xsl:apply-templates/>
								</xsl:for-each>
							</td>
						</xsl:otherwise>
					</xsl:choose>
				</tr>
				<!--Closes the indexentry.-->
			</xsl:for-each>
		</table>
		<p>
			<a href="#">Return to the Table of Contents</a>
		</p>
		<hr></hr>
	</xsl:template>

<!-- OTHER FINDING AIDS/GUIDES TO THIS MATERIAL -->
<!-- if lists are needed here (outside of p tags, -->
<!-- need to add otherfindaid to list templates below -->

<!--	<xsl:template match="descgrp/otherfindaid">
		<xsl:apply-templates />
	</xsl:template>
	
	<xsl:template match="otherfindaid">
		<xsl:apply-templates />
		<HR/>
	</xsl:template>-->


 <!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

 <!-- CERTAIN LISTS -->
 
	<xsl:template match="p/list | descgrp/list | arrangement/list | odd/list | separatedmaterial/list | relatedmaterial/list | scopecontent/list">
		<xsl:apply-templates select="head" />
		<xsl:choose>
			<xsl:when test="@type='marked'">
				<UL>
					<xsl:apply-templates select="item"/>
				</UL>
			</xsl:when>
			<xsl:when test="@type='simple'">
				<UL class="simple">
					<xsl:apply-templates select="item"/>
				</UL>
			</xsl:when>
			<xsl:when test="@type='ordered'">
				<xsl:choose>
					<xsl:when test="@numeration='arabic'">
						<OL>
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:when>
					<xsl:when test="@numeration='upperalpha'">
						<OL class="upperalpha">
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:when>
					<xsl:when test="@numeration='loweralpha'">
						<OL class="loweralpha">
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:when>
					<xsl:when test="@numeration='upperroman'">
						<OL class="upperroman">
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:when>
					<xsl:when test="@numeration='lowerroman'">
						<OL class="lowerroman">
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:when>
					<xsl:otherwise>
						<OL>
							<xsl:apply-templates select="item"/>
						</OL>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="@type='deflist'">
				<TABLE>
					<xsl:apply-templates select="defitem"/>
				</TABLE>
			</xsl:when>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="p/list/head | descgrp/list/head | arrangement/list/head | odd/list/head | separatedmaterial/list/head | relatedmaterial/list/head | scopecontent/list/head">
		<U>
			<xsl:apply-templates/>
		</U>
	</xsl:template>
		
	<xsl:template match="p/list/item | descgrp/list/item | arrangement/list/item | odd/list/item | separatedmaterial/list/item | relatedmaterial/list/item | scopecontent/list/item">
		<LI>
			<xsl:apply-templates/>
		</LI>
	</xsl:template>

	<xsl:template match="p/list/defitem | descgrp/list/defitem | arrangement/list/defitem | odd/list/defitem | separatedmaterial/list/defitem | relatedmaterial/list/defitem | scopecontentl/list/defitem">
		<TR><TD>
			<xsl:apply-templates select="label"/>
		</TD>
		<TD>
			<xsl:apply-templates select="item"/>
		</TD></TR>
	</xsl:template>


 <!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
 
 <!-- LOW LEVEL ELEMENTS -->
	<xsl:template match="lb">
		<xsl:apply-templates/>
		<BR/>
	</xsl:template>
	
	<xsl:template match="p">
		<DIV class="para"><xsl:apply-templates/></DIV></xsl:template>
		
	<xsl:template match="note">
			<DIV class="note"><xsl:value-of select="." /></DIV>
	</xsl:template>
	
	<xsl:template match="title">
		<I><xsl:apply-templates/></I></xsl:template>

	<xsl:template match="emph[@render='bold']">
		<B><xsl:apply-templates/></B></xsl:template>
	<xsl:template match="emph[@render='italic']">
		<I><xsl:apply-templates/></I></xsl:template>
	<xsl:template match="emph[@render='underline']">
		<U><xsl:apply-templates/></U></xsl:template>
	<xsl:template match="emph[@render='super']">
		<SUP><xsl:apply-templates/></SUP></xsl:template>
	<xsl:template match="emph[@altrender='red']">
		<FONT COLOR="red"><xsl:apply-templates/></FONT></xsl:template>
		

<!-- LINKS AND POINTERS -->

	<xsl:template match="extref">
		<xsl:choose>
			<xsl:when test="@show='new'">
<A target="_blank"><xsl:attribute name="HREF"><xsl:value-of select="@href"/></xsl:attribute><xsl:apply-templates/></A></xsl:when>
			<xsl:otherwise>
<A target="_top"><xsl:attribute name="HREF"><xsl:value-of select="@href"/></xsl:attribute><xsl:apply-templates/></A></xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="ref">
		<a href="#{@target}">
			<xsl:apply-templates/>
		</a>
	</xsl:template>

	<xsl:template match="extptr">
		<DIV CLASS="image">
	  		<IMG>
	  			<xsl:attribute name="SRC"><xsl:value-of select="@href"/></xsl:attribute>
	  			<xsl:attribute name="ALT"><xsl:value-of select="@title"/></xsl:attribute>
	  		</IMG>
	  	</DIV>
	  	<DIV CLASS="caption"><xsl:value-of select="@title"/></DIV>
	</xsl:template>


 <!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
 
 <!-- functions -->
 
<xsl:template name="depth-of-node">
 	<xsl:value-of select="count(ancestor::node())" />
</xsl:template>


<!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
  
</xsl:stylesheet>
