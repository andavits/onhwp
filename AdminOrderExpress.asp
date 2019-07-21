<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							AdminOrderExpress.asp
'======================================================================
'
' Filename: AdminOrderExpress.asp	
' Description:	Main page for administration of Order Express
'				catalogue. Provides the adiministrator with the 
'				following options:
'					Update "New Items" on main page
'					Database - new item, update and delete items
'					Review orders
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Macromedia Flash (optional)
'										Disclaimer.inc
'
'	Called From:	/orderexpress/admin/  -virtual directory redirects
'	Calls:			productnew.asp
'					productdelete.asp
'					productupdate.asp
'					vieworders.asp
'					adminnewitems.asp
'					www.newhome.on.ca/site-nav.htm
'
'	Version:	Version 1.0					Date: Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		Select option: Update New Items, Update Database, view orders
'	end
'======================================================================
%>
<% Response.Expires = 0 %>
<% REM Previous line is used to work around proxy caching problems. %> 

<% REM Set up error handling. %>
<% on error resume next %>

<HTML>
<HEAD>
	<TITLE>Order Express Administration</TITLE>
</HEAD>
<body bgcolor="#FFFFFF" topmargin="0">
<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=top alt="ONHWP logo">
<br>

<center>
<IMG SRC="/orderexpress/images/admin.gif" height = "92" width = "358" 
		border=0 alt="Return to Order Express Administration">
<H2 ALIGN=CENTER>Administration Page for Order Express</H2>

<table border="0" width="80%">
	<tr>
		<td>
			<OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
			WIDTH="85" HEIGHT="85" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
			<PARAM NAME="MOVIE" VALUE="/orderexpress/images/doghelpma.swf">

				<EMBED SRC="/orderexpress/images/doghelpma.swf" WIDTH="85" HEIGHT="85"
					PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
				</EMBED>
				<NOEMBED>
				<A HREF="/order/adminhelp.asp">
					<img src="/orderexpress/images/doghelp.gif" WIDTH="85" HEIGHT="85" ALT="Click here for help" border=0>
				</A>
				</NOEMBED>
			</OBJECT>
		</td>
		<td >
			This is the main page for Order Express 
			Administration. These "restricted access" pages enable  
			Order Express to be updated and maintained. 
			Please select one of the following options:
			<P>
		</td>
	</tr>
</table> 
 
<table border="0" width="80%"> 
	<tr>
		<td Align=center>
			<%' Table to present user options %>
			<table Cellpadding=10 Cellspacing=10>
				<tr BGCOLOR="f7efde">
					<td>
	   					<A HREF="/order/admin/adminnewitems.asp">
	   						<H3>Update "New Items" of Order Express main page.</H3>
	  					</a>
	   				</td>
				</tr>
				<tr BGCOLOR="f7efde">
					<td>
						<H3>Update database of products and services.</H3>
						<ul>
							<li><A HREF="/order/admin/productnew.asp">
									ADD NEW Product
								</a>
							<li><A HREF="/order/admin/productupdate.asp">
									UPDATE Existing Product
								</a>
							<li><A HREF="/order/admin/productdelete.asp">
									DELETE Existing Product
								</a>
						</ul>
					</td>
				</tr>
				<tr BGCOLOR="f7efde">
					<td>
						<A Href="/order/admin/vieworders.asp">
							<H3>Review Orders Received</H3> 
						</a>
					</td>
				</tr>
			</table>
			<%' End table of user options %>
		</td>
   </tr>
</table>

<%' navigation: link to Order Express main page, and ONHWP site navigation %>
<A HREF="/order/orderexpress.asp">Return to Order Express</a> , 
<A HREF="http://www.newhome.on.ca/site_nav.htm">Return to Site Navigation</a>
</center>

</BODY>
</HTML>