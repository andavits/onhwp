<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							OrderExpress.asp
'======================================================================
'
' Filename:	OrderExpress.asp
' Description:	Order Express Main Page
'		Provides the option of viewing the various publications 
'		based on one of three Categories: Consumer, Builder/Vendor 
'		and Technical Information. The inset box presents information
'		about what new products.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Macromedia Flash (Optional)
'										Disclaimer.inc
'										newsitems.txt
'
'	Called From:	order_express.htm (on ONHWP Server)
'					Prodservcat.asp
'					Basket.asp
'	Calls:			Prodservcat.asp
'					www.newhome.on.ca/site_nav.htm
'					basket.asp
'
'	Version:	Version 1.0						Date: Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers:	RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		Read "New Items" information from text file
'		Output HTML and user menu options
'	end
'======================================================================
%>

<% Response.Expires = 0 %>
<% REM Previous line is used to work around proxy caching problems. %> 

<% REM Set up error handling. %>
<% on error resume next %>

<% 
' Output Main Page graphics
%>

<HTML>
<HEAD>
	<TITLE>Order Express</TITLE>
</HEAD>
<body  background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0">

<center>
<table border="0" width="100%">
	<tr>
        <td>
			<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=top alt="ONHWP logo"><br>
			<center><IMG src="/orderexpress/images/order3.GIF"
					width="523" height="95" alt="Order Express"></center>
		</td>
        <td>
			&nbsp;
		</td>
    </tr>
</table>

<table width=100% >
	<tr>
		<td valign=top>
<%
			' Flash Component - Dog holding "Grand Opening" sign
			' Installs ActiveX control, or Netscape plugin if supported
			' Displays a gif of the dog if components are not supported
%>
			<OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
				WIDTH="75" HEIGHT="75" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
				<PARAM NAME="MOVIE" VALUE="/orderexpress/images/dogonhwp2.swf">

				<EMBED SRC="/orderexpress/images/dogonhwp2.swf" WIDTH="75" HEIGHT="75"
					PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
				</EMBED>
				<NOEMBED>
					<img src="/orderexpress/images/dogonhwp2.gif" width=75 height=75 border=0 alt="Grand Opening">
				</NOEMBED>
			</OBJECT>

<%
			'Display a button for viewing the shopping basket if there are any items in it.
			 If Session("ItemCount") > 0 Then	
				'items in cart, therefore show button and link it to the shopping basket page
%>
				<P>
				<FORM ACTION="/order/basket.asp" METHOD=POST>
				<INPUT TYPE=SUBMIT NAME="Action" VALUE="Review Basket">
				</FORM>
			<% End If ' End of Shopping Basket Show %>
		</td>
	    <td align=center valign=top >
			<%' Display the "Menu" of catagories to choose from.  %>
			<img src="/orderexpress/images/menu.gif" width="60" height="28" alt="Menu">
			<br>
	  		<b>Of special interest to</b><br>
			<A HREF="/order/prodservcat.asp?Category=Consumer">
				<img src="/orderexpress/images/consumers.gif" border=0 width="88" 
					height="34" alt="Consumer Information">
			</A>
			<P>
			<b>Of special interest to</b><br>
			<A HREF="/order/prodservcat.asp?Category=Builder">
				<img src="/orderexpress/images/vendor.gif" border=0 width="94" 
					height="51" alt="Builder/Vendor Information">
			</A>
			<P>
			<b>Special interest:</b><br>
			<A HREF="/order/prodservcat.asp?Category=Technical">
				<img src="/orderexpress/images/technical.gif" border=0 width="98" 
					height="35" alt="Technical Information">
			</A>
			<P>
		</td>
		<td align=left> 
			Welcome to Order Express! 
			The Ontario New Home Warranty Program's online catalogue of 
			publications and videos designed for new home builders, 
			consumers and industry professionals in Ontario. 
			<P>
			Order Express lets you easily find materials suited to your 
			needs and develop your own customized "shopping list". And 
			even better, many of our publications are free of charge. 
			<P>
			Celebrating 21 years of service to Ontario new home builders, 
			consumers and industry professionals! 
			<P>
			If at any time you require help in using Order Express, 
			click on the dog icon.
		</td>
		<td rowspan=10 width=1 bgcolor="#000000">
			<%' insert a verical black line %>
			<img width=1 height=1 SRC="/orderexpress/images/line.gif" ALT="">
		</td>
		<td align=right valign=top>
			<!-- new items start--->
			<%' insert table (box) containing information about new items %>	
			<table width=180 cellpadding=2 bgcolor="#f7efde" >
				<tr>
					<td colspan=2 >
						<font color="#DD0022" size=+1><b>New Products available in 1997:</b></font>
					</td>
				</tr>

<% 
				' Open Text Stream for Input. 
				' File "newsitems.txt" contains details about new items
				Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
				Set Instream = FileObject.OpenTextFile (Server.MapPath ("/orderexpress") & "\newitems.txt", 1, FALSE, FALSE)

				'Read each "New Item" (one  per line)
				Do while Instream.AtEndOfStream <> True
					newsstring = Instream.ReadLine
					'parse line of text to determine graphic, title and text
					ss=";"
					mynum = InStr(newsstring,ss)
					'determine graphic to display: "new", "revised" or "comingsoon"
					newsgraphic = Left(newsstring,mynum -1)	
					'create graphic file from previous word	
					imageurl = "/orderexpress/images/" & newsgraphic & ".gif"  
					'determine what title and text for new item
					tmpstring = Right(newsstring,Len(newsstring)-mynum) 
					mynum = InStr(tmpstring,ss)
					titlestring = Left(tmpstring,mynum-1)
					txtstring=Right(tmpstring,Len(tmpstring)-mynum) 
					'Output the graphic, title and text in a row in the table
%>
				<tr>
					<td valign=top>
						<img src="<%=imageurl %>">
					</td>
					<td>
						<font><b><%=titlestring %></B></font><br><%=txtstring %><P>
					</td>
				</tr>
<%
				Loop		'loop till no more "new items" in text file
				Instream.Close		'close text stream
				'Close "New Items" table
%>
	 
			</table>   
			<!---- news items end --->
		</td>
		<td width=50>
			&nbsp;
		</td>		
    </tr>
</table>

<P>

Email your comments to:<A HREF="mailto: orderexp@newhome.on.ca">orderexp@newhome.on.ca</A>
<P>
<A HREF="http://www.newhome.on.ca/site_nav.htm">Return to Site Navigation</A>
</center>
<P>
<!--#include virtual="/order/disclaimer.inc"-->

</BODY>
</HTML>