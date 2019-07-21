<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							ProductUpdate.asp
'======================================================================
'
' Filename:	ProductUpdate.asp
' Description:	Presents the Order Express Administrator a list of
'				all the products (Product ID and Name). The user
'				can select the product to update from this list.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):
'
'	Called From:	AdminOrderExpress.asp
'	Calls:			Update.asp
'					AdminOrderExpress.asp
'
'	Version:	Version 1.0						Date: Sept.4,1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		read and output all product ID's and Names
'		redirect user to update page with that item's details
'	end
'======================================================================
%>
<% Response.Expires = 0 %>
<% REM Previous line is used to work around proxy caching problems. %> 

<% REM Set up error handling. %>
<% on error resume next %>

<% 
'open database products table and output ProductID and Product Name
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT ProductID, ProductName FROM Products" 
Set rs = Conn.Execute(SQL)

'output page header
%>

<HTML>
<HEAD>
	<TITLE>Order Express Administration</TITLE>
</HEAD>
<body bgcolor="#FFFFFF" topmargin="0">
<table width=100%>
	<tr>
		<td width= 152 valign=top>
			<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=top alt="ONHWP logo">
		</td>
		<td align=center>
			<center>
			<A HREF="/order/admin/AdminOrderExpress.asp">
				<IMG SRC="/orderexpress/images/admin.gif" 
					height = "92" width = "358" border=0 
					alt="Return to Order Express Administration">
			</A>
			</center>
		</td>
	</tr>
</table>
<center>
<H2 ALIGN=CENTER>Order Express Product Update</H2>

<table border="0" width="80%">
	<tr>
		<td colspan=2>
			Select the Product to Update:
			<P>
		</td>
	</tr>
	<tr>
		<td>
			Product ID
		</td>
		<td>
			Product Name
		</td>
	</tr>
<%
'Loop through all the products and output name and ID
Do While Not rs.EOF 
%>
	<tr>
		<td>
			<A HREF="/order/admin/update.asp?ProductID=<%=rs("ProductID") %>"> 
				<%=rs("ProductID") %>
			</A>
	   	</td>
		<td>
			<A HREF="/order/admin/update.asp?ProductID=<%=rs("ProductID") %>"> 
				<%=rs("ProductName") %>
			</a>
		</td>
	</tr>

<%
	rs.MoveNext		'move to next item in recordset
	Loop
	rs.Close
	Conn.Close
%>

</table>
<P>
<A HREF="/order/admin/AdminOrderExpress.asp">Return to Order Express Administration Home</A>
</center>
<P>

</BODY>
</HTML>