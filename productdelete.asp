<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							productdelete.asp
'======================================================================
'
' Filename:	productdelete.asp
' Description:		Used to present a list of items from the database
'					that the administrator can select the desired 
'					product from.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):
'
'	Called From:	AdminOrderExpress.asp
'	Calls:			delete.asp
'					AdminOrderExpress.asp
'
'	Version:	Version 10					Date: Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers:	 RV AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		presents all the products, select which on to view and delete
'	end
'======================================================================
%>
<% Response.Expires = 0 %>
<% 'Previous line is used to work around proxy caching problems. %> 

<% 'Set up error handling. %>
<% on error resume next %>

<% 'This page is for updating the Product and Services database. %>
<% 'open table and output prod ID and Product name %>
<% 
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT ProductID, ProductName FROM Products" 

Set rs = Conn.Execute(SQL)
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
<H2 ALIGN=CENTER>Order Express Product Delete</H2>
<table border="0" width="80%">
    <tr>
	<td colspan=2>Select the Product to Delete:
	<P>
	</td>
   </tr>
	<tr>
	<td>Product ID</td>
	<td>Product Name</td>
	</tr>
<% Do While Not rs.EOF %>
<tr>
	<td>
	<A HREF="/order/admin/delete.asp?ProductID=<%=rs("ProductID") %>">
	<%=rs("ProductID") %>
	</A>
   	</td>
	<A HREF="/order/admin/delete.asp?ProductID=<%=rs("ProductID") %>">
	<td><%=rs("ProductName") %>
	</a>
	</td>
</tr>

<%
	rs.MoveNext
Loop
rs.Close
Conn.Close
%>
 
   
</table>
</center>

<A HREF="/order/admin/AdminOrderExpress.asp">Return to Order Express Administration Home</A>
<P>

</BODY>
</HTML>