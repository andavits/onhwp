<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							vieworders.asp
'======================================================================
'
' Filename:		vieworders.asp
' Description:	Presents a list of both no-charge orders and orders
'				that require payment, on the screen. The user selects 
'				the order based on the data (or OrderID).
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):
'
'	Called From:	AdminOrderExpress.asp
'	Calls:			view.asp 
'					AdminOrderExpress.asp
'	Version:	Version 1.0						Date: Sept.4,1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers:	RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		present user with a list of all the no-charge orders
'		present user with a list of all orders requiring payment
'	end
'======================================================================
%>
<% Response.Expires = 0 %>
<% 'Previous line is used to work around proxy caching problems. %> 

<% 'Set up error handling. %>
<% on error resume next %>

<HTML>
<HEAD>
<TITLE>Order Express Administration View Orders</TITLE>
</HEAD>
<body bgcolor="#FFFFFF" topmargin="0">

<center>
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

<H2 ALIGN=CENTER>Review Orders Received</H2>

<%
' open connection to database and retreive no-charge orders 
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")

' generate SQL query for no charge orders
SQL = "SELECT Orders.* FROM Orders "
SQL = SQL & "WHERE Orders.OrderTotal = 0 AND Orders.FreightCharge = 0" 

'run the query and output the orders
Set rs = Conn.Execute(SQL)	
%>
<H2 ALIGN=CENTER>No-Charge Orders</H2>
<table border="0" width="80%">
    <tr>
		<td colspan=2>
			Select the Order to View:
		</td>
	</tr>
	<tr>
		<td>
			<b>Order ID</b>
		</td>
		<td>
			<b>Order Date and Time</b>
		</td>
	</tr>
<% Do While Not rs.EOF %>
	<tr>
		<td>
			<A HREF="/order/admin/view.asp?OrderID=<%=rs("OrderID") %>">
				<%=rs("OrderID") %>
			</A>
   		</td>
		<td>
			<A HREF="/order/admin/view.asp?OrderID=<%=rs("OrderID") %>">
				<%=rs("OrderDate") %>
			</A>
		</td>
	</tr>
<%
	rs.MoveNext
	Loop		'loop through the result set
	rs.Close
	Conn.Close	'close data connection
%>
</table>

<% 
' open connection to database and retreive orders that require payment
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")

'generate query for orders that require payment
SQL = "SELECT Orders.* FROM Orders "
SQL = SQL & "WHERE Orders.OrderTotal > 0 " 

'execute the query 
Set rs = Conn.Execute(SQL)
'display  a list of orders
%>
<H2 ALIGN=CENTER>Orders That Require Payment</H2>
<table border="0" width="80%">
    <tr>
		<td colspan=2>
			Select the Order to View:
		</td>
	</tr>
	<tr>
		<td>
			<b>Order ID</b>
		</td>
		<td>
			<b>Order Date and Time</b>
		</td>
	</tr>
<% Do While Not rs.EOF %>
	<tr>
		<td>
			<A HREF="/order/admin/view.asp?OrderID=<%=rs("OrderID") %>">
				<%=rs("OrderID") %>
			</A>
   		</td>
		<td>
			<A HREF="/order/admin/view.asp?OrderID=<%=rs("OrderID") %>">
				<%=rs("OrderDate") %>
			</A>
		</td>
	</tr>
<%
	rs.MoveNext
	Loop		'loop through the resultset
	rs.Close
	Conn.Close	'close the data connection
%>
</table>


<P>
<A HREF="/order/admin/AdminOrderExpress.asp">Return to Order Express Administration </a>
</center>
<P>

</BODY>
</HTML>