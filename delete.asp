<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							delete.asp
'======================================================================
'
' Filename:	delete.asp
' Description:	This page presents the selected product information
'				to the administrator. The administrator can verify
'				the delete process or cancel to return to the
'				main admin page.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	disclaimer.inc
'
'	Called From:	productdelete.asp
'	Calls:			delete.asp (recursive)
'					AdminOrderExpress
'
'	Version:	Version 1						Date:Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		get productID for product to delete
'		read and present prod info	
'		confirm delete product or cancel
'	end
'======================================================================
%>
<% Response.Expires = 0 %>
<% REM Previous line is used to work around proxy caching problems. %> 

<% 
'This page requires that a product ID be passed to it.
'If no ProductID, then redirect user to Amin page
Product = Request.QueryString("ProductID")
If Product = "" OR IsNumeric(Product)=False Then
	Response.Redirect("/order/admin/productdelete.asp")
End If
%>

<% REM Set up error handling. %>
<% on error resume next %>

<%
'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively. The following section 
'	is used to check the information submited in the form.
'	 The first time this form is displayed, the following section
'	is not processed because Request("Action") = "" (user has 
'	not pressed the Delete Product button, named "Action", 
'	at the bottom of the page). Pressing the enter button, 
'	sets "Action" = "Delete Product" or "Cancel"
'----------------------------------------------------------------------

SELECT CASE Request("Action")
   
	CASE "Delete Product"
	
		'delete item
		'create SQL statement to delete item	
		sql = "DELETE FROM Products WHERE ProductID=" & CLng(Product)
		'connect to database
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
				
		rs.Close
		Conn.Close

		'once item deleted, redirect user to main Admin page
		Response.Redirect "/order/admin/adminOrderExpress.asp"

CASE "Cancel"

	'do not delete item
	'redirect user to main Admin page
	Response.Redirect "/order/admin/AdminOrderExpress.asp"

END SELECT
'----------------------------------------------------------------------
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
<H2 ALIGN=CENTER>Order Express Product DELETE</H2>

<h1><font color="#FF0000">
	Confirm that you wish to delete the following product.
</font>
</h1>

<% 
'present all the product information to the user
' so they can verify that they will be deleting the correct item

'connect to database and get all the product information for requested product
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT *,* FROM Products, ProdType WHERE Products.ProductID=" & CLng(Product)
SQL = SQL & " AND Products.ProductType = ProdType.ProdTypeID "
Set rs = Conn.Execute(SQL)

'output the product info
%>

<table border="0" width="80%">
	<TR>
		<TD align=right>
			Product ID
		</TD>
		<TD>
			<%=rs("ProductID")%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Product Name
		</TD>
		<TD>
			<%=rs("ProductName")%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Product Despcription
		</TD>
		<TD>
			<%=rs("ProductDescription")%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Consumer Price
		</TD>
		<TD>
			<%=FormatCurrency(rs("ConsumerPrice"))%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Builder Price
		</TD>
		<TD>
			<%=FormatCurrency(rs("BuilderPrice"))%>
		</TD>
	</TR>
		<TR>
		<TD align=right>
			Product Type
		</TD>
		<TD>
			<%=rs("ProdTypeName")%>
		</TD>
	</TR>

	<TR>
		<TD align=right>
			Product Image Filename
		</TD>
		<TD>
			<%=rs("ProductImage")%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Product Flag
		</TD>
		<TD>
			<%=rs("ProductFlag")%>
		</TD>
	</TR>
	<TR>
		<TD align=right>
			Product Preview URL
		</TD>
		<TD>
			<%=rs("ProductPreviewURL")%>
		</TD>
	</TR>
</TABLE>


<FORM ACTION="/order/admin/delete.asp?ProductID=<%=rs("ProductID")%>" METHOD=POST>

	<INPUT TYPE=SUBMIT NAME="Action" VALUE="Delete Product">
	<INPUT TYPE=SUBMIT NAME="Action" VALUE="Cancel"> 

</Form>


<%
'close database connection
rs.Close
Conn.Close
%>


</center>
<P>

</BODY>
</HTML>