<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							view.asp
'======================================================================
'
' Filename:	View.asp	
' Description:	Presents order details based on an OrderID that 
'				is sent to the file as a query string 
'				from the file vieworders.asp
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	
'
'	Called From:	vieworders.asp
'	Calls:			AdminOrderExpress.asp
'
'	Version:	Version 1.0					Date: Sept.4.1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		obtain orderID from vieworders.asp
'		present order information to screen
'	end
'======================================================================
%>
<%
' Determine the OrderID to display on the page:
oid=Request.QueryString("OrderID")
' If no OrderID has was sent or OrderID is not a numeric value
' then return user to main Order Express page.
If oid = "" OR IsNumeric(oid)=False Then
	Response.Redirect("/order/admin/vieworders.asp")
End If
%>
<HTML>
<HEAD>
	<TITLE>Order Express - View Order</TITLE>
</HEAD>
<Body bgcolor="#FFFFFF" topmargin="0">

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


<h2 align=center>View Order</h2>

<P align=center>
	<b>Order reference number:
	<%=Request.QueryString("OrderID")%>
	</b>
</P>


<%
'connect to database
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
'generate SQL query
SQL = "SELECT Customers.* " &_ 
	"FROM Customers, Orders " &_
	"WHERE Orders.OrderID =" & oid &_ 
	"AND Orders.CustomerID = Customers.CustomerID " 
Set rs = Conn.Execute(SQL)
'execute the query and output the information to the screen
%>
<table>
	<tr>
		<td Align=top>
   			<b>Order Information:</B>
		</td>
		<td>
		&nbsp;
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
		<td>
			<table >
				<tr>
					<td align=right>
						Title:
					</td>
					<td width=250 BGCOLOR="f7efde">
						<%=rs("CustomerTitle")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						First Name:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("ContactFirstName")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Last Name:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("ContactLastName")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						CompanyName:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("CompanyName")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Address:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("BillingAddress")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Apt:/Suite:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("Suite")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						City:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("City")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Province/State:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("StateOrProvince")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Postal Code:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("PostalCode")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Country:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("Country")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Phone:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("PhoneNumber")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Ext.:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("Extension")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Email:
					</td>
					<td BGCOLOR="f7efde">
						<%=rs("EmailAddress")%>
					</td>
				</tr>
				<tr>
					<td align=right>
						Registered Builder Number:
					</td>
					<td BGCOLOR="f7efde">\
						<%=rs("BuilderNumber")%>
					</td>
				</tr>
			</table>
		</td>
	</tr>


<%
'close record set and data connection
rs.Close
Conn.Close
%>

	<tr>
		<td>
			<b>Order Details</b>
		</td>
		<td>
			&nbsp;
		</td>
	</tr>
</table>

<%
'open connection to data source 
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")

'create query for order details
SQL = "SELECT OrderDetails.Quantity, Products.* " &_ 
	"FROM OrderDetails, Products " &_
	"WHERE OrderDetails.OrderID =" & oid  &_
	"AND Products.ProductID=OrderDetails.ProductID"

'generate result set and output results
Set rs = Conn.Execute(SQL)
%>

<% 'Output Summary of Order in a table %>
	
<center>
<Table COLSPAN=8 CELLPADDING=5 BORDER=1>
<!-- BEGIN column header row -->
	<TR>
		<TD ALIGN=CENTER BGCOLOR="#800000">
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>
				Product Name
			</FONT>
		</TD>
		<TD ALIGN=CENTER BGCOLOR="#800000">
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>
				Quantity
			</FONT>
		</TD>
		<TD ALIGN=CENTER BGCOLOR="#800000">
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>
				Consumer Price
			</FONT>
		</TD>
		<TD ALIGN=CENTER BGCOLOR="#800000">
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>
				Unit Total
			</FONT>
		</TD>
	</TR>

<%
' Loop through all items in the order an output each in a row
iSubtotal = 0
Do While Not rs.EOF  %>



	<TR>
		<TD BGCOLOR="f7efde" ALIGN=CENTER>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<%=rs("ProductName") %>
			</FONT>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=CENTER>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<%=rs("Quantity") %>
			</FONT>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<%If rs("OneFree")="True" then %>
				<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
					1 copy FREE
				</FONT>
				<BR>
				<FONT SIZE=1>
					Additional copies<% = FormatCurrency(rs("ConsumerPrice"))%>
				</FONT>
			<%Else%>
				<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
					<% = FormatCurrency(rs("ConsumerPrice"))%>
				</FONT>
			<%End If%>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<% cquantity = rs("Quantity")
			If rs("OneFree")="True" then 
				cquantity=cquantity-1
			End If
			%>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(rs("ConsumerPrice") * cquantity)%>
			</FONT>
		</TD>
	</TR>

<%
	'Add product price to subtotal
	If rs("ConsumerPrice") <> "" Then
		 iSubTotal = iSubtotal + (rs("ConsumerPrice") * cquantity)
	End If

	'repeat for next item in order
		rs.MoveNext
	Loop
%>

	<TR>
		<TD COLSPAN=2>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=LEFT>
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#800000" SIZE=2>
				Subtotal:
			</FONT>
		</TD>
		<% 
		if iSubTotal=0 then %>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				FREE
			</FONT>
		</TD>
		<% Else%>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(iSubtotal)%>
			</FONT>
		</TD>
		<% End If %>
	</TR>
	<TR>
		<TD COLSPAN=2>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=LEFT>
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#FF0000" SIZE=2>
				Registered Builder Discount:
			</FONT>
		</TD>
		<%'Calc builder discount
		'determine if builder, discount = isubtotal - ordertotal
		'else discount = 0
		If builder <> "" then
			discount = iSubtotal - total
		%>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(discount)%>
			</FONT>
		</TD>
		<% Else
			discount=0 %>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(discount)%>
			</FONT>
		</TD>
		<% End If %>
	</TR>
	<TR>
		<TD COLSPAN=2>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=LEFT>
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#800000" SIZE=2>
				Shipping:
			</FONT>
		</TD>
		<% ' enter shipping info %>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(shipping)%>
			</FONT>
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=LEFT>
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#800000" SIZE=2>
				ORDER TOTAL:
			</FONT>
		</TD>
		<% 'Calculate order total %>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#800000" SIZE=2>
				<% = FormatCurrency(iSubtotal - discount + shipping)%>
			</FONT>
		</TD>
	</TR>
</TABLE>
	<%

	rs.Close	'close record set and data connection
	Conn.Close
	%>


<P>
&nbsp;
<P>
<A HREF="AdminOrderExpress.asp">[Return to Order Express Administration Page]</A>
<A HREF="vieworders.asp">[Return to Review Orders Page]</A>
</center>

</BODY>
</HTML>