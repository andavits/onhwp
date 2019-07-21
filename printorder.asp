<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							printorder.asp
'======================================================================
'
' Filename: printorder.asp
' Description:	Prints orders that require payment to the screen
'				for the end user to print out and process via mail 
'				other method.
'				The final order (subtotal, shipping, builder discount
'				if applicable) is shown.
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Macromedia Flash (optional)
'										Disclaimer.inc
'
'	Called From:	CustomerInfo.asp
'					
'	Calls:			OrderExpress.asp
'					www.newhome.on.ca/site-nav.htm
'					www.newhomw.on.ca/homeowners/main_manor/main_manor.htm
'
'	Version:	Version 1.0					Date: Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		outputs customer info
'		summarizes order for printing
'		provides payment info and options
'	end
'======================================================================
%>

<%
Const MaxItems=30
Const PRODPRICE= 0
Const ITEMCHECKED = 1
Const PRODID = 2
Const PRODNAME =3
Const PRODQUANTITY = 4
Const PRODONEFREE = 5
Const PRODINSTOCK = 6

iCount = Session("ItemCount")
basket = Session("MyBasket")
total = Session("Total")
shipping = Session("Shipping")
builder = ""

' Determine the OrderID to display on the page:
oid=Request.QueryString("OrderID")
' If no OrderID has was sent or OrderID is not a numeric value
' then return user to main Order Express page.
If oid = "" OR IsNumeric(oid)=False Then
	Response.Redirect("/order/orderexpress.asp")
End If
%>

<HTML>
<HEAD>
	<TITLE>Order Express - Thank you</TITLE>
</HEAD>
<Body background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0">

<center>
<table width=100%>
	<tr>
		<td width= 152 valign=top>
			<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=top alt="ONHWP logo">
		</td>
		<td align=center>
			<center>
			<a href="/order/OrderExpress.asp">
				<img src="/orderexpress/images/orderexpresssm.gif" border=0 
				alt="Back to Main Page" width = 178 height = 61 align=top>
			</a>
			
			</center>
		</td>
	</tr>
	<tr>
		<td>
			<OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
				 WIDTH="75" HEIGHT="75" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
				<PARAM NAME="MOVIE" VALUE="/orderexpress/images/dogthanks.swf">

				<EMBED SRC="/orderexpress/images/dogthanks.swf" WIDTH="75" HEIGHT="75"
					 PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
				</EMBED>
				<NOEMBED>
				<A HREF="http://205.210.250.157/order/help.asp">
					<img src="/orderexpress/images/dogthankyou.gif" WIDTH="75" HEIGHT="73" Border=0 ALT="Thank You">
				</A>
				</NOEMBED>
			</OBJECT>		
		</td>
		<td valign=top>
			<h2 Align=center>Thank you for using Order Express!</h2>
			To complete your order you will need to print out a 
			copy of this page, enter in your payment details. 
			Then fax or mail your order to ONHWP at the
			address listed below.
			<br><b>Order reference number:  
			
			<%  =Request.QueryString("OrderID") %>
			</b>
		</td>
	</tr>
</table>
</center>

<%
'Obtain customer information based on OrderID
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT Customers.* " &_ 
	"FROM Customers, Orders " &_
	"WHERE Orders.OrderID =" & oid &_ 
	"AND Orders.CustomerID = Customers.CustomerID " 
Set rs = Conn.Execute(SQL)
'obtain builder number for later use
builder = rs("BuilderNumber")
%>

<% ' output customer info to screen %>
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
						Apt./Suite:
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
						Daytime Phone:
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
					<td BGCOLOR="f7efde">
						<%=rs("BuilderNumber")%>
					</td>
				</tr>
			</table>
		</td>
	</tr>

<%
' end of customer informatio output to screen
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
<center>

<%'Output Summary of Order in a table%>

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
For i = 1 to iCount
%>

	<TR>
		<TD BGCOLOR="f7efde" ALIGN=CENTER>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<%=basket(PRODNAME,i)%>
			</FONT>
			<% If basket(PRODINSTOCK,i) < 1 then %>
				<br>
				<FONT COLOR="#FF000" Size=1>
					Out of Stock
				</FONT>
			<% End If %>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=CENTER>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<%=basket(PRODQUANTITY,i)%>
			</FONT>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<%If basket(PRODONEFREE,i)="True" then %>
				<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
					1 copy FREE
				</FONT>
				<BR>
				<FONT SIZE=1>
					Additional copies<% = FormatCurrency(basket(PRODPRICE,i))%>
				</FONT>
			<%Else%>
				<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
					<% = FormatCurrency(basket(PRODPRICE,i))%>
				</FONT>
			<%End If%>
		</TD>
		<TD BGCOLOR="f7efde" ALIGN=RIGHT>
			<% cquantity = basket(PRODQUANTITY,i)
			If basket(PRODONEFREE,i)="True" then 
				cquantity=cquantity-1
			End If
			%>
			<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
				<% = FormatCurrency(basket(PRODPRICE,i) * cquantity)%>
			</FONT>
		</TD>
	</TR>

<%
	'Add product price to subtotal
	If (basket(PRODPRICE,i)) <> "" Then
		 iSubTotal = iSubtotal + (basket(PRODPRICE,i) * cquantity)
	End If

	'repeat for next item in order
	Next
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

<P>
GST#:R121546931
</center>

<P>
<b>Method of Payment</b>
<center>
<P>
<h3>
Cheques or Money Orders must be made payable to:<br>
	Ontario New Home Warranty Program</h3>
<table Border=1 >
	<tr>
		<td width=20 bgcolor="#FFFFFF">
			&nbsp;
		</td>
		<td>
			Money Order  (Canadian Funds)
		</td>
		<td Rowspan=3>
			<table>
				<tr>
					<td valign=top>
						Forward all orders to:
					</td>
					<td>
						Ontario New Home Warranty Program<br>
						5160 Yonge St., 6th Floor<br>
						North York, Ontario<br>
						Canada<br>
						M2N 6L9<br>
						<P>
						Fax: (for visa orders only)<br>
						416-229-3845
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width=20 bgcolor="#FFFFFF">
			&nbsp;
		</td>
		<td>
			Cheque (must be drawn on a Canadian Bank)
		</td>
	</tr>
	<tr>
		<td width=20 bgcolor="#FFFFFF">
			&nbsp;
		</td>
		<td>
			
			Visa <img src="/orderexpress/images/visa.gif" width="49" height="30" alt="VISA">
			(minimum order $10.00 CDN)
			<table>
				<tr>
					<td align=right>
						Name on Card:
					</td>
					<td width=200 >
						&nbsp;__________________________
					</td>
				</tr>
				<tr>
					<td align=right>
						Bank:
					</td>
					<td width=200 >
						&nbsp;__________________________
					</td>
				</tr>
				<tr>
					<td align=right>
						Card #:
					</td>
					<td width=200 >
						&nbsp;__________________________
					</td>
				</tr>
				<tr>     
 					<td align=right>
						Expiry Date:
					</td>
					<td width=200 >
						&nbsp;__________________________
					</td> 
				</tr>
				<tr>  
					<td align=right>
						Signature Line:
					</td>
					<td width=200 >
						<P>
						&nbsp;__________________________
					</td>  
				</tr>
			</table>

		</td>
	</tr>
</table>


<b>Payment MUST accompany order form.</b>
<P>
All prices are in effect as of:   September 1, 1997 
<P>
<a href="/order/orderexpress.asp">Return to Order Express</A> , 
<A HREF="http://www.newhome.on.ca/site_nav.htm">Return to Site Navigation</a>
<P>
<HR>
<h3>Explore what else this site has to offer!</h3>
<BR>
<OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
  WIDTH="450" HEIGHT="130" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
 <PARAM NAME="MOVIE" VALUE="/orderexpress/images/maintmanor2.swf">

    <EMBED SRC="/orderexpress/images/maintmanor2.swf" WIDTH="450" HEIGHT="130"
     PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
    </EMBED>
	<NOEMBED>
	<A HREF="http://www.newhome.on.ca/homeowners/main_manor/main_manor.htm>
	<img src="/orderexpress/images/maintmanor.gif" WIDTH="450" HEIGHT="130" ALT="Maintenance Manor" border=0>
	</A>
	</NOEMBED>
</OBJECT>
</center>
<P>
<!--#include virtual="/order/disclaimer.inc"-->

</BODY>
</HTML>