<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							thankyou.asp
'======================================================================
'
' Filename: thankyou.asp
' Description:  For Free orders. Presents the order summary to the user
'				and thanks them for their order.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Disclaimer.inc
'										Macromedia Flash (optional)
'
'	Called From:	CustomerInfo.asp
'	Calls:			www.newhome.on.ca/site-nav.htm
'					www.newhome.on.ca/newhome_shoppers/info-home
'
'	Version:	Version 1.0						Date: Sept 4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		retreive customer information from database
'		retreive order information from database
'		present to screen
'	end
'======================================================================
%>

<%
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
	<TITLE>Order Express - Thankyou</TITLE>
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
			<h2 Align=center>Thank you for your Order!</h2>
			Your order has been automatically sent to Order Express 
			for processing. ONHWP will ship your order within 
			five business days. 
		</td>
	</tr>
</table>
</center>



<% set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT Customers.* " &_ 
	"FROM Customers, Orders " &_
	"WHERE Orders.OrderID =" & oid &_ 
	"AND Orders.CustomerID = Customers.CustomerID " 
Set rs = Conn.Execute(SQL)

%>
<% REM output customer info to screen %>
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
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Title:
						</FONT>
					</td>
					<td width=250 BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("CustomerTitle")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						First Name:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("ContactFirstName")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Last Name:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("ContactLastName")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						CompanyName:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("CompanyName")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Address:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("BillingAddress")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Apt./Suite:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("Suite")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						City:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("City")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Province/State:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("StateOrProvince")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Postal Code:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("PostalCode")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Country:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("Country")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Daytime Phone:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("PhoneNumber")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Ext.:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("Extension")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Email:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("EmailAddress")%>
						</FONT>
					</td>
				</tr>
				<tr>
					<td align=right>
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						Registered Builder Number:
						</FONT>
					</td>
					<td BGCOLOR="f7efde">
						<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<%=rs("BuilderNumber")%>
						</FONT>
					</td>
				</tr>
	</table>
	</font>
   </td>
</tr>


<%
rs.Close
Conn.Close
%>

<% set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT OrderDetails.Quantity, Products.ProductName " &_ 
	"FROM OrderDetails, Products " &_
	"WHERE OrderDetails.OrderID =" & oid  &_
	"AND Products.ProductID=OrderDetails.ProductID"
Set rs = Conn.Execute(SQL)
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

<% REM output order details %>
	
<center>
<Table COLSPAN=8 CELLPADDING=5 BORDER=1>

<!-- BEGIN column header row -->

<TR>
<TD ALIGN=CENTER BGCOLOR="#800000">
<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>Product Name</FONT>
</TD>
<TD ALIGN=CENTER BGCOLOR="#800000">
<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>Quantity</FONT>
</TD>
</TR>


<% Do While Not rs.EOF  %>

<!-- BEGIN first row of inserted product data -->
<TR>
<TD BGCOLOR="f7efde" ALIGN=CENTER>
<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2><%=rs("ProductName")%></FONT>

</TD>
<TD BGCOLOR="f7efde" ALIGN=CENTER>
<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2><%=rs("Quantity")%></FONT>
</TD>
</TR>

	<%
	rs.MoveNext
	Loop
	rs.Close
	Conn.Close
	%>
</TABLE>
<P>
Print out a copy of this page and retain your order number for future
reference. Should you have questions about your order, contact us at
1-800-668-0124.
<P>
<b>Order Reference Number: 
<%=Request.QueryString("OrderID")%>
</b>
<HR>
<h3>Explore what else this site has to offer!</h3>

<OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
  WIDTH="550" HEIGHT="140" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
 <PARAM NAME="MOVIE" VALUE="/orderexpress/images/infohome4.swf">

    <EMBED SRC="/orderexpress/images/infohome4.swf" WIDTH="550" HEIGHT="140"
     PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
    </EMBED>
	<NOEMBED>
	<A HREF="http://www.newhome.on.ca/newhome_shoppers/info-home/index.htm"
	<img src="/orderexpress/images/infohome.gif" WIDTH="550" HEIGHT="140" ALT="InfoHome" border=0></A>
	</NOEMBED>
</OBJECT>
<P>
<A HREF="http://www.newhome.on.ca/site_nav.htm">Return to Site Navigation</a>
</center>
<HR>

<!--#include virtual="/order/disclaimer.inc"-->

</BODY>
</HTML>