<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							Basket.asp
'======================================================================
'
' Filename: Basket.asp
' Description:	The shopping basket for the user. This page displays
'		the contents of the users shopping basket. 
'		Prices shown are the consumer price. Any applicable discounts
'		are applied later.
'		The user can:
'			1. Specify the quantity desired
'			2. Remove items from the basket by "unselecting" them
'			3. Return to main page to shop for more
'			4. Proceed with the order
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Macromedia Flash (optional)
'										Disclaimer.inc
'
'	Called From:	prodservcat.asp
'					OrderExpress.asp
'	Calls:			CustomerInfo.asp
'					OrderExpress.asp
'					Basket.asp (resursive)
'
'	Version:	Version 1.0					Date:	Sept.4, 1997
'
'	Enhancements/Fixes: 
' Developed by:	Codefusion Communications Inc.
' Programmers:	RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		presents table of shopping basket contents
'		If select checkout, redirect to customer information page
'		If select "shop for more", redirect to main page 
'		If choose to cancel order, basket emptied, return to main page
'		If change quantity, recalculate costs and re-display basket
'	end
'======================================================================
%>

<%
'========================== Subroutine Start ==========================
'	Subroutine Name: RemoveUnselected
'	Description:	Removes items from the shopping basket that 
'					have the "confirm" box un-selected.
'					Removes items from the Session level basket.
'	Parameters:
'	Assumptions:
'----------------------------------------------------------------------
%>

<SCRIPT LANGUAGE=VBScript RUNAT=Server>
SUB RemoveUnselected ()
	iCount = Session("ItemCount")
	basket = Session("MyBasket")
	For i = 1 to iCount
		If Request("Confirm" & CStr(i)) = "" Then
			iCount = iCount - 1
			For x = 0 to UBound(basket,1)
				basket(x,i) = ""
			Next
			n = i
			while n < UBound(basket,2)
				For x = 0 to UBound(basket,1)
					basket(x,n) = basket(x,n + 1)
					basket(x,n + 1) = ""
				Next
				n = n + 1
			wend	
		End If
    	Next
	Session("MyBasket") = basket
	Session("ItemCount") = iCount
END SUB
</SCRIPT>
<%
'========================= Subroutine End =============================
%>

<%
'========================== Subroutine Start ==========================
'	Subroutine Name: VerifyQuantity
'	Description:	Verifies that the quantity value is a numberic
'					between 1 and 9999.
'					If the quantity entry is not a nummeric
'						 -> set quantity = 1
'					If the quantity is greater than 9999,
'						 then set it to 9999
'	Parameters:
'	Assumptions:
'----------------------------------------------------------------------
%>

<SCRIPT LANGUAGE=VBScript RUNAT=Server>
SUB VerifyQuantity ()
	iCount = Session("ItemCount")
	basket = Session("MyBasket")
	'convert to a numeric
	For i = 1 to iCount       
		Quantity = Request("Quantity" & CStr(i))
	
		If IsNumeric(Quantity) Then
			If CLng(Quantity) > 9999 Then
				Quantity = 9999
			End If
			basket(PRODQUANTITY,i) = abs(CLng(Quantity))
		Else
			basket(PRODQUANTITY,i) = 1
		End If
	Next
	Session("MyBasket") = basket
	Session("ItemCount") = iCount
END SUB
</SCRIPT>
<%
'========================= Subroutine End =============================
%>

<%
' Constants and Variables
Const MaxItems=30			'Maximum items allowed in basket

' Constants for shopping basket array
Const PRODPRICE= 0			'Index for Price 
Const ITEMCHECKED = 1		'Index indicating item selected/desired 
Const PRODID = 2			'Index for ProductID
Const PRODNAME =3			'Index for ProductName
Const PRODQUANTITY = 4		'Index for Quantity requested
Const PRODONEFREE = 5		'Index for "first on free" Flag
Const PRODINSTOCK = 6		'Index for "products in stock" Flag

'set variables to use locally
iCount = Session("ItemCount")	'number of items in basket
basket = Session("MyBasket")	'reference to basket

'determine the product selected to be added to the basket
' assign ProductID to a local varialble
prod = CInt(Request.QueryString("ProductID"))	


' Obtain information about the desired product and 
' place the product information into the shopping basket array
'
' if a product has been requested, obtain info and add it to basket
If prod <> 0 Then
	' increment counter of the number of items in basket
	If iCount < MaxItems Then
		iCount = iCount + 1
	End if
	'write item count to session variable so it is 
	' available to the other files
	Session("ItemCount") = iCount

	' Connect to the database
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Session("dbConnectionString")
	'create query to find required product information
	SQL = "SELECT ProductID, ProductName, ConsumerPrice, OneFree, UnitsInStock FROM Products WHERE ProductID = " & prod 
	Set RS = Conn.Execute(SQL)
	'assign product information to basket 
	If Not IsEmpty(RS) Then
		basket(ITEMCHECKED,iCount) = "CHECKED"
		basket(PRODPRICE,iCount) = RS("ConsumerPrice")
		basket(PRODID,iCount) = RS("ProductID")
		basket(PRODNAME,iCount) = RS("ProductName")
		basket(PRODQUANTITY,iCount) = 1
		basket(PRODONEFREE,iCount) = RS("OneFree")
		basket(PRODINSTOCK,iCount) = RS("UnitsInStock")
		Session("MyBasket") = basket
	End If
	RS.Close
	Conn.Close
End If  'end of product info into basket array
  
 
'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively and the following 
'	select case deals with the various actions 
'	that a user can select.
'	Once the shopping basket is displayed, the user can select
'	on of 4 options: 
'			Shop for More (return to main page)
'			Recalculate (basket totals, removes any unselected items)
'			Place Order 
'			Cancel (clears order)
' The first time this page is written, the "Action" = ""
' and the following Select Case options are ignored.
'----------------------------------------------------------------------

SELECT CASE Request("Action")
   
CASE "Shop for More"
	
	VerifyQuantity	'verify that quantity between 1 and 9999

	RemoveUnselected	'Remove any uncofirmed products from basket

	'redirect the user to the main page for more shopping
	Response.Redirect "/order/orderexpress.asp"

CASE "Recalculate"
	
	VerifyQuantity	'verify that quantity between 1 and 9999

	RemoveUnselected	'Remove any uncofirmed products from basket

	'do NOT redirect, continue with HTML below
	
CASE "Cancel Order"
	iCount = 0	
	Session("ItemCount") = iCount
	Response.Redirect "/order/orderexpress.asp"

CASE "Place Order"
	
	VerifyQuantity	'verify that quantity between 1 and 9999

	RemoveUnselected	'Remove any uncofirmed products from basket

	
	'Redirect user to the Customer Information page
	' to begin processing the order.
	Response.Redirect "/order/Customerinfo.asp"
	
END SELECT
' ---------------------------------------------------------------------
%>

<HTML>
<HEAD>
	<TITLE>Order Express = Shopping Basket</TITLE>
</HEAD>
<BODY background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0"> 
<table width=100%>
	<tr>
		<td width= 152 valign=top>
			<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=top alt="ONHWP Logo">
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
				 WIDTH="85" HEIGHT="85" CODEBASE="http://active.macromedia.com/flash2/cabs/swflash.cab#version=2,0,0,0">
				<PARAM NAME="MOVIE" VALUE="/orderexpress/images/doghelpci.swf">

				<EMBED SRC="/orderexpress/images/doghelpci.swf" WIDTH="85" HEIGHT="85"
					 PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash2">
				</EMBED>
				<NOEMBED>
				<A HREF="http://205.210.250.157/order/help.asp">
					<img src="/orderexpress/images/doghelp.gif" WIDTH="85" HEIGHT="85" ALT="Click here for help" border=0>
				</A>
				</NOEMBED>
			</OBJECT>
		</td>
		<td valign=top>
			<h2 Align=center>Shopping Basket</h2>
		</td>
	</tr>
</table>

<center>
<FORM ACTION="/order/basket.asp?" METHOD=POST>
	<Table COLSPAN=8 CELLPADDING=5 BORDER=0>
		<%' Create the Column Title Row %>
		<TR>
			<TD ALIGN=CENTER BGCOLOR="#800000">
				<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#ffffff" SIZE=2>
					Confirm
				</FONT>
			</TD>
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
		<%' End Column Title Row %>
<%
		' Output all the items in the shopping basket
		' Create a new row for each item
		iSubtotal = 0		'variable to calculate order subtotal
		For i = 1 to iCount
%>

		<TR>
			<TD ALIGN=CENTER BGCOLOR="f7efde">
				<% 'Display a checkbox for each item where the user can confirm
					' the item or remove it from the order by unselecting
					' the checkbox
				If basket(ITEMCHECKED,i) = "CHECKED" Then
				%>
					<INPUT TYPE="CHECKBOX" NAME=<%Response.Write "Confirm" & CStr(i)%> 
							VALUE="Confirmed" CHECKED>
				<%End If%>
			</TD>
			<TD BGCOLOR="f7efde" ALIGN=CENTER>
				<%' Display the product Name %>
				<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
					<%=basket(PRODNAME,i)%>
				</FONT>
				<%' Display the words "Out of Stock" if item is out of stock
				 If basket(PRODINSTOCK,i) < 1 then 
				 %>
					<br><FONT COLOR="#FF000" Size=1>Out of Stock</FONT>
				<% End If %>
			</TD>
			<TD BGCOLOR="f7efde" ALIGN=CENTER>
				<%If basket(ITEMCHECKED,i) = "CHECKED" Then%>
					<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<INPUT TYPE=TEXT NAME=<%Response.Write "Quantity" & CStr(i)%> 
							VALUE="<%=basket(PRODQUANTITY,i)%>" SIZE=2 MAXLENGTH=4>
					</FONT>
				<%End If%>
			</TD>
			<TD BGCOLOR="f7efde" ALIGN=RIGHT>
				<%'Display Product Price
				'If the first copy of this produt is free, then display this 
				'information and display price of additional copies
				If basket(PRODONEFREE,i)="True" then
				 %>
					<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						1 copy FREE
					</FONT>
					<BR>
					<FONT SIZE=1>
						Additional copies<% = FormatCurrency(basket(PRODPRICE,i))%>
					</FONT>
				<%' Display only consumer price if no free copies for product
				Else
				%>
					<FONT STYLE="Verdana, Arial, Helvetica" SIZE=2>
						<% = FormatCurrency(basket(PRODPRICE,i))%>
					</FONT>
				<%End If%>
			</TD>
			<TD BGCOLOR="f7efde" ALIGN=RIGHT>
				<% ' To calculate cost, reduce the quantity by one
				cquantity = basket(PRODQUANTITY,i)
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
		' Add the product price to the subtotal
		If (basket(PRODPRICE,i)) <> "" Then
			 iSubTotal = iSubtotal + (basket(PRODPRICE,i) * cquantity)
		End If
		' repeat for the next item in the the basket
		Next	' loop back and output next item in another row
%>

		<%' Output the basket subtotal at the bottom of the table%>
		<TR>
			<TD COLSPAN=3>
			</TD>
			<TD BGCOLOR="f7efde" ALIGN=LEFT>
				<FONT STYLE="Verdana, Arial, Helvetica" COLOR="#800000" SIZE=2>
					Subtotal:
				</FONT>
			</TD>
			<%' write the subtotal to the session variable
			Session("SubTot") = iSubTotal
			'If the order cost is $0.00 then display the 
			'words "Free" instead of subtotal
			if iSubTotal=0 then 
			%>
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
			<TD ALIGN=RIGHT >
			</TD>
			<TD COLSPAN=3 ALIGN=RIGHT>
			<%' Output the appropriate navigation buttons %>
			<% If iCount < MaxItems Then %>
			   <INPUT TYPE=SUBMIT NAME="Action" VALUE="Shop for More">
			<% End If %>
			<% If iCount > 0 Then %>
				<INPUT TYPE=SUBMIT NAME="Action" VALUE="Place Order">
				<INPUT TYPE=SUBMIT NAME="Action" VALUE="Recalculate">
			<% End If %>
				<INPUT TYPE=SUBMIT NAME="Action" VALUE="Cancel Order">
			</TD>
		</TR>
	</TABLE>
</FORM>


</center>
<!--#include virtual="/order/disclaimer.inc"-->

</BODY>
</HTML>