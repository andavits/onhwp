<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							CustomerInfo.asp
'======================================================================
'
' Filename: CustomerInfo
' Description:	This file provides an entry form for the customer
'		to enter Name, phone number, email address,
'		Shipping address and builder number (if applicable).
'		The required fields can not be left blank.
'		None of the information (eg. builder number) is verified.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	SendMail Component from MPS
'										Macromedia Flash (optional)
'										Disclaimer.inc
'
'	Called From:	Basket.asp
'	Calls:			Thankyou.asp - for no charge orders
'					PrintOrder.asp - for orders requiring payment
'					CustomerInfo.asp (recursive)
'	Version:	Version 1.0					Date: Sept.4,1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers:	RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		Present Form for customer information input
'		If form incomplete, show form again
'		If complete form and Free order, 
'			process order (email and log) and redirect to thankyou.asp
'		If complete form and required payment,
'			log order to database (for reference, not processed)
'			redirect order to printorder.asp for printout to screen
'	end
'======================================================================
%>
<%
'========================== Function Start ============================
'	Function Name:	CheckString
'	Description:	Parses a string and formats it for entry into
'					databse via SQL statement
'	Parameters:	Accepts a string to format, s
'				Accepts a termination character for the string
'	Returns:	reformatted string
'	Assumptions:
'----------------------------------------------------------------------
%>
<SCRIPT LANGUAGE=VBScript RUNAT=Server>
FUNCTION CheckString (s, endchar)
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
   CheckString="'" & s & "'" & endchar
END FUNCTION
</SCRIPT>
<%
'========================== Function End ==============================
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
stotal = Session("SubTot")

REM Script here will verify the input customer fields
REM and write order details to database

emailmsg=""

' The variable "msg" is used as a flag and an information string
' to display to the prompting the user to fill in all the information
' Initialize to an empty string.
msg=""


'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively. The following section 
'	is used to check the information submited in the form.
'	 The first time this form is displayed, the following section
'	is not processed because Request("Action") <> "Submit Order" (user has 
'	not pressed the Enter button, named "Action", at the bottom of 
'	the page). Pressing the enter button, sets "Action" = "Enter"
'	and the submitted form fields can be verified.
'----------------------------------------------------------------------

If Request("Action")="Submit Order" Then
	'Validate that the required fields are not blank
	If Request("ContactFirstName") = "" OR _
		Request("ContactLastName") = "" OR _
		Request("BillingAddress") = "" OR _
		Request("City") = "" OR _
		Request("StateOrProvince") = "" OR _
		Request("PostalCode") = "" OR _
		Request("Country") = "" OR _
		Request("PhoneNumber") = "" OR _
		Request("EmailAddress") = "" Then
		'set message (flag) to prompt user to complete missing fields
		msg="<B><I>Please fill in all required fields.</I></B>"
	End If
	
' The form is valid and no missing fields (msg flag not set)
	If msg = "" Then
		'generate the SQL statement to enter the fields into the 
		'database.
		sql = "insert into Customers (" &_
				"CompanyName, " &_
				"CustomerTitle, " &_
				"ContactFirstName, " &_
				"ContactLastName, " &_
				"BillingAddress, " &_
				"Suite, " &_
				"City, " &_
				"StateOrProvince, " &_
				"PostalCode, " &_
				"Country, " &_
				"PhoneNumber, " &_
				"Extension, " &_
				"BuilderNumber, " &_
				"EmailAddress) " &_
				"VALUES ( "
		sql = sql & CheckString(Request("CompanyName"),",")
		sql = sql & CheckString(Request("CustomerTitle"),",")
		sql = sql & CheckString(Request("ContactFirstName"),",")
		sql = sql & CheckString(Request("ContactLastName"),",")
		sql = sql & CheckString(Request("BillingAddress"), ",")
		sql = sql & CheckString(Request("Suite"), ",")
		sql = sql & CheckString(Request("City"), ",")
		sql = sql & CheckString(Request("StateorProvince"), ",")
		sql = sql & CheckString(Request("PostalCode"), ",")
		sql = sql & CheckString(Request("Country"), ",")
		sql = sql & CheckString(Request("PhoneNumber"), ",")
		sql = sql & CheckString(Request("Extension"), ",")
		sql = sql & CheckString(Request("BuilderNumber"), ",")
		sql = sql & CheckString(Request("EmailAddress"), ")")

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
		sql = "select @@identity"
		set rs = Conn.Execute(sql)
		CustomerID = CLng(rs(0))
		rs.Close
		Conn.Close
			
' Shipping Calculations -----------------------------------------------
		'	if order is free and within North America, then shipping = 0
		'	if cost order, no shipping charge in Canada, $15 elsewhere
		Session("Shipping") = 0
		If stotal = 0  then
			Session("Shipping") = 0
		Elseif UCase(Request("Country")) = "CANADA" then
			Session("Shipping") = 0
		Else
			Session("Shipping") = 15
		End If
		
		'Session("CustomerID")= CustomerID

'Enter order into Orders and OrderDetails tables ----------------------		

		sql = "INSERT INTO Orders(CustomerID, HowOrdered, OrderDate, FreightCharge) "
		sql = sql & "VALUES( "
		sql = sql & CustomerID
		sql = sql & ",1,"
		sql = sql & "{fn now()}, "
		sql = sql & Session("Shipping") & ")"
					
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
		sql = "select @@identity"
		set rs = Conn.Execute(sql)
		OrderID = rs(0)		'obtain OrderID
		rs.Close

' Generate Order Detail record for each item in shopping cart ---------

		'initialize email message of order (for no-charge orders)
		emailorder = ""	

		'initialize order total to zero
		total=CCur(0)

		'Process each of the items in the basket
		For i = 1 to iCount
			
			'find builder price
			sqlbprice = "SELECT Products.BuilderPrice " &_
				"FROM Products " &_
				"WHERE Products.ProductID =" & basket(PRODID,i)
			set rsbprice = Conn.Execute(sqlbprice)
			bprice = rsbprice("BuilderPrice")
			rsbprice.Close

			'determine quantity to calc line total
			'remove on from quantity if first copy is free
			If basket(PRODONEFREE,i) = "True" then
				calcnum = basket(PRODQUANTITY,i) -1
			Else
				calcnum = basket(PRODQUANTITY,i)
			End If

			'initialize line subtotal
			linetot=CCur(0)

			'calc line subtotal with either the consumer price
			'or builder price
			if Request("BuilderNumber") = "" then
				' not a builder, sum consumer price
				linetot=calcnum * CCur(basket(PRODPRICE,i))

			Else
				'if a builder, sum builder price
				'calculate total cost for item based on builder price
				linetot=calcnum * CCur(bprice)

			End If

			'Insert basket item into OrderDetails table
			sql = "INSERT INTO OrderDetails(OrderID, ProductID, Quantity, ConsumerPrice, BuilderPrice, LineTotal) "           
			sql = sql & "VALUES( "
		    	sql = sql & OrderID & ","
			sql = sql & basket(PRODID,i) & ","
			sql = sql & basket(PRODQUANTITY,i) & ","
			sql = sql & basket(PRODPRICE,i) & ","
			sql = sql & bprice & ","
			sql = sql & linetot & ")"
			Conn.Execute(sql)

			'add item (line total) to order total
			total = total + linetot

			'generate email message of order for no charge orders
			If Session("SubTot") = 0 then
				emailorder = emailorder & basket(PRODNAME,i) & ":     " & basket(PRODQUANTITY,i) & vbCrLf
			End If

		Next	'process next item in basket
	
		'update ordertotal (without shipping) in 
		' both Orders database and session variable
		sql = "UPDATE Orders set OrderTotal= "& CCur(total) & " WHERE OrderID =" & OrderID
		Conn.Execute(sql)
		Conn.Close
		Session("Total") = total

'Redirect user to correct output page ---------------------------------
		
		'Determine if a Free order 
		If CLng(Session("SubTot")) >0  then
			' Order requires payment.
			'order was entered into database for logging purposes only
			'redirect to print order page
			Response.Redirect "/order/printorder.asp?OrderID=" & OrderID
		Else
			'order was entered into database for filling order
			'Email of order is sent to ONHWP processing
			'Redirect user to "thankyou" page 
			'Create email message
			emailmsg = emailmsg & "Company Name: " & Request("CompanyName") & vbCrLf
			emailmsg = emailmsg & "Builder Number: " & Request("BuilderNumber") & vbCrLf
			emailmsg = emailmsg & "Title: " & Request("CustomerTitle")  & vbCrLf
			emailmsg = emailmsg & "First Name: " & Request("ContactFirstName")  & vbCrLf
			emailmsg = emailmsg & "Last Name: " & Request("ContactLastName") & vbCrLf
			emailmsg = emailmsg & "Address: " & Request("BillingAddress") & vbCrLf
			emailmsg = emailmsg & "Apt./Suite: " & Request("Suite") & vbCrLf
			emailmsg = emailmsg & "City: " & Request("City") & vbCrLf
			emailmsg = emailmsg & "Prov/State: " & Request("StateorProvince") & vbCrLf
			emailmsg = emailmsg & "Postal Code: " & Request("PostalCode") & vbCrLf
			emailmsg = emailmsg & "Country: " & Request("Country") & vbCrLf
			emailmsg = emailmsg & "Phone: " & Request("PhoneNumber") & vbCrLf
			emailmsg = emailmsg & "Ext: " & Request("Extension") & vbCrLf
			emailmsg = emailmsg & "Email: " & Request("EmailAddress") & vbCrLf
			emailmsg = emailmsg & "Date/Time: " & now() & vbCrLf  & vbCrLf
			emailmsg = emailmsg & "Order ID: " & OrderID & vbCrLf
			emailmsg = emailmsg & "Product Name / Quantity Ordered" & vbCrLf
			emailmsg = emailmsg & emailorder
	
			'Set up Send Mail component. 	
			set sm = Server.CreateObject ("mps.sendmail") 	
			if IsObject(sm) = FALSE then 
				mailerr= "<b>The Send Mail component error</b><p>"
			else
				receiver = "oexpress@codefusion.com" 
				subject = "Web Order - No Charge"
				feedback = sm.SendMail ( "OrderExpress Order", receiver,subject, emailmsg )
		
				'send thank you email msg to customer here 
				receiver = Request("EmailAddress")
				subject = "Thank you for using Order Express"
				emailmsg = "Thank you for using Order Express!" & vbCrLf
				emailmsg = emailmsg & "Your order is being processed."
				emailmsg = emailmsg & " If your have any questions about"
				emailmsg = emailmsg & " your order, please send your questions"
				emailmsg = emailmsg & " to orderexp@newhome.on.ca. Please"
				emailmsg = emailmsg & " refer to your order number in your"
				emailmsg = emailmsg & " correspondence. Your order number is: "
				emailmsg = emailmsg & OrderID
				feedback = sm.SendMail("ONHWP Order Express", receiver, subject, emailmsg)

			end If 'end send email

			'redirect user to thank you page
			Response.Redirect "/order/thankyou.asp?OrderID=" & OrderID
	   End If ' end no charge order

	End If  'end of msg = ""

End If  'end of Action = "Submit Order"
'----------------------------------------------------------------------
%>

<HTML>
<Head>
	<Title>Order Express - Customer Information </Title>
</Head>
<Body  background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0">

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
			<h2 Align=center>Customer Information</h2>
			To place your order simply complete the Customer Information 
			form. Once you hit submit, your order will be automatically 
			sent to ONHWP Order Express for processing. 
			<P>
			This form is for orders placed within Continental North America. 
			For orders outside this area, please forward your inquiries to
			<A HREF="mailto:orderexp@newhome.on.ca">
			orderexp@newhome.on.ca</a>.
		</td>
	</tr>
</table>

<center>

<%'Output message to user if form is not fully completed %>
<font size=+1 Color="#FF0000"><%=msg %></font>
<br>

<P>

<%'Begin form for input of customer information %>
<FORM ACTION="/order/CustomerInfo.asp" METHOD=POST>
	<table width=100%>
		<tr>
			<td align=right>
				Title: 
			</td>
			<td>
				<SELECT NAME="CustomerTitle" SIZE="1">
					<OPTION VALUE="Mr.">Mr.
					<OPTION VALUE="Ms.">Ms.
					<OPTION VALUE="Mrs.">Mrs.
				</SELECT>
			</td>
		</tr>
		<tr>
			<td align=right>
				First Name:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="ContactFirstName" 
					VALUE="<%=Request("ContactFirstName")%>" SIZE=31 
					MAXLENGTH=35>
			</td>
		</tr>
		<tr>
			<td align=right>
				Last Name:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="ContactLastName" 
					VALUE="<%=Request("ContactLastName")%>" SIZE=31 
					MAXLENGTH=35>
			</td>
		</tr>
		<tr>
			<td align=right>
				Company:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="CompanyName" 
					VALUE="<%=Request("CompanyName")%>" SIZE=31 
					MAXLENGTH=35>
				<%'(Optional) %>
			</td>
		</tr>
		<tr>
			<td align=right>
				Address:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="BillingAddress" 
					VALUE="<%=Request("BillingAddress")%>" SIZE=36 
					MAXLENGTH=75>
				&nbsp;&nbsp;
				Apt./Suite:
				&nbsp;
				<INPUT TYPE="Text" NAME="Suite" 
					VALUE="<%=Request("Suite")%>" Size=4
					MAXLENGTH=10>
			</td>
		</tr>
		<tr>
			<td align=right>
				City:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="City" 
					VALUE="<%=Request("City")%>" Size=8>
				&nbsp;&nbsp;
				Province/State:
				&nbsp;
				<INPUT TYPE="Text" NAME="StateOrProvince" 
				VALUE="<%=Request("StateOrProvince")%>" Size=2>
				(2 letter code)
			</td>
		</tr>
		<tr>
			<td align=right>
				Country:
			</td>
			<td>
				<SELECT NAME="Country" SIZE="1">
					<OPTION VALUE="Canada">Canada
					<OPTION VALUE="USA">USA
					<OPTION VALUE="Mexico">Mexico
				</SELECT>
			</td>
		</tr>
		<tr>
			<td align=right>
				Postal/Zip Code:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="PostalCode" 
					VALUE="<%=Request("PostalCode")%>" Size=5>
			</td>
		</tr>
		<tr>
			<td align=right>
				Daytime Phone:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="PhoneNumber" 
					VALUE="<%=Request("PhoneNumber")%>" Size=21>
				(Include Area Code)
			</td>
		</tr>
		<tr>
			<td align=right>
				Extension:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="Extension" 
					VALUE="<%=Request("Extensionr")%>" Size=5>
				<%' (Optional) %>
			</td>
		</tr>
		<tr>
			<td align=right>
				Email Address:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="EmailAddress" 
					VALUE="<%=Request("EmailAddress")%>" Size=35>
			</td>
		</tr>
		<tr>
			<td align=right>
				ONHWP Builder Reference Number:
			</td>
			<td>
				<INPUT TYPE="Text" NAME="BuilderNumber" 
					Value="<%=Request("BuilderNumber")%>" Size=15>
				(Optional: Required for Registered Builder Pricing)
			</td>
</tr>
</table>


	<h3><font color="#FF0000">
	Please make sure that your information is accurate before 
	submitting your order.
	</font></h3>
	<INPUT TYPE=SUBMIT NAME="Action" VALUE="Submit Order">
</center>

</form>

<center>
<A HREF="/order/basket.asp">Return to shopping cart page without placing order</A>
</center>
<P>
<!--#include virtual="/order/disclaimer.inc"-->
</Body>
</HTML>