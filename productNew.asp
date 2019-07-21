<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							productNew.asp
'======================================================================
'
' Filename:	productNew.asp
' Description:	For adding a new product to the Order Express Database.
'				User is presented with an empty form to fill out. 
'				UserOptions:	Enter Product - adds it to database
'								Cancel - do not add info
'				No content validation performed.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):
'
'	Called From:	AdminOrderExpress.asp
'	Calls:			productNew.asp (recursive)
'					AdminOrderExpress.asp
'
'	Version:	Version 1.0					Date: Sept.4,1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		present empty form to fill in
'		add contents to database, if desired
'		return user to Order Express Admin page
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

<% Response.Expires = 0 %>
<% REM Previous line is used to work around proxy caching problems. %> 

<% REM Set up error handling. %>
<% on error resume next %>

<%
'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively. The following section 
'	is used to check the information submitted in the form.
'	 The first time this form is displayed, the following section
'	is not processed because Request("Action") = "" (user has 
'	not pressed the Enter Product button, named "Action", 
'	at the bottom of the page). Pressing the enter button, 
'	sets "Action" = "Enter Product" or "Cancel"
'----------------------------------------------------------------------
  

SELECT CASE Request("Action")
   
	CASE "Enter Product"

		' Note: Validation that consumer and builder prices are  >= 0, is 
		' done on the client side (VBScript). Administrator is assumed
		' to use MS Internet Explorer 3.x or better.

		'Assume data OK and enter it
		'All Variables must match their SQL data types for passing to the Table

		'create SQL statement to enter product data
		sql = ""
		sql = "INSERT INTO Products (ProductName,ProductDescription,ProductType,UnitsInStock,ConsumerPrice,BuilderPrice,OneFree,ProductImage,ProductFlag,ProductOnline,ProductPreviewURL) VALUES ("
		sql = sql & CheckString(Request("ProductName"),",")
		sql = sql & CheckString(Request("ProductDescription"),",")
		sql = sql & CheckString(Request("ProductType"),",") 
		sql = sql & CLng(Request("UnitsInStock")) &","
		sql = sql & CCur(Request("ConsumerPrice")) &","
		sql = sql & CCur(Request("BuilderPrice")) &","
		sql = sql & CBool(Request("OneFree")) &","
		sql = sql & CheckString(Request("ProductImage"),",")
		sql = sql & CheckString(Request("ProductFlag"),",")
		sql = sql & CBool(Request("ProductOnline")) &","
		sql = sql & CheckString(Request("ProductPreviewURL"),")")
	
		'create connection to database and commit data
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
		'get product ID
		sql = "select @@identity"
		set rs = Conn.Execute(sql)
		ProductID = CLng(rs(0))
		rs.Close
		Conn.Close

		'add information to categories table
		sql = "INSERT INTO bndProdCat(ProductID,CategoryID) VALUES ("
		sql = sql & ProductID & ","	
		sql = sql & CLng(Request("Category")) & ")"	
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
		rs.Close
		Conn.Close

		'redirect user to Order Express Admin page
		Response.Redirect "/order/admin/adminOrderExpress.asp"

	CASE "Cancel"
		
		'did not want to add new product
		' redirect to Order Express Admin page
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
				height="56" align=top >
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
<Font color="#FF0000">
	Requires Internet Explorer 3.x or better
	(VBScript support).</font>
</center>


<FORM METHOD=POST ACTION="productNew.asp" NAME="New">
<H2 ALIGN=CENTER>Add a New Product</H2>

	<TABLE>
		<TR>
			<TD align=right>
				Product Name
			</TD>
			<TD>
				<input type=text size=80 maxlength=150 
				name="ProductName">
			</TD>
		</TR>
		<TR>
			<TD align=right valign=top>
				Product Despcription
			</TD>
			<TD>
				<textarea rows=8 columns=80 maxlength=600 
				name="ProductDescription"></textarea>
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Category
			</TD> 
			<TD>
				<SELECT NAME="Category">
					<option value="1">Consumer
					<option value="2">Builder
					<option value="3">Technical
				</SELECT>	
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Units in Stock
			</TD> 
			<TD>
				<SELECT NAME="UnitsInStock">
					<option value="1" checked>yes
					<option value="0">no
				</SELECT>	
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Consumer Price
			</TD>
			<TD>
				$<input type=text size=8 maxlength=8 value="0"
				name="ConsumerPrice">
				<fONT color="#FF0000">(Number required)</font>
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Builder Price
			</TD>
			<TD>
				$<input type=text size=8 maxlength=8 value="0"
				name="BuilderPrice">
				<fONT color="#FF0000">(Number required)</font>
			</TD>
		</TR>
		<TR>
			<TD align=right>Product Type
			</TD>
			<TD>
				<SELECT NAME="ProductType">
					<option value="0" SELECTED>
					<option value="1" >Manuals
					<option value="2" >Booklets
					<option value="3" >Case Study
					<option value="4" >Building Smart
					<option value="5" >Video/Study Kits
				</SELECT>
				(For products in the Technical Category)
			</TD>
		</TR>
		<TR>
			<TD  align=right>
				Is the First One Free?
			</TD>
			<TD>
				<input type=radio name="OneFree" 
				value="true" >yes &nbsp; &nbsp;
                <input type=radio name="OneFree" 
				value="false" checked>no
			</TD>
		</TR>
		<TR>
			<TD  align=right>
				Product Image Filename
			</TD>
			<TD>
				<input type=text size=50 maxlength=50 
				name="ProductImage">(no path required)
			</TD>
		</TR>
		<TR>
			<TD align=right valign=top>
				Product Flag
			</TD>
			<TD>
				<SELECT NAME="ProductFlag">
					<option value="" SELECTED>-none-
					<option value="new"		>New
					<option value="revised"	>Revised
				</SELECT>	
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Is Product (or Product Preview) On-line?
			</TD>
			<TD>
				<input type=radio name="ProductOnline" 
				value="true" >yes &nbsp; &nbsp;
                <input type=radio name="ProductOnline" 
				value="false" checked>no
			</TD>
		</TR>
		<TR>
			<TD align=right>
				Product Preview URL (full URL reg'd for online document)
			</TD>
			<TD>
				<input type=text size=70 maxlength=70 
				name="ProductPreviewURL" value="http://">
			</TD>
		</TR>
	</TABLE>
<P>
	<center>
		<INPUT TYPE=SUBMIT NAME="Action" VALUE="Enter Product">
		<INPUT TYPE=SUBMIT NAME="Action" VALUE="Cancel">
	</center>

</FORM>
<P>
<!--#include virtual="/order/disclaimer.inc"-->


</BODY>

<%' validation of prices - client side %>
<SCRIPT LANGUAGE="VBSCRIPT">
	
	SUB Action_OnClick
	Dim F
	Set F=Document.New
	C = "False"
	B = "False"
		If IsNumeric(F.ConsumerPrice.Value) Then
			F.ConsumerPrice.value = abs(F.ConsumerPrice.Value)
			C = "True"
		Else		
			MsgBox "Consumer Price must be a number zero or greater", 16
		 	C = "False"
		End If
		If IsNumeric(F.BuilderPrice.Value) Then 			
			B = "True"
			F.BuilderPrice.value = abs(F.BuilderPrice.Value)
		Else
			MsgBox "Builder Price must be a number zero or greater", 16 		
		End If
	
		If C = "True" and B = "True" Then
		
			F.Submit	
		End If
	END SUB
	
</SCRIPT>



</HTML>