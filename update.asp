<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							update.asp
'======================================================================
'
' Filename:	update.asp
' Description:	Presents the existing product information in a form
'				for updating. The productID must be passed to this
'				page when it is called.
'				User has option to update changes or cancel.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files): requires ProductID as querystring
'
'	Called From:	productupdate.asp
'	Calls:			update.asp (recursive)
'					AdminOrderExpress.asp
'
'	Version:	Version 1.0					Date: Sept.4,1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		read product information from database
'		populate form with information
'		update information or cancel
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

<% 
'page requires product ID. If not sent, redirect user to admin page
Product = Request.QueryString("ProductID")
If Product = "" OR IsNumeric(Product)=False Then
	Response.Redirect("/order/admin/productupdate.asp")
End If
%>

<% REM Set up error handling. %>
<% on error resume next %>

<%
'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively. The following section 
'	is used to check the information submitted in the form.
'	 The first time this form is displayed, the following section
'	is not processed because Request("Action") = "" (user has 
'	not pressed the Update Product button, named "Action", 
'	at the bottom of the page). Pressing the enter button, 
'	sets "Action" = "Update Product" or "Cancel"
'----------------------------------------------------------------------

SELECT CASE Request("Action")
   
	CASE "Update Product"
		'update item
		sql = "UPDATE Products SET ProductName=" & CheckString(Request("ProductName"),",")
		sql = sql & "ProductDescription=" & CheckString(Request("ProductDescription"),",")
		sql = sql & "ProductType =" & CheckString(Request("ProductType"),",")
		sql = sql & "UnitsInStock =" & CLng(Request("UnitsInStock")) &","
		sql = sql & "ConsumerPrice= " & CCur(Request("ConsumerPrice")) &","
		sql = sql & "BuilderPrice= " & CCur(Request("BuilderPrice")) &","
		sql = sql & "OneFree=" & CBool(Request("OneFree")) &","
		sql = sql & "ProductImage=" & CheckString(Request("ProductImage"),",")
		sql = sql & "ProductFlag=" & CheckString(Request("ProductFlag"),",")
		sql = sql & "ProductOnline=" & CBool(Request("ProductOnline")) &","
		sql = sql & "ProductPreviewURL=" & CheckString(Request("ProductPreviewURL"),"")
		sql = sql & " Where ProductID = " & CLng(Product)
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)
		rs.Close
		Conn.Close

		'add information to categories table
		sql = "UPDATE  bndProdCat Set CategoryID= " & CLng(Request("Category"))
		sql = sql & " Where ProductID = " & CLng(Product)

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Session("dbConnectionString")
		Conn.Execute(sql)

		rs.Close
		Conn.Close

		'redirect user to Order Express Admin page
		Response.Redirect "/order/admin/adminOrderExpress.asp"

	CASE "Cancel"
		
		'changes ignored and user redirected to Admin page
		Response.Redirect "/order/admin/AdminOrderExpress.asp"

END SELECT
'----------------------------------------------------------------------
%>

<% 
'open table and output all product information
set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open Session("dbConnectionString")
SQL = "SELECT * FROM Products WHERE ProductID=" & CLng(Product) 
Set rs = Conn.Execute(SQL)
SQL2 = "SELECT CategoryID FROM bndProdCat WHERE ProductID=" & CLng(Product) 
Set rsCat = Conn.Execute(SQL2)

'populate form with product info 
'(exept ProductID, since not to be changed and not relevant to enduser
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

<FORM ACTION="/order/admin/update.asp?ProductID=<%=rs("ProductID")%>" METHOD=POST>
	<table border="0" width="80%">
		<TR>
			<TD>
				Product Name
			</TD>
			<TD>
				<input type=text size=80 maxlength=250 
				name="ProductName" value="<%=rs("ProductName")%>">
			</TD>
		</TR>
		<TR>
			<TD>
				Product Despcription
			</TD>
			<TD>
				<textarea rows=8 columns=80 maxlength=500 
				name="ProductDescription" ><%=rs("ProductDescription")%>
				</textarea>
			</TD>
		</TR>
		<TR>
			<TD>
				Category
			</TD> 
			<TD>
<%
				' determine previous settings and enter into Category statement
				SELECT CASE rsCat("CategoryID")
					CASE 1
						mysel1="SELECTED"
						mysel2=""
						mysel3=""
					CASE 2
						mysel1=""
						mysel2="SELECTED"
						mysel3=""
					CASE 3
						mysel1=""
						mysel2=""
						mysel3="SELECTED"
				END SELECT
%>
				<SELECT NAME="Category">
					<option value="1" <%=mysel1%>>Consumer
					<option value="2" <%=mysel2%>>Builder
					<option value="3" <%=mysel3%>>Technical
				</SELECT>
			</TD>
		</TR>
		<TR>
			<TD>
				Units in Stock
			</TD> 
<% 
				' determine previous settings and enter into Units in Stock statement
				SELECT CASE rs("UnitsInStock")
					CASE 0
						mysel1=""
						mysel2="SELECTED"
					CASE 1
						mysel1="SELECTED"
						mysel2=""
				END SELECT
%>
			<TD>
				<SELECT NAME="UnitsInStock">
					<option value="0" <%=mysel2%>>no
					<option value="1" <%=mysel1%>>yes
				</SELECT>	
			</TD>
		</TR>
		<TR>
			<TD>
				Consumer Price
			</TD>
			<TD>
				<input type=text size=8 maxlength=8 
				name="ConsumerPrice" value="<%=rs("ConsumerPrice")%>">
			</TD>
		</TR>
		<TR>
			<TD>
				Builder Price
			</TD>
			<TD>
				<input type=text size=8 maxlength=8 
				name="BuilderPrice" value="<%=rs("BuilderPrice")%>">
			</TD>
		</TR>
		<TR>
			<TD>
				Product Type
			</TD>
			<TD>
<%
				' determine previous settings and enter into Category statement
				'contents based on ProdType table
				SELECT CASE rs("ProductType")
					CASE 1
						mysel1="SELECTED"
						mysel2=""
						mysel3=""
						mysel4=""
						mysel5=""
						mysel6=""
					CASE 2
						mysel1=""
						mysel2="SELECTED"
						mysel3=""
						mysel4=""
						mysel5=""
						mysel6=""
					CASE 3
						mysel1=""
						mysel2=""
						mysel3="SELECTED"
						mysel4=""
						mysel5=""
						mysel6=""
					CASE 4
						mysel1=""
						mysel2=""
						mysel3=""
						mysel4="SELECTED"
						mysel5=""
						mysel6=""
					CASE 5
						mysel1=""
						mysel2=""
						mysel3=""
						mysel4=""
						mysel5="SELECTED"
						mysel6=""
					CASE 0
						mysel1=""
						mysel2=""
						mysel3=""
						mysel4=""
						mysel5=""
						mysel6="SELECTED"

				END SELECT
%>
				<SELECT NAME="ProductType">
					<option value="1" <%=mysel1%>>Manuals
					<option value="2" <%=mysel2%>>Booklets
					<option value="3" <%=mysel3%>>Case Study
					<option value="4" <%=mysel4%>>Building Smart
					<option value="5" <%=mysel5%>>Video/Study Kits
					<option value="0" <%=mysel6%>>
				</SELECT>
			</TD>

		<TR>
			<TD>
				One Free
			</TD>
<%
				' determine previous settings and enter into OneFree statement
				SELECT CASE rs("OneFree")
					CASE -1
						mysel1="CHECKED"
						mysel2=""
					CASE 0
						mysel1=""
						mysel2="CHECKED"
				END SELECT
%>
			<TD>
				<input type=radio name="OneFree" value="true" <%=mysel1%>>
				yes &nbsp; &nbsp;
                <input type=radio name="OneFree" value="false" <%=mysel2%>>
				no
			</TD>
		</TR>
		<TR>
			<TD>
				Product Image Filename
			</TD>
			<TD>
				<input type=text size=50 maxlength=50 
				name="ProductImage" value="<%=rs("ProductImage")%>">
			</TD>
		</TR>
		<TR>
			<TD>
				Product Flag
			</TD>
			<TD>
<% 
				' determine previous settings and enter into Category statement
				SELECT CASE rs("ProductFlag")
					CASE ""
						mysel1="SELECTED"
						mysel2=""
						mysel3=""
					CASE "new"
						mysel1=""
						mysel2="SELECTED"
						mysel3=""
					CASE "revised"
						mysel1=""
						mysel2=""
						mysel3="SELECTED"
				END SELECT
%>
				<SELECT NAME="ProductFlag">
					<option value=""		<%=mysel1%>>-none-
					<option value="new"		<%=mysel2%>>New
					<option value="revised"	<%=mysel3%>>Revised
				</SELECT>	
			</TD>
		</TR>
		<TR>
			<TD>
				Product On-line
			</TD>
<%
				' determine previous settings and enter into ProductOnline
				If rs("ProductOnline") Then
						mysel1="CHECKED"
						mysel2=""
				Else
						mysel1=""
						mysel2="CHECKED"
				End If
%>
			<TD>
				<input type=radio name="ProductOnline" 
				value="true" <%=mysel1%>>yes &nbsp; &nbsp;
                <input type=radio name="ProductOnline" 
				value="false" <%=mysel2%>>no
			</TD>
		</TR>
		<TR>
			<TD>
				Product Preview URL
			</TD>
			<TD>
				<input type=text size=70 maxlength=70 
				name="ProductPreviewURL" value="<%=rs("ProductPreviewURL")%>">
			</TD>
		</TR>
	</TABLE>

	<center>
	<INPUT TYPE=SUBMIT NAME="Action" VALUE="Update Product">
	<INPUT TYPE=SUBMIT NAME="Action" VALUE="Cancel"> 
	</center>
<%
'close data connection
rs.Close
rsCat.Close
Conn.Close
%>

</FORM>
<P>


</BODY>
</HTML>