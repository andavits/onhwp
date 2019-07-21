<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							prodservcat.asp
'======================================================================
'
' Filename:	prodservcat
' Description:	Displays the contents of the Order Express database
'		based on the category that the user has selected.
'		This page queries the database based on the category and 
'		returns the information in a table format to the user.
'		The user views a thumbnail graphic, Description and pricing
'		information. The product can also be selected for ordering.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Disclaimer.inc
'										Macromedia Flash (optional)
'
'	Called From:	OrderExpress.asp
'	Calls:			Basket.asp
'					OrderExpress.asp
'
'	Version:	Version 1.0						Date: Sept.4, 1997
'
'	Enhancements/Fixes: Enh- Ability to sort product by 'type'
'							in addition to 'category'
' Developed by:	Codefusion Communications Inc.
' Programmers: RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		display products from category requested
'		if technical, show by product type
'	
'	end
'======================================================================
%>
<%
' Determine what catagory of the user has requested:
'	Consumer, Builder or Technical
Category = Request.QueryString("Category")
' If no category selected, return user to main Order Express page.
If Category = "" Then
	Response.Redirect("/order/orderexpress.asp")
End If
%>

<HTML>
<Head>
<Title>Order Express Catalog - <% = Category %> </Title>
</Head>

<Body  background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0">

<%
'select the appropriate page title and image to display on the page
SELECT CASE Category
	CASE "Consumer"
		img = "consumers.gif"
		title= "Of Special Interest to Consumers"
	CASE "Builder"
		img="builders.gif"
		title="Of Special Interest to Builders/Vendors"
	CASE "Technical"
		img="technical.gif"
		title="Technical Information"
	CASE ELSE 
		title=" "
END SELECT
 %>


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
					<img src="/orderexpress/images/doghelp.gif" border=0 WIDTH="85" HEIGHT="85" ALT="Click here for help">
				</A>
				</NOEMBED>
			</OBJECT>
		</td>
		<td valign=top>
			<h2 Align=center><%= Title %></h2>
		</td>
	</tr>
</table>

<%
'if the category is Consumer or Builder, then output all itmes
'if the category id technical, output information with 
'	based on the product type (eg. videos, manuals etc.)

If Category = "Consumer" OR Category = "Builder" then
	
	'open database, run query and output data in one category

	' Create the database object
	set Conn=Server.CreateObject("ADODB.Connection")
	Conn.Open Session("dbConnectionString")
	'prepare SQL query based on the category selected by the user
	SQL = "SELECT Products.*, Categories.* " &_ 
		"FROM Products, Categories, bndProdCat " &_
		"WHERE Products.ProductID = bndProdCat.ProductID " &_
		"AND Categories.CategoryID = bndProdCat.CategoryID " &_
		"AND Categories.CategoryName = '" & Category  & "'"
	' generate the result set
	Set rs = Conn.Execute(SQL)

%>

	<center>

	<table cellpadding=5>

<%
		'Output a row for each product in the category (result set)
		Do While Not rs.EOF
%>
		<tr>
			<td align=center>

			<% ' show if item is new, revised or coming soon
			If Not(rs("ProductFlag") = "") then
				'create file name and location 
				imageurl = "/orderexpress/images/" & rs("ProductFlag") & ".gif"
			%>
				<img src="<%=imageurl %>"><br>
			<% Else %>
				&nbsp;
			<% End If %>

			<% 
			' show whether a product preview is available online
			' If it is available online, create a link to it
			'"ProductOnline" is a flag indicating if product is available online
			' " ProductPreviewURL" contains the FULL URL for the item
			If rs("ProductOnline") Then
			 %>
				<a href="<%=rs("ProductPreviewURL")%>">
					<font color=#"0000FF">Preview<br>Online</font>
				</a>
			<% Else %>
				&nbsp;
			<% End If %>
			</td>
			<td valign=top>
				<%
				'calculate image filename and file location string, output image
				imageurl = "/orderexpress/images/" & rs("ProductImage")
				 %>
				<img src="<%=imageurl %>">
			</td>
			<td Valign=top>
				<%
				'output product Name, Description and Builder Price
				%>
				<Font size=+1><b><%= rs("ProductName") %></b></font>
				<br>
				<%=rs("ProductDescription")%>
				<%
				'don't display builder price on Consumer page
				If Category <> "Consumer" then
					price=rs("BuilderPrice")
				%>
					<br>
					<font Color="#0000FF" Size=-1>ONHWP registered builders/vendors:
					<%= FormatCurrency(price)%></font>
				<% 
				End If 'end builder price display
				 %>
			</td>
			<td valign=top>
				<% price = rs("ConsumerPrice") %>
				<%= FormatCurrency(price)%>
				<%
				'output note if first copy is free
				If rs("OneFree") then
				%>
					<br>
					<nobr><FONT STYLE="Verdana, Arial, Helvetica" SIZE=2 color="#FF0000">
					First copy
					</nobr> FREE </FONT>
				<% End If %>
			</td>
			<td valign=top>
				<% 
				' Create link to Order item.
				' Create a link to the shopping basket and pass the 
				' product ID 
				%>
				<A HREF="/order/basket.asp?ProductID=<%=rs("ProductID") %>">
					<img src="/orderexpress/images/order.gif" border=0 align=top>
				</A>
			</td>
		</tr>
<%
		' go to next product in recordset
		rs.MoveNext		'more next product in recordset
		Loop			'output the HTML code for another row in the output table
		rs.Close
%>
	</table>

<%
ElseIf Category = "Technical" Then
%>
	
<%
	' read in all the product Type info from the prodType table 

	SQL = "SELECT * FROM ProdType"
	' Create the database object
	set Conn=Server.CreateObject("ADODB.Connection")
	Conn.Open Session("dbConnectionString")
	' generate the result set
	Set rstype = Conn.Execute(SQL)
	'Loop through each "Type" and output a title and all the 
	' technical products which are of that type.
	Do While Not rstype.EOF

	'	Response.Write rstype("ProdTypeID") & " " & rstype("ProdTypeName")
		'Determine Product Sub-Heading (Type)
		title = rstype("ProdTypeName")
		Response.Write "<h3>" & title & "</h3>"

		'Query database for all technical products of this type
		'prepare SQL query based on the category selected by the user
		SQL = "SELECT Products.*, Categories.* " &_ 
			"FROM Products, Categories, bndProdCat " &_
			"WHERE Products.ProductID = bndProdCat.ProductID " &_
			"AND Categories.CategoryID = bndProdCat.CategoryID " &_
			"AND Categories.CategoryName = '" & Category & "'" &_
			"AND Products.ProductType = '" & rstype("ProdTypeID") & "'"
		' generate the result set
		Set rs = Conn.Execute(SQL)
%>
		<table cellpadding=5>

<%
			'Output a row for each product in the category (result set)
			Do While Not rs.EOF
%>
		<tr>
			<td align=center>

			<% ' show if item is new, revised or coming soon
			If Not(rs("ProductFlag") = "") then
				'create file name and location 
				imageurl = "/orderexpress/images/" & rs("ProductFlag") & ".gif"
			%>
				<img src="<%=imageurl %>"><br>
			<% Else %>
				&nbsp;
			<% End If %>

			<% 
			' show whether a product preview is available online
			' If it is available online, create a link to it
			'"ProductOnline" is a flag indicating if product is available online
			' " ProductPreviewURL" contains the FULL URL for the item
			If rs("ProductOnline") Then
			 %>
				<a href="<%=rs("ProductPreviewURL")%>">
					<font color=#"0000FF">Preview<br>Online</font>
				</a>
			<% Else %>
				&nbsp;
			<% End If %>
			</td>
			<td valign=top>
				<%
				'calculate image filename and file location string, output image
				imageurl = "/orderexpress/images/" & rs("ProductImage")
				 %>
				<img src="<%=imageurl %>">
			</td>
			<td Valign=top>
				<%
				'output product Name, Description and Builder Price
				%>
				<Font size=+1><b><%= rs("ProductName") %></b></font>
				<br>
				<%=rs("ProductDescription")%>
				<%
				price=rs("BuilderPrice")
				%>
				<br>
				<font Color="#0000FF" Size=-1>ONHWP registered builders/vendors:
				<%= FormatCurrency(price)%></font>
			</td>
			<td valign=top>
				<% price = rs("ConsumerPrice") %>
				<%= FormatCurrency(price)%>
				<%
				'output note if first copy is free
				If rs("OneFree") then
				%>
					<br>
					<nobr><FONT STYLE="Verdana, Arial, Helvetica" SIZE=2 color="#FF0000">
					First copy
					</nobr> FREE </FONT>
				<% End If %>
			</td>
			<td valign=top>
				<% 
				' Create link to Order item.
				' Create a link to the shopping basket and pass the 
				' product ID 
				%>
				<A HREF="/order/basket.asp?ProductID=<%=rs("ProductID") %>">
					<img src="/orderexpress/images/order.gif" border=0 align=top width="38" height="7" alt="Click here to Order">
				</A>
			</td>
		</tr>

			
			<%
			' go to next product in recordset
			rs.MoveNext		'more next product in recordset
			Loop			'output the HTML code for another row in the output table
			rs.Close
			%>
		</table>

<%			
		' go to next product in recordset
		rstype.MoveNext		'more next product in recordset
	Loop			
	rstype.Close	'close "type" resultset
%>

<% 
End If

Conn.Close

%>

<%
' ------------------- Navigation Arrows ---------------------
%>

<HR>

<A HREF="/order/OrderExpress.asp">[Order Express Main Page]</A>
<%
'Display a button for viewing the shopping basket if there are any items in it.
If Session("ItemCount") > 0 Then	
	'items in cart, therefore show button and link it to the shopping basket page
%>
	<A HREF="/order/basket.asp">[Shopping Basket]</A>
<% End If ' End of Shopping Basket Show %>

<% if Category = "Consumer" then %>
	<A HREF="/order/prodservcat.asp?Category=Builder">[ONHWP Builder/Vendor Info]</A> 
	<A HREF="/order/prodservcat.asp?Category=Technical">[Technical Information]</A> 
<% end if %>

<% if Category = "Builder" then %>
	<A HREF="/order/prodservcat.asp?Category=Consumer">[Consumer Information]</A> 
	<A HREF="/order/prodservcat.asp?Category=Technical">[Technical Information]</A> 
<% end if %>

<% if Category = "Technical" then %>
	<A HREF="/order/prodservcat.asp?Category=Consumer">[Consumer Information]</A> 
	<A HREF="/order/prodservcat.asp?Category=Builder">[ONHWP Builder/Vendor Info]</A> 
<% end if %>
</center>
<P>

<!--#include virtual="/order/disclaimer.inc"-->

<P>&nbsp;
</Body>
</HTML>