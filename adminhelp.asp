<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							AdminHelp.asp
'======================================================================
'
' Filename: Help.asp
' Description:	Provides help and policy information for 
'				Administrators of Order Express
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	Disclaimer.inc
'
'	Called From:	AdminOrderExpress.asp
'	Calls:			AdminOrderExpress.asp
'
'	Version:	Version 1.0						Date: Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers:	RV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		Displays Help information only, no scripts used	
'	end
'======================================================================
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



<H2 ALIGN=CENTER>Administration Help for Order Express</H2>

<h3>Update New Items Panel on Order Express Main Page</h3>

The front page contains a panel (side bar) 
that displays information about new publications.  This information is 
updated through an administrative interface (private web pages). 
No coding or file editing is required. The Order Express administrator 
can focus on the content itself without being concerned about how it is 
implemented. Product text is entered into text boxes and the displayed 
image is selected from a drop-down box. 
<P>
Currently a maximum of 10 items are permitted.
<P>
 The New Items can display the following information 
<ul>
	<li>Flag indicating whether product is "New", "Revised" or "Coming Soon" (required)
	<li>Product Title to display
	<li>Additional description or text
</ul>
The user is presented with all the currently displayed information in text boxes where it may be edited. Additional empty text boxes are provided if there were less than the maximum items displayed.


<h3>Add New Products</h3>

To add a new product, select this option from the Order Express Main 
administration page. The following information is required:
<center>
<table>
	<tr>
		<td>
			Product Name
		</td>
		<td>
			Name that you wish to appear and the product title
		</td>
	</tr>
	<tr>
		<td>
			Product Description
		</td>
		<td>
			Any additional details about the product, if any
		</td>
	</tr>
	<tr>
		<td>
			Category
		</td>
		<td>
			Select from the drop down categories either: 
				Consumer, Builder/Vendor, Technical Information
		</td>
	</tr>
	<tr>
		<td>
			Units in Stock
		</td>
		<td>
			Select "yes" of "no" from drop down box
		</td>
	</tr>
	<tr>
		<td>
			Consumer Price
		</td>
		<td>
			Required. Enter a price in text box. $0 or greater.
		</td>
	</tr>
	<tr>
		<td>
			Builder Price
		</td>
		<td>
			Required.If no special builder price is 
			applicable, enter the same price as the Consumer Price
		</td>
	</tr>
	<tr>
		<td>
			 Product Type
		</td>
		<td>
			For "Technical Information" Category Products. Select from 
			the drop down list box: Manuals, Booklets, Case Study, Building 
			Smart, Videos.<br>
			For other categories, leave the blank selection visible.
		</td>
	</tr>
	<tr>
		<td>
			Is the First One Free?
		</td>
		<td>
			Select "Yes" from the radio buttons if the first copy of the 
			product is free.
		</td>
	</tr>
	<tr>
		<td>
			Product Flag
		</td>
		<td>
			Used if a special image ("New" "Revised" ) is applicable for 
			the product. Select from the drop down list box.
		</td>
	</tr>
	<tr>
		<td>
			Is Product (or Product Preview) on-line?
		</td>
		<td>
			Select "Yes" if the document (or a preview of the document) is 
			available on the ONHWP site. Selecting this option will link 
			produce a link "Preview Online" beside the product to the URL 
			listed below.
		</td>
	</tr>
	<tr>
		<td>
			Product Preview URL
		</td>
		<td>
			Full URL to the online version of the product. This field 
			will be used if "yes" is selected above. The URL must start
			with "http://".
		</td>
	</tr>
</table>
</center>
<h3>Update Products</h3>

When "Update" a product is selected, the administrator selects 
the product of interest from a list of all the products. A product 
form is filled with the selected product details. Any of the information 
can be changed. The product information is identical to that presented 
with a "New Product" (see previous section). The change only takes when 
the update button is selected. If the Cancel button is pressed, the changes 
are not committed.

<h3>Delete Products</h3>

Similar to updating a product, deleting a product first involves 
selecting the product to delete. The product details are then shown 
and the administrator must confirm the deletion of the product. If 
Cancel is selected, nothing is deleted and the administrator is 
returned to the main administration page. 

<h3>View Orders</h3>

This option presents a list of no charge orders followed by a list of 
orders that require payment. Order ID's and the date and time of the 
order are shown. Selecting an order from one of these lists will present 
the full order details.


<hr>

<A HREF="AdminOrderExpress.asp">[Return to Main Administration Page]</A>
</center>

</BODY>
</HTML>