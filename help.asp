<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							Help.asp
'======================================================================
'
' Filename: Help.asp
' Description:	Provides help and policy information for end users
'				using Order Express
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):
'
'	Called From:	CustomerInfo.asp
'					Basket.asp
'					orderexpress.asp
'	Calls:			orderexpress.asp
'
'	Version:	Version 1.0					Date: Sept.4, 1997
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
<TITLE>Order Express - Help</TITLE></HEAD>
<Body background="/orderexpress/images/SAND_BKGRND.GIF" topmargin="0">
<img src="/orderexpress/images/logo2w1c.gif" width="150"
				height="56" align=middle>
<center>
	<a href="/order/orderexpress.asp">
	<img src="/orderexpress/images/orderexpresssm.gif" 
		border=0 align=middle>
	</a>
</center>

Watch this area for information about using Order Express and 
policy information about:
<ul>
	<li>Orders from continental North America
	<li>Orders outside continental North America
	<li>Processing no-charge orders
	<li>Processing orders that require payment
	<li>Orders from ONHWP Registered Builders
	<li>Shipping policies and costs
</ul>

<h3>How to use the Shopping Basket</h3>

The shopping basket is only visible if it contains a product. To view the 
contents of your basket, select the "Review Basket" button on the front 
page, or select a link from one of the product pages. You will also be 
taken to the shopping basket when you select a product.
<p>
To place a product into your shopping basket, click on the "Order" image 
beside the product you desire. This will place the product in your 
shopping basket and show you the contents of the basket. You may then 
change the quantity (default is 1 unit) and press "Recalculate" to 
view the updated costs. 
<P>
Items can be removed from the basket at any time by "un-selecting" the 
Confirm check box. The entire order will be cleared if the "Cancel Order" 
button is selected. 
<P>
The "Shop For More" button returns you to the Order Express main page. 
You may select products from any of the three product categories.


<hr>
<A HREF="orderexpress.asp">[Return to Order Express Main Page]</A> 
<A HREF="CustomerInfo.asp">[Return to Customer Information Page]</A> 
<P>
<!--#include virtual="/order/disclaimer.inc"-->

</BODY>
</HTML>