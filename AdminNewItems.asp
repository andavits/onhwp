<%
'cci!==================================================================
'cci!		Copyright of Codefusion Communications Inc. 1997
'cci!==================================================================
%>
<%
'======================================================================
'							AdminNewItems.asp
'======================================================================
'
' Filename: AdminNewItems.asp
' Description:	Allows the Order Express administrator to update the
'				"new items" section on the front page of 
'				Order Express. "New Items" are added by new entering 
'				information into the provided text boxes. Items
'				are deleted by "unselecting" the specified checkbox.
'				The following information is displayed:
'					1. Graphic indicating "new", "revised"
'							or "Coming Soon"
'					2. Title of Item
'					3. Description of Item
'				Information is stored in a text file "NewItems.txt"
'				in the /orderexpress virtual directory.
'
'	Platform:	IIS Active Server 3.0
'	Languages:	VBScript, HTML
'	Dependencies (components, files):	/orderexpress/newitems.txt
'										(virtual directory)
'										disclaimer.inc
'	Called From: AdminOrderExpress.asp
'	Calls:		AdminNewItmes.asp (recursive)
'				AdminOrderExpress.asp
'
'	Version:	Version 1.0						Date:	Sept.4, 1997
'
'	Enhancements/Fixes:
' Developed by:	Codefusion Communications Inc.
' Programmers: RV, AV
'----------------------------------------------------------------------
' High Level Function  (pseudo-code)
'	start
'		Read text file of "new items"
'		Display "new items" in text boxes for editing
'		Write change to text file
'	end
'
'======================================================================
%>

<% Response.Expires = 0 
' Previous line is used to work around proxy caching problems.  

' Set up error handling. 
 on error resume next 

'Variables and Constants
ss=";"		'character used to seperate fields in text file
Const MAXITEMS =10		'maximum number of "new items" permitted

'----------------------------------------------------------------------
' Section Notes:
'	This page is called recursively and the following 
'	select case deals with the various actions 
'	that a user can select.
'	The first time this page is written,  "Action" = ""
'	and the following Select Case options are ignored.
'----------------------------------------------------------------------
%>
<%
SELECT CASE Request("Action")

CASE "Cancel - Return to Admin"

	'redirect the user to Order Express Administration Main Page
	Response.Redirect "/scripts/onhwp/admin/adminOrderExpress.asp"

CASE "Update Site"

	'Update the " New Items" text file with revised/new content

	'create file access objects
	Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
	Set Outstream = FileObject.CreateTextFile (Server.MapPath ("/orderexpress") & "\newitems.txt", TRUE, FALSE)

	'loop through all the rows in the fill in form
	'write them to the text file
	inum=1		'initialize counter
	Do while inum <= MAXITEMS 
		
		'create text strings for each field name
		FieldNameImg = "Image" & CStr(inum)
		FieldNameTitle = "Title" + CStr(inum)
		FieldNameText = "Text" & CStr(inum)
		FieldNameShow = "Show" & CStr(inum)
		
		'if the "show" checkbox is selected, then display that item
		'by read content for line, and writing it to the text file
		If Request.Form(FieldNameShow)="TRUE" then

			'create text string to write to file
	 		outstring = Request.Form(FieldNameImg) & ";" & Request.Form(FieldNameTitle) & ";" & Request.Form(FieldNameText)  
			Application.lock
			OutStream.WriteLine(outstring)
			Application.unlock

		End If
		
		inum=inum+1	'increment counter
	Loop			
	Outstream.Close
	
	'file has been updated, return user to Order Express Admin Page
	Response.Redirect "/scripts/onhwp/admin/AdminOrderExpress.asp"

END SELECT
'----------------------------------------------------------------------
%>

<HTML>
<HEAD>
	<TITLE>Order Express Administration - Update News Items</TITLE>
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
<h2> Update New Items</H2>

<FORM METHOD=POST ACTION="/scripts/onhwp/admin/adminnewitems.asp">
	<table border="1" width="80%" >
		<%'Display the instructions at the top of the form %>
		<tr BGCOLOR="f7efde">
			<td colspan=2>
				This page will update the "New Items" section of the 
				Order Express front page. (maximum 10 items) 
				<ul>
					<li>To remove an item, either uncheck the 
						"Show Item" box or replace existing
						text to display a new item.
					<li>To add an item, check the "Show Item" 
						box and enter the information to display. 
					<li>Each item has an image associated with it. 
						Select an appropriate image from the drop down menu.
					<li>Press the "Update Site" button when finished. 
				</ul>
			</td>
		</tr>
		<%'Display the column titles %>
		<tr >
			<th BGCOLOR="#800000">
				<Font color="#FFFFFF">Show Item</font>
			</th>
			<th  BGCOLOR="#800000">
				<Font color="#FFFFFF">New Item Details</font>
			</th>
		</tr>
<%
		'Populate the form with any existing "new items" text

		''create file access objects
		Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
		Set Instream = FileObject.OpenTextFile (Server.MapPath ("/orderexpress") & "\newitems.txt", 1, FALSE, FALSE)

		'loop through all the items in the text file
		'populate the text fields with this information
		inum=1	'initialize counter
		Do while Instream.AtEndOfStream <> True

			'read align
			newsstring = Instream.ReadLine

			'parse line of text for content
			mynum = InStr(newsstring,ss)	'intermediate variable
			newsgraphic = Left(newsstring,mynum -1)
			imageurl = "/orderexpress/images/" & newsgraphic & ".gif"
			tmpstring = Right(newsstring,Len(newsstring)-mynum) 
			mynum = InStr(tmpstring,ss)
			titlestring = Left(tmpstring,mynum-1)
			txtstring=Right(tmpstring,Len(tmpstring)-mynum) 

			'create text field names
			FieldNameImg = "Image" & CStr(inum)
			FieldNameTitle = "Title" + CStr(inum)
			FieldNameText = "Text" & CStr(inum)
			FieldNameShow = "Show" & CStr(inum)

			'output this information in form
%>
		<tr>
			<td Align=center BGCOLOR="f7efde">
				<Input type=checkbox CHECKED value="TRUE" 
				name="<%=FieldNameShow%>">
			</td>
			<td  BGCOLOR="f7efde">
				Select Image to display<BR>
				Image: 
				<SELECT NAME="<%=FieldNameImg%>" >  
				<% If newsgraphic="new" then %>         
					<OPTION VALUE="new" SELECTED>New  
				<% Else %>
					<OPTION VALUE="new">New  
				<%End If%>

				<% If newsgraphic="revised" then %> 
					<OPTION VALUE="revised" SELECTED>Revised
				<% Else %>
					<OPTION VALUE="revised">Revised
				<%End If%>

				<% If newsgraphic="comingsoon" then %> 
					<OPTION VALUE="comingsoon" SELECTED>Coming Soon
				<% Else %>
					<OPTION VALUE="comingsoon">Coming Soon
				<%End If%>
				</SELECT>
			
				Title to display for new item:
  				<input type=text size=50 maxlength=50 
				name="<%=FieldNameTitle%>" value="<%=titleString%>"> 
				<br>
				Text to display:<br>
  				<input type=text size=90 maxlength=90 
				name="<%=FieldNameText%>" value="<%=txtstring %>"> 
				<P>
			</td>
		</tr>
<%
	inum=inum+1		'increment counter
	Loop
	Instream.Close
	
	'output additional EMPTY text boxes till the total of
	'input boxes = MAXITEMS

	Do While inum <= MAXITEMS

		'generate input box names
		FieldNameImg = "Image" & CStr(inum)
		FieldNameTitle = "Title" + CStr(inum)
		FieldNameText = "Text" & CStr(inum)
		FieldNameShow = "Show" & CStr(inum)

		'output empty input boxes to screen
%>
		<tr>
			<td Align=center BGCOLOR="f7efde">
				<Input type=checkbox value="TRUE" name="<%=FieldNameShow%>">
			</td>
			<td BGCOLOR="f7efde">
				Select Image to display<BR>
				Image: 
				<SELECT NAME="<%=FieldNameImg%>" >  
					<OPTION VALUE="new" SELECTED>New  
					<OPTION VALUE="revised">Revised
					<OPTION VALUE="comingsoon">Coming Soon
				</SELECT>
			
				Title to display for new item:
  				<input type=text size=50 maxlength=50 name="<%=FieldNameTitle%>"> 
				<br>
				Text to display:<br>
  				<input type=text size=90 maxlength=90 name="<%=FieldNameText%>" > 
				<P>
			</td>
		</tr>
<%
	inum=inum+1		'increment counter
	Loop		'loop back and output remaining empty input boxes
%>
	</table>

	<P>
		<input type=submit name="Action" value="Update Site">
		<input name="Action" type=submit value="Cancel - Return to Admin">
	</p>

</FORM>

<P>
<A HREF="/scripts/onhwp/admin/adminorderexpress.asp">Return to Order Express Administration Home</a>
</center>

</BODY>
</HTML>