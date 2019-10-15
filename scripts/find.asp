<%@ Language=VBScript %>
<!--#include file="code.asp" -->

<%
'*****************************************************
'Name: find.asp

'Description:  Query form for Historic Property Inventory 
'              

'Requires: 
'          

'Called by: default.htm

'Results: Calls summary.asp for property summary
'Added code.asp for efficient string concatenation (array)

'Last Update: 12/26/2013

'****************************************************

'==========
'Variables
'==========

  Dim objConn, objRS
  Dim theString(4), theMString(4), strSQL(11), theList(4), theListR(4), theTString(6)
  numTItems=6  'Number of text fields to query
  numMItems=4  'Number of menu fields to query

'Text field descriptions
  theTString(0) = "<a href=""#"" onclick=""sendURL('../help.htm#Keywords');return false;"" title=""Any word or phrase""><b>Keywords:</b></a>"
  theTString(1) = "<a href=""#"" onclick=""sendURL('../help.htm#Property');return false;"" title=""Historic property name""><b>Name:</b></a>"
  theTString(2) = "<a href=""#"" onclick=""sendURL('../help.htm#Parcel');return false;"" title=""10 digit tax parcel number""><b>Parcel #:</b></a>"
  theTString(3) = "<a href=""#"" onclick=""sendURL('../help.htm#Years');return false;"" title=""Year a building was built""><b>Years Built:</b></a>"
  theTString(4) = "<a href=""#"" onclick=""sendURL('../help.htm#Architect');return false;"" title=""Architect's name""><b>Architect:</b></a>"
  theTString(5) = "<a href=""#"" onclick=""sendURL('../help.htm#Builder');return false;"" title=""Builder's name""><b>Builder:</b></a>"

'Menu choice field descriptions
  theMString(0) = "<a href=""#"" onclick=""sendURL('../help.htm#Address');return false;"" title=""Select street address""><b>Address:</b></a>"
  theMString(1) = "<a href=""#"" onclick=""sendURL('../help.htm#Survey');return false;"" title=""Select year(s) of historic survey""><b>Recorded:</b></a>"
  theMString(2) = "<a href=""#"" onclick=""sendURL('../help.htm#Style');return false;"" title=""Select architectural style""><b>Style:</b></a>"
  theMString(3) = "<a href=""#"" onclick=""sendURL('../help.htm#Form');return false;"" title=""Select form/type of building""><b>Form/Type:</b></a>"


'SQL query strings
  strSQL(0) = "SELECT DISTINCT Loc_FullAddress FROM [dt_HistoricInventory]"
  'strSQL(0) = "SELECT Loc_FullAddress FROM [Query_dt_HistoricInventory]"
  strSQL(1) = "SELECT DISTINCT year(DateRecorded) FROM [dt_HistoricDetails]"
  'strSQL(1) = "SELECT Year FROM [QueryYear_dt_HistoricDetails]"
  strSQL(2) = "SELECT DISTINCT StylesList FROM [dt_HistoricDetails]"
  'strSQL(2) = "SELECT StylesList FROM [QueryStylesList_dt_HistoricDetails]"
  strSQL(3) = "SELECT DISTINCT FormsList FROM [dt_HistoricDetails]"
  'strSQL(3) = "SELECT FormsList FROM [QueryFormsList_dt_HistoricDetails]"

'Database
  dbPath = Server.MapPath("fpdb/HistoricPropertyData.mdb")
  Set objConn = Server.CreateObject("ADODB.Connection")


'================================
'Build parts
'================================

'Build array item values
   objConn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath

   FOR i=0 to (numMItems-1)
    set theList(i) = objConn.Execute(strSQL(i))
    theListR(i) = theList(i).getrows
    set theList(i) = nothing
   NEXT


  'Close database connect & clean up
   objConn.Close 
   Set objConn = Nothing


'////////////////////////////////////////////////////////

'Loop through each query field & create the pulldown menus

   FOR i=0 to (numMItems-1)

    'See code.asp
    Set tmp = New StrConCatArray

    numrows=ubound(theListR(i),2)
   
    for counter=0 to numrows
      if (len(theListR(i)(0,counter))>30) then
       'theString(i) = theString(i) & "<option value=""" & theListR(i)(0,counter) & """>" & left(theListR(i)(0,counter),30) & " - etc</option>"   'Update extra long strings
       tmp.Add "<option value=""" & theListR(i)(0,counter) & """>" & left(theListR(i)(0,counter),30) & " - etc</option>"   'Update extra long strings
      else
       'theString(i) = theString(i) & "<option value=""" & theListR(i)(0,counter) & """>" & theListR(i)(0,counter) & "</option>"   'Short strings - verbatum in case of spaces
       tmp.Add "<option value=""" & theListR(i)(0,counter) & """>" & theListR(i)(0,counter) & "</option>"   'Short strings - verbatum in case of spaces
      end if
    next

    theString(i) = tmp.value
    Set tmp = Nothing

   NEXT

'////////////////////////////////////////////////////////


'================================
'Create page
'================================

  response.write "<META HTTP-EQUIV=""Pragma"" CONTENT=""no-cache"">"

  response.write "<HTML>"
  response.write "<HEAD>"
  response.write "<link href=""../css/master.css"" rel=""stylesheet"">"
  response.write " <style>"
  response.write "  body {overflow:auto}"
  response.write " </style>"

    response.write "<SCRIPT LANGUAGE=""javascript"">" 
%>

	var t = parent.MapFrame;

function popupWin(){
	url = "http://search.tpl.lib.wa.us/buildings/bldglogon1.asp?xNumber=747&xPrefix=-&xName=Market&xExact=&xType=-&xDirection=-&city=&style=&built=&keyword=&Bool=all&Disp=many&OPERATE=Perform+Search";
	setTimeout('windowProp(url)', 0);   //no delay before opening
}

function windowProp(url) {
	newWindow = window.open(url,'newWin','width=100,height=100, left=0, top = 10000');
	setTimeout('closeWin(newWindow)', 3000); //delay 3 seconds before closing
	parent.focus(); //move window behind main page
}

function closeWin(newWindow) {
	newWindow.close();
}

<%
    response.write "function testForEnter() "
    response.write "{ "   
    response.write "	if (event.keyCode == 13) "
    response.write "	{  "      
    response.write "		event.cancelBubble = true;"
    response.write "		event.returnValue = false;"
    response.write "         }"
    response.write "} "

    response.write "function sendURL(url) {"
    response.write "  spawn = window.open(url,'popupWindow','scrollbars=yes,resizable=yes,left=10,top=10,width=300,height=100');"
    response.write "  if (spawn.blur) spawn.focus();"
    response.write "}"

    response.write "</SCRIPT> "
    response.write "</HEAD> " 
    response.write "<body>"

    response.write "<style>"
    response.write "a {color:#424242; text-decoration: none}" 
    response.write "a:hover {color:#0066CC; text-decoration: underline}" 
    response.write "</style>" 

'Information (top) section
    response.write "<table id=""content"" width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"

    response.Write " <tr>"
    response.Write "  <td colspan=""3"" >"
    response.write    "<b>Welcome to Tacoma's <br>Historic Inventory Database</b><br>"
    response.write    "Use this database to search, map, and print information on historic properties in Tacoma."
    response.Write "  </td>"
    response.Write " </tr>"

    response.Write " <tr valign=""top"">"
    response.write "  <td><b><a href=""../help.htm"" target=""TextFrame""><font color=#0066CC>Help</font></a></b><br>&nbsp;</td>"
    response.Write "  <td align=""left""><b><a href=""../splash.htm"" target=""_blank""><font color=#0066CC>About this Site</font></a></b></td>"

    response.Write " <td></td>"
    response.Write " </tr>"

    response.Write " <tr>"
    response.Write "  <td colspan=""3"" background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td>"
    response.Write " </tr>"

    response.Write " <tr bgcolor=""#E8F2ED"">"
    response.Write "  <td colspan=""3"" align=""left""><b>SEARCH USING MAP</b></td>"
    response.Write " </tr>"

    response.Write " <tr>"
    response.Write "  <td colspan=""3"" background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td>"
    response.Write " </tr>"

    response.Write " <tr>"
    response.Write "  <td align=""left"" colspan=""2""><b>TOOLS:&nbsp;</b>"
    'response.Write "  <td colspan=""2"">"

%>
		<!--<img src="../images/nav_zoomin1.png" width=16 height=16 hspace=1 vspace=0 border=0 Title="To zoom in - click on map tool &amp; then drag box"  onmousedown="parent.MapFrame.getReady();" >-->
    <img src="../images/nav_fullextent.png" width=16 height=16 hspace=1 vspace=1 border=0 Title="Zoom To Tacoma" onmousedown="parent.MapFrame.getReady();parent.MapFrame.zoomTacoma();">
    <img src="../images/select_circle_1.gif" width=16 height=16 hspace=0 vspace=1 border=0 Title="Select properties by distance" onmousedown="parent.MapFrame.getReady('true');parent.TextFrame.location='../addmatch.htm?v=20171023';" >
    <img src="../images/pan_1.gif" width=16 height=16 hspace=0 vspace=1 border=0 Title="To pan hold mouse button down on map &amp; drag"  onmousedown="parent.MapFrame.getReady();">
    <img src="../images/identify_1.gif" width=16 height=16 hspace=0 vspace=1 border=0 Title="Click on map for property info"  onmousedown="parent.MapFrame.getReady();" >

    &nbsp;<i>Enlarge map: F11</i>

<%

    response.Write "  </td>"
    response.Write " </tr>"

    response.Write " <tr>"

    response.write "  <td colspan=""2""><img src=""../images/blue.gif"">&nbsp;Historic properties in database</td>"
    response.Write " <td></td>"
    response.Write " </tr>"
    response.Write " <tr>"
    response.write "  <td colspan=""2""><img src=""../images/magenta.gif"">&nbsp;Properties selected <br>&nbsp;</td>"
    response.Write " <td></td>"
    response.Write " </tr>"

'Search Form
    response.Write " <tr>"
    response.Write "  <td colspan=""3"" background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td>"
    response.Write " </tr>"

    response.Write " <tr bgcolor=""#E8F2ED"">"
    response.Write "  <td align=""left"" colspan=""3""><b>SEARCH USING TEXT</b></td>"
    response.Write " </tr>"

    response.Write " <tr>"
    response.Write "  <td colspan=""3"" background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td>"
    response.Write " </tr>"


   'Update map on selection & change map tool to zoom
    response.write "<tr><td colspan=""3""><form name=""search"" onSubmit=""parent.MapFrame.getReady();document.search.action='summary.asp'; document.search.submit();""  method=""POST"" target=""TextFrame""></td></tr>"



   'Text items:
    FOR i=0 to (numTItems-1)
       response.write " <tr>"
       response.write "  <td align=""left"">"
       response.write "   " & theTString(i) & "</td>"
       response.write "  <td  colspan=""2"" align=""right"">"
       response.write "   <input type=""text"" NAME="& i & " style=""font-size: 8pt;background-color: #FFFFCC; width: 195px;"" onkeydown=""testForEnter();"" ></td>"
       response.write " </tr>"
    NEXT

   'Select menu items:
    FOR i=0 to (numMItems-1)

       response.write " <tr>"
       response.write "  <td align=""left"">"

      If (i=1) then
	response.write "  " & theMString(i) & "</td>"
       response.write "  <td  colspan=""2"" align=""right"">"
	response.write "   <select NAME="& i + numTItems & "  style=""font-size: 8pt;background-color: #FFFFCC; width: 195px;"">"& theString(i) & "</select></td>"
      Else	
	response.write "     " & theMString(i) & "</td>"
       response.write "  <td  colspan=""2"" align=""right"">"
	response.write "     <select NAME="& i + numTItems & "  style=""font-size: 8pt;background-color: #FFFFCC; width: 195px;"">"& theString(i) & "</select></td>"
      End If

       response.write " </tr>"

    NEXT

     response.write " <tr>"
     response.write "  <td align=""left"">&nbsp;</td><td colspan=""2"" align=""right""><input type=submit value=""SEARCH"" style=""font-size: 8pt"" >&nbsp;&nbsp;<input type=reset value=""Clear"" style=""font-size: 8pt"" ></td>"
     response.write " </tr>"

     response.write "</form>"
     response.write "  </table>"


     response.write "</HTML>"


'================================
'Clean Up
'================================

   FOR i=0 to (numMItems-1)
    set theListR(i) = nothing
   NEXT

%>


