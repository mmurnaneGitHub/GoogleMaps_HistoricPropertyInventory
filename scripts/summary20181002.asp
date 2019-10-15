<%@ Language=VBScript %>

<%
'*****************************************************
'Name: summary.asp

'Description: Creates a summary web page of selected
' properties within the Historic Property Inventory
' database  

'Requires: 
'          

'Called by: 
'1. find.asp (summary.asp?0=&1=&2=&3=&4=&5=&6=&7=&8=&9=)
'2. Select by Circle & Identify tools

'Results: Summary of Historic Property Inventory

'Last Update: 4/14/2010
'****************************************************

'==========
'Variables
'==========

'---Sent from user---
  iMapID = Request.querystring("ID")
  iMap = Replace(Replace(Request.querystring("map"),"(",""),")","")

  if (iMap <> "") then
    theLatLong = Split(iMap,",")
    'more than 6 decimals won't work in current StreetView page 
    theLat = Round(theLatLong(0),6) 
    theLong = Round(theLatLong(1),6) 
  end if

  if (iMapID = "") then
   iMapID = Request.form("ID")
  end if

  Dim objConn, objRS, strSQL1
  Dim fldSQL(13),fldSQL2(13),iItem(10), theMString(10)

  numItems=13  'Number of fields to query
  numTItems=6  'Number of text fields to query
  numMItems=4  'Number of menu fields to query
  numCItems= numTItems + numMItems  'number of criteria fields

  theKeyword=""
  theText=""
  theCriteria=""	'for SQL search string
  theCriteria2="" 'for web page title

'---Database---
  dbPath = Server.MapPath("fpdb/HistoricPropertyData.mdb")
  Set objConn = Server.CreateObject("ADODB.Connection")

'---Selection items from user form - replace any apostrophes---
  iItem(0) = Replace(Request.form("0"),"'","' & chr(39) & '")
  iItem(1) = Replace(Request.form("1"),"'","' & chr(39) & '")
  iItem(2) = Replace(Request.form("2"),"'","' & chr(39) & '")
  iItem(3) = Replace(Request.form("3"),"'","' & chr(39) & '")
  iItem(4) = Replace(Request.form("4"),"'","' & chr(39) & '")
  iItem(5) = Replace(Request.form("5"),"'","' & chr(39) & '")
  iItem(6) = Replace(Request.form("6"),"'","' & chr(39) & '")
  iItem(7) = Replace(Request.form("7"),"'","' & chr(39) & '")
  iItem(8) = Replace(Request.form("8"),"'","' & chr(39) & '")
  iItem(9) = Replace(Request.form("9"),"'","' & chr(39) & '")


'----------------------
'SQL query table.fields
'----------------------

'For individual field query by user input
  fldSQL(0) = "dt_HistoricInventory.Historic_ID"
  fldSQL(1) = "dt_HistoricInventory.SiteNameHistoric"
  fldSQL(2) = "dt_HistoricInventory.TaxParcel_No"
  fldSQL(3) = "dt_HistoricDetails.DateOfConstruction"
  fldSQL(4) = "dt_HistoricDetails.Architect"
  fldSQL(5) = "dt_HistoricDetails.Builder"
  fldSQL(6) = "dt_HistoricInventory.Loc_FullAddress"
  fldSQL(7) = "year(dt_HistoricDetails.DateRecorded)"
  fldSQL(8) = "dt_HistoricDetails.StylesList"
  fldSQL(9) = "dt_HistoricDetails.FormsList"
  fldSQL(10) = "dt_HistoricDetails.StatementofSignificance"
  fldSQL(11) = "dt_HistoricDetails.PhysicalAppearance"
  fldSQL(12) = "dt_HistoricPhoto.PhotoFileName"

'For list of fields to include in query
  fldSQL2(0) = "dt_HistoricInventory.Historic_ID"
  fldSQL2(1) = "dt_HistoricInventory.SiteNameHistoric"
  fldSQL2(2) = "dt_HistoricInventory.TaxParcel_No"
  fldSQL2(3) = "Last(dt_HistoricDetails.DateOfConstruction)"
  fldSQL2(4) = "Last(dt_HistoricDetails.Architect)"
  fldSQL2(5) = "Last(dt_HistoricDetails.Builder)"
  fldSQL2(6) = "dt_HistoricInventory.Loc_FullAddress"
  fldSQL2(7) = "Last(year(dt_HistoricDetails.DateRecorded))"
  fldSQL2(8) = "Last(dt_HistoricDetails.StylesList)"
  fldSQL2(9) = "Last(dt_HistoricDetails.FormsList)"
  fldSQL2(10) = "Last(dt_HistoricDetails.StatementofSignificance)"
  fldSQL2(11) = "Last(dt_HistoricDetails.PhysicalAppearance)"
  fldSQL2(12) = "Last(dt_HistoricPhoto.PhotoFileName)"

	'-----------------------------
	'Create list of fields string
	'-----------------------------
	  FOR i=0 to (numItems-1)
	   if i=0 then
	     fldString = fldSQL2(i)
	   else
	     fldString = fldString & ", " & fldSQL2(i)
	   end if
	  NEXT

'-------------------
'Fields to group by
'-------------------
 groupBy = " GROUP BY dt_HistoricInventory.Historic_ID, dt_HistoricInventory.SiteNameHistoric, dt_HistoricInventory.Loc_FullAddress, dt_HistoricInventory.TaxParcel_No, dt_HistoricDetails.Historic_ID, dt_HistoricUTM.Historic_ID"


'-----------
'Page titles
'-----------
  theMString(0) = "<b>Keyword(s):</b><br>"
  theMString(1) = "<b>Property Name:</b><br>"
  theMString(2) = "<b>Parcel Number:</b><br>"
  theMString(3) = "<b>Years Built:</b><br>"
  theMString(4) = "<b>Architect:</b><br>"
  theMString(5) = "<b>Builder:</b><br>"
  theMString(6) = "<b>Address:</b><br>"
  theMString(7) = "<b>Date Recorded:</b><br>"
  theMString(8) = "<b>Style:</b><br>"
  theMString(9) = "<b>Form/Type:</b><br>"


'--------------------------
'SQL query string - Part 1  (***dt_HistoricDetails has duplicate ID entries - currently no merge of information)
'--------------------------
  strSQL1 = "SELECT " & fldString
  strSQL1 = strSQL1 & " FROM (dt_HistoricInventory" 
  strSQL1 = strSQL1 & " INNER JOIN (dt_HistoricDetails" 
  strSQL1 = strSQL1 & " LEFT JOIN dt_HistoricPhoto" 
  strSQL1 = strSQL1 & " ON dt_HistoricDetails.Historic_Details_ID = dt_HistoricPhoto.Historic_Details_ID)" 
  strSQL1 = strSQL1 & " ON dt_HistoricInventory.Historic_ID = dt_HistoricDetails.Historic_ID)" 
  strSQL1 = strSQL1 & " INNER JOIN dt_HistoricUTM" 
  strSQL1 = strSQL1 & " ON dt_HistoricInventory.Historic_ID = dt_HistoricUTM.Historic_ID"


'----------------------------------------------------
'SQL query string - Part 2: Determine WHERE criteria
'----------------------------------------------------

IF (iMapID <> "") THEN   'Search by ID#
    if (iMapID = "()")  then	
     iMapID = "(0)"
    end if

    strSQL1 = strSQL1 & " WHERE " & fldSQL(0) & " in " & iMapID

    'Current fix for duplicates
      strSQL1 = strSQL1 & groupBy

ELSE  'Search by user defined field & value

 '===================================================
  FOR i=0 to (numCItems-1)  'loop through USER FORM field names

  If ((i=0) and (iItem(i)<>"")) Then
   '-------------------------------
   'Keyword(s) search - all fields
   '-------------------------------
    theKey = Split(iItem(i),",")

    FOR a=0 to Ubound(theKey)
     For x=0 to (numItems-1)  'loop through ALL field names
      if (theKeyword = "") then
        theKeyword = fldSQL(x) & " LIKE '%" & theKey(a) & "%'"
      else
        theKeyword = theKeyword & " OR " & fldSQL(x) & " LIKE '%" & theKey(a) & "%'"
      end if
     Next
    NEXT

    theTitle(i)


  ElseIf ((i>0) and (i<numTItems) and (iItem(i)<>"")) Then
   '------------------------------
   'Word(s) search - single field
   '------------------------------
    theKey = Split(iItem(i),",")

    FOR a=0 to Ubound(theKey)
      if (theText = "") then
        theText = fldSQL(i) & " LIKE '%" & theKey(a) & "%'"
      else
        theText = theText & " OR " & fldSQL(i) & " LIKE '%" & theKey(a) & "%'"
      end if
    NEXT

    theTitle(i)

  ElseIf ((i>(numTItems-1)) and (iItem(i)<>"")) Then
   '------------------------------
   'Menu selection search
   '------------------------------
    if (theCriteria="") then

     theCriteria = fldSQL(i) & "='" & iItem(i) & "'"

    else

     theCriteria = theCriteria & " AND " & fldSQL(i) & "='" & iItem(i) & "'"

    end if

    theTitle(i)

  Else
   'nothing
  End If

  NEXT
 '==========================================================


	'///////////////////////////////////
	'---Concatenate criteria strings---

	 '---Add Word(s) search - single field---
	  IF (theText<>"") THEN
	   if (theCriteria="") then
	     theCriteria = theText
	   else
	     theCriteria = theCriteria & " AND (" & theText & ")"
	   end if
	  END IF

	 '---Keyword(s) search---
	  IF (theKeyword<>"") THEN
	   if (theCriteria="") then
	     theCriteria = theKeyword
	   else
	     theCriteria = theCriteria & " AND (" & theKeyword & ")"
	   end if
	  END IF

	  '---Add WHERE clause to criteria---
	   IF (theCriteria<>"") THEN
	     theCriteria = " WHERE " & theCriteria & groupBy   'Fix for duplicate records....
	    else
		theCriteria = groupBy
	   END IF
	'///////////////////////////////////


 strSQL1 = strSQL1 & theCriteria

END IF


'-------------------------------------------------
'SQL query string - Part 3: Determine sort order
'-------------------------------------------------
  strSQL1 = strSQL1 & " ORDER BY dt_HistoricInventory.SiteNameHistoric ASC"





'================================
'Build web page parts
'================================

'Top of web page
  response.write "<HTML>"
  response.write "<HEAD>"
  response.write "<link href=""../css/master.css"" rel=""stylesheet"">"
  response.write " <style>"
  response.write "  body {overflow:auto}"
  response.write " </style>"

%>



<SCRIPT LANGUAGE="JavaScript">


 function toggleDisplay(e)  {
    if (e.style.display == "none") {
      e.style.display = "";
  	} else {
  	      e.style.display = "none"
  	}
	}

 //function sendQuery(theQuery) {
 function sendQuery(theQuery,zoom) {
	var t = parent.MapFrame;
  if (zoom=="zoom1") {
  	t.executeQuery(theQuery,zoom);
  } else {
  	t.executeQuery(theQuery);
  }
 }

 function sendAddress(address) {
   splitstreet = address.split(" ");
   number = splitstreet[0];
   theDir = splitstreet[1].replace('.', '');
   street = splitstreet[2].replace(',', '');

   if (theDir=="N") {
       theDir="NO"
   } else if (theDir=="S") {
       theDir="SO"
   } else if ((theDir != "E") && (theDir != "W")) {
       street = theDir
       theDir=""
   }

  url = "http://search.tpl.lib.wa.us/buildings/bldglogon1.asp?xNumber=" + number + "&xPrefix=" + theDir + "&xName=" + street + "&xExact=&xType=-&xDirection=-&city=&style=&built=&keyword=&Bool=all&Disp=many&OPERATE=Perform+Search";
  spawn = window.open(url,'popupWindow','scrollbars=yes,resizable=yes,left=20,top=20,width=450,height=550');
  if (spawn.blur) spawn.focus();

  }

 function sendAPN(APN) {
  splitAPN = APN.split(" ");
  theAPN = splitAPN[0].replace('-', '').replace('-', '').replace('R', '');

  url = "http://epip.co.pierce.wa.us/CFApps/atr/epip/searchResults.cfm?s1=" + theAPN + "&stype=parcel";
  spawn = window.open(url,'popupWindow','scrollbars=yes,resizable=yes,left=20,top=20,width=450,height=550');
  if (spawn.blur) spawn.focus();

  }

 function seeBirdsEye(theLocation, theName, myLAT, myLONG) {
	thePage = 'VEphoto.asp?title=' + theLocation + '&details=' + theName + '&lat=' + myLAT + '&long=' + myLONG 
	win = window.open(thePage,"QueryWindow","width=650,height=425,scrollbars=yes,resizable=yes");	win.window.focus();
	win.window.focus();
 }

 function seeStreetView(theLocation, theName, myLAT, myLONG) {
	thePage = 'SVphoto.asp?title=' + theLocation + '&details=' + theName + '&lat=' + myLAT + '&long=' + myLONG 
	win = window.open(thePage,"QueryWindow","width=650,height=425,scrollbars=yes,resizable=yes");	win.window.focus();
	win.window.focus();
 }

</SCRIPT>
</head>
<%




'Build an array of information from query
   objConn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath
   set theAList = objConn.Execute(strSQL1) 


'================================
'Build web page body
'================================


'Write out selected record information

If theAList.eof Then  '---No Records found---

	CleanUp
	Response.Write "<body>"
	Response.Write "<hr><center>Sorry, no properties found meeting criteria.<hr>"
  	If (theCriteria2<>"") then
  	 Response.Write "<table width=""100%"" id=""content""><tr><td bgcolor=""#FFFFCC"">" & theCriteria2 & "</td></tr></table><hr>"
  	End If
	theFooter

Else   '---Records found---

	theHistoricList = theAList.getrows
	CleanUp



'Create string of Historic Property Inventory Map #s
 FOR i=0 to ubound(theHistoricList,2) 'Loop through all selected rows/records
	if (i=0) then
	 theQuery = theHistoricList(0,i)
	else
	 theQuery = theQuery & "," & theHistoricList(0,i)
	end if
 NEXT


	'Determine onload event - need theQuery string
	  selectMap


 IF (iMap = "") Then  '***Create SUMMARY PAGE***

  If (theCriteria2<>"") then
    response.write "<table id=""content"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
    response.Write " <tr><td background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td></tr>"
   Response.Write "<tr><td bgcolor=""#E8F2ED"" border=""1""><b>SEARCH RESULTS</b></td></tr>"
    response.Write " <tr><td background=""../images/hor-dash.gif""><img width=""100%"" height=""1px"" src=""../images/trans.gif""></td></tr>"
   Response.Write "<tr><td>Click on individual properties to access property details.  Property details will appear in new window. (HINT: Turn off any web browser pop-up blockers.)</td></tr>"
   Response.Write "<tr><td bgcolor=""#FFFFCC"">" & theCriteria2 & "</td></tr>"
   Response.Write "</table></center>"
  End If

  Response.Write "Number of properties = " & (ubound(theHistoricList,2)+1) & "<br>"
  Response.Write "<center><table width=""100%"" border=""1"" cellspacing=""0"" id=""content"" bordercolorlight=""#C0C0C0"" bordercolordark=""#FFFFFF"">"
  Response.Write "<tr bgcolor=""#6FA0E6"">"
  Response.Write "<td id=""content"" bordercolordark=""#C0C0C0"" bordercolorlight=""#C0C0C0"" bordercolor=""#C0C0C0""><b><center>Name</center></b></td>"
  Response.Write "<td id=""content"" bordercolordark=""#C0C0C0"" bordercolorlight=""#C0C0C0"" bordercolor=""#C0C0C0""><b><center>Address</center></b></td>"
  Response.Write "<td id=""content"" bordercolordark=""#C0C0C0"" bordercolorlight=""#C0C0C0"" bordercolor=""#C0C0C0""><b><center>Zoom</center></b></td>"
  Response.Write "<td id=""content"" bordercolordark=""#C0C0C0"" bordercolorlight=""#C0C0C0"" bordercolor=""#C0C0C0""><b><center>Details</b></center></td>"
  Response.Write "</tr>"

  FOR i=0 to ubound(theHistoricList,2) 'Loop through all selected rows/records
	Response.Write "<tr>"

	If (len(theHistoricList(1,i))>0) then
	 Response.Write "<td><center>" & theHistoricList(1,i) & "</center></td>"
	Else
	 Response.Write "<td><center>-</center></td>"
	End If

	If (InStr(theHistoricList(6,i),",")>0) then
	 Response.Write "<td><center>" & left(theHistoricList(6,i),InStr(theHistoricList(6,i),",")-1) & "</center></td>"
	Else
	 Response.Write "<td></center>" & theHistoricList(6,i) & "</center></td>"
	End If

	 Response.write "<td align=""center""><a href=""#"" onClick=""sendQuery('" & theHistoricList(0,i) & "','zoom1');""><img src=""../images/zoomin_1.gif""  border=""0"" Title=""Zoom to Map#" & theHistoricList(0,i) & """></a></td>"

	 Response.write "<td align=""center""><input type=""button"" value=""More"" style=""background-color: #3180CB;color: #FFFFFF;font-size: 8pt;width:40""  onClick=""sendQuery('" & theHistoricList(0,i) & "');"" Title=""Details for Map#" & theHistoricList(0,i) & """></td>"

	Response.Write "</tr>"
  NEXT

  Response.Write "</table></center>"

  

 ELSE '***Create FULL DESCRIPTION PAGE*******************************

  Response.Write "<center><table width=""95%"" border=""0"" cellspacing=""0"" id=""content"">"

  Response.Write "<tr><td><a href=""http://www.tacomaculture.org"" target=""_blank""><img border=""0"" alt=""tacoma culture"" src=""../images/logo.gif"" width=""206"" height=""40""></a></td></tr>"

  Response.Write "<tr>"
  Response.Write "<td>"

 'Count number of records in each property category

  FOR i=0 to ubound(theHistoricList,2) 'Loop through all selected rows/records

	'Property Picture
	If (not isNull(theHistoricList(12,i))) then
  Response.Write "</td></tr><tr height=""150""><td>"
	  Response.Write "<center><a href=""#"" onClick=""javascript:popWin=open('http://wspdsmap.cityoftacoma.org/website/Historic/Photos/" & theHistoricList(12,i) & "','popupWindow','scrollbars=yes,resizable=yes,width=450,height=300,top=10,left=10');popWin.focus();return false;""><img src='../../Historic/Photos/" & theHistoricList(12,i) & "' width=""150"" alt=""Click to enlarge""></a>"
	  Response.Write "</center>"
  Response.Write "</td></tr><tr><td>"
	Else
	 Response.Write "Sorry, no picture available.<br><br>"
	End If
%>

<a href="#null" onClick="seeBirdsEye('<%=formatString(theHistoricList(1,i))%>', '<%=formatString(left(theHistoricList(6,i),InStr(theHistoricList(6,i),",")-1))%>', '<%=theLat%>', '<%=theLong%>')">  Bird's Eye Photo (Microsoft Virtual Earth)</a><br><br>

<a href="#null" onClick="seeStreetView('<%=formatString(theHistoricList(1,i))%>', '<%=formatString(left(theHistoricList(6,i),InStr(theHistoricList(6,i),",")-1))%>', '<%=theLat%>', '<%=theLong%>')">  Street View Photo (Google Maps)</a><br><br>

<%

	Response.Write "<b>Scan#: </b>" & theHistoricList(0,i) & "<br>"
	Response.Write "<b>Historic Name: </b>" & theHistoricList(1,i) & "<br>"
	Response.Write "<b>Property Address: </b>" & theHistoricList(6,i) & "<br>"
	Response.Write "<b>Tax No./Parcel No.: </b>" & theHistoricList(2,i) & "<br><br>"
	Response.Write "<b>Date Recorded: </b>" & theHistoricList(7,i) & "<br>"
	Response.Write "<b>Style: </b>" & theHistoricList(8,i) & "<br>"
	Response.Write "<b>Form/Type: </b>" & theHistoricList(9,i) & "<br>"
	Response.Write "<b>Date of Construction: </b>" & theHistoricList(3,i) & "<br>"
	Response.Write "<b>Architect: </b>" & theHistoricList(4,i) & "<br>"
	Response.Write "<b>Builder: </b>" & theHistoricList(5,i) & "<br>"
	Response.Write "<b>Significance: </b>" & theHistoricList(10,i) & "<br><br>"
	Response.Write "<b>Appearance: </b>" & theHistoricList(11,i) & "<br><br>"

  Response.Write "</td></tr><tr><td>"

	'Links & Map

	 googleMap
	 
	 Response.write "<a href=""#"" onClick=""javascript:window.open('directions.asp?Address=" & formatString(theHistoricList(6,i)) & "&Name=" & formatString(theHistoricList(1,i)) & "','popupWindow','scrollbars=yes,resizable=yes,width=640,height=480,top=10,left=10');return false;""><b>Driving Directions</b></a><br>"
	 Response.write "<a href=""#"" onclick=""sendAddress(" & chr(39) & theHistoricList(6,i) & chr(39) & ");return false;""><b>Photos (Tacoma Public Library)</b></A><br>"
	 Response.write "<a href=""#"" onclick=""sendAPN(" & chr(39) & theHistoricList(2,i) & chr(39) & ");return false;""><b>Pierce County Assessor Data </b></A>"

  NEXT

  Response.Write "</td></tr></table></center>"

 End If  'Check for records

 theFooter

END IF 'check for summary page



'================================
'Clean Up, Other subs & functions
'================================

  Set theHistoricList = Nothing

Sub CleanUp
    objConn.Close 
    set objConn  = nothing
    set theAList = nothing
    set theNList = nothing
End Sub


Sub selectMap
 IF (iMap <> "") THEN  'Don't need to update map
    Response.Write vbcr & "<BODY>"
 ELSE   'Zoom out to City or single property
    Response.Write vbcr & "<BODY onload=" & chr(34) & "sendQuery('" & theQuery & "');" & chr(34) & ">" & vbcr
 END IF
End Sub


Sub theFooter
  Response.Write "</body>"
  Response.Write "</html>"
End Sub


Function formatString(strText)
  if (strText <>"") then
    aryChars = Array(chr(34),"'","\","&")
    aryReplacement = Array("","","/","%26")

     FOR x = 0 To UBound(aryChars)
      strText = Replace(strText,aryChars(x), aryReplacement(x))
     NEXT

    formatString = strText
  else
    formatString = ""
  end if
End Function

Function theTitle(x)
    if (theCriteria2="") then
     theCriteria2 = theMString(x) & iItem(x) & "<br>"
    else
     theCriteria2 = theCriteria2 & theMString(x) & iItem(x) & "<br>"
    end if
End Function

Sub googleMap
    Response.Write "<div align=center>"
    Response.Write "<img src='http://maps.google.com/maps/api/staticmap?center="
    Response.Write iMap
    Response.Write "&zoom=18&size=450x450"
    Response.Write "&maptype=hybrid"
    Response.Write "&markers=color:blue|"
    Response.Write iMap
    Response.Write "&sensor=false"
    Response.Write "' border=1></center>"
    Response.Write "</div>"
End Sub

%>