
<%@ Language=VBScript %>

<%
'*****************************************************
'NAME: Calendar.asp
'AUTHOR: Mike Murnane
'LAST UPDATE: 8/23/2018

'DESCRIPTION: Edits calendar file (ics)
'             
'REQUIRES: 
' Request sent with calendar values (ics,etc)

'RESULTS: 
' 

'****************************************************
%>


<html><head>
 <meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
 </head>
 <body bgcolor="#FFFFFF">Open file (then save) to add appointment to your calendar.

<% 

'----------------------------------------------------------------------------
'Variables 
'----------------------------------------------------------------------------
'Sent by user
  Set theDay = request.querystring("Day")
  Set theStudio = request.querystring("Studio")
  Set theAddress = request.querystring("Address")
  Set theName = request.querystring("Name")
  
  'format address string spaces for Google link - don't use Set
  theGAddress = Replace(request.querystring("Address"), " ","%20")

'Variables for calendar ics text file - user must have permission to write to file

'basic template depends on the day
  If (theDay=1) Then
    fileString = Server.MapPath("StudioTourSat.ics")
  ElseIf (theDay=2) Then
    fileString = Server.MapPath("StudioTourSun.ics")
  Else
    fileString = Server.MapPath("StudioTourSS.ics")
  End If
  
  newfileString = Server.MapPath("TacomaStudioTour.ics")

'-------------------------------------------------------------------------------------------


  Const ForReading = 1
  Const ForWriting = 2
  Const ForAppending = 8
  
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFSO.OpenTextFile(fileString, ForReading)
  
  'Open existing calendar file and use for new string value
   strText = objFile.ReadAll
   objFile.Close

  'Replace generic items with specific values in string ------------------------
  'Studio
   strNewText = Replace(strText, "Studio #", "Studio #" + theStudio)
  
  'Continue to update strNewText
  'Address
   strNewText = Replace(strNewText, "Address:", "Address: " + theAddress)
  
  'Artist Name
   strNewText = Replace(strNewText, "Artist Name:", "Artist Name: " + theName)
  
  'Google Directions
   strNewText = Replace(strNewText, "Google Directions:", "Google Directions: https://maps.google.com/maps?daddr=" + theGAddress + ",%20Tacoma,%20WA")
   
  '-----------------------------------------------------------------------------
  
  'Overwrite existing calendar file with new string
   Set objFile = objFSO.OpenTextFile(newfileString, ForWriting)
   objFile.WriteLine strNewText
   objFile.Close


'Show text file (ics) - TIME added to avoid any browser caching problems
'(webcal: if you want to open mail program on client machine)
  response.write "<BODY onLoad=""javascript:location='TacomaStudioTour.ics?i=" & TIME & "'"">"

'Clean up variables
  Set objFSO = Nothing
  Set objFile = Nothing

%>


</body></html>


