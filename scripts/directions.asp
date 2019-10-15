<%@ Language=VBScript %>

<%
'*****************************************************
'Name: directions.asp

'Description: Input form for getting driving directions to public art  
'              

'Requires: Address & Art Title sent in URL string
'          

'Called by: art.asp & summary.asp

'Results: MapQuest driving directions & map

'Last Update: 5/27/04

'****************************************************

'==========
'Variables
'==========
  Dim iAddress, iName

  iAddress = Request.QueryString("Address")  
  iName = Request.QueryString("Name")  

%>


<html>
<head>
	<title>Driving Directions</title>
</head>
<body onload="document.getElementById('1a').focus();">
<font face=verdana> <font size=-2>

<form name="dirForm" action="http://www.mapquest.com/directions/main.adp" method="get">
<input type="hidden" name=go value=1>
<input type="hidden" name=do value=nw>
<input type="hidden" name=ct value="NA">
<table cellspacing="0" cellpadding="10" style="border: 1px solid #330066; background-color: #F4F7F4;" >
  <tr> 
    <td>
<table width="390" cellpadding=6 cellspacing=0 border=0 bgcolor="#fafad2" style="border: 1px solid #330066; padding: 6px; margin-top: 4px">
	<tr>
	<td>
	<table border="0" cellpadding="1" cellspacing="0" class="text1" width="100%"> 
  <tr> 
    <td align=left colspan=3 class="title4"><b>Enter Your Start Address</b></td>
  </tr>
  <input type=hidden name=1y value=US>
  <tr> 
    <td colspan=3> Address/Intersection: </td>
  </tr>
  <tr> 
    <td colspan=3> 
      <input type="text" id=1a name=1a size=28 style="width:361" maxlength=80 value="747 Market St" class="inputColor">
    </td>
  </tr>
  <tr> 
    <td valign=bottom>  City:</td>
    <td valign=bottom>State:</td>
    <td valign=bottom>Zipcode:</td>
  </tr>
  <tr> 
    <td width=201 height="29" valign="top"> 
     <input type="text" name=1c size=11 maxlength=80 style="width:160" value="Tacoma" class="inputColor">
    </td>
    <td height="29" valign="top"> 
      <input type="text" name=1s maxlength=2 size=1 style="width:28" value="WA" class="inputColor">
    </td>
    <td height="29" valign="top"> 
      <input type="text" name=1z size=6 maxlength=10 style="width:95" value="" class="inputColor">
    </td>
  </tr>
  <tr>
    <td colspan="3" align="center">
      <input type="submit" value="Get Directions!" name="submit" class="btnGrey" class="inputColor">
    </td>
  </tr>
</table>
</td>
</tr>
</table>

<table width=376 cellpadding=6 cellspacing=0 border=0 bgcolor="#fafad2" style="border: 1px solid #330066; padding: 6px; margin-top: 4px">
	<tr>
	<td>
	<input type="hidden" name=2a value="<% Response.Write iAddress %>">
	<input type="hidden" name=2c value="Tacoma">
	<input type="hidden" name=2s value="WA">
	<input type="hidden" name=2z value="">
<table width="376" cellpadding="3" cellspacing="0" border="0" bgcolor="#fafad2" class="text1"> 
  <tr> 
    <td class="title4" width="20%"><b>Destination: </b></td>
		<td class="title4"><% Response.Write iName %></td>
  </tr>
  <input type=hidden name=1y value=US>
  <tr> 
    <td class="text1" valign="top">&nbsp;</td>
		<td class="text1"><% Response.Write iAddress %><br>
											Tacoma, WA</td>
  </tr>
</table>

</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</body>
</html>
