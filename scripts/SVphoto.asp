<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
    <title>Street View Photo (Google Maps) </title>
	<meta http-equiv="Content-Type" content="text/html/xml; charset=utf-8">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<META HTTP-EQUIV="Expires" CONTENT="-1">



    <%@ Language=VBScript %>

    <% 'Get variables sent by user
     	Set theTitle =request.querystring("title")  
     	Set theDetails =request.querystring("details") 
     	Set theLat =request.querystring("lat") 
     	Set theLong =request.querystring("long")

    	'Other variables:

	'Very small increment to add to URL Lat and avoid IE cache problems (XML)
    	 theTime = Second(Now()) 
 
     %>
 
</head>

<body bgcolor="black">

 <iframe width="600" height="230" frameborder="0" scrolling="no" marginheight="0" marginwidth="0" src="http://maps.google.com/maps/sv?cbp=12,0,,0,5&amp;cbll=<%=theLat%><%=theTime%>,<%=theLong%>&amp;v=1&amp;panoid=&amp;gl=&amp;hl=en">
 </iframe>

 <FONT style="FONT-SIZE: 12px" face=arial color=#cccccc size=2>

 <p>
  <b><font color=#00CCFF>INSTRUCTIONS:</font></b> Click on the photo and then navigate 360&deg; with either the arrow keys on the keyboard or the mouse.  &nbsp;Initial view is facing north.
 </p>
 
 <p>
  <b><font color=#00CCFF>LOCATION:</font></b> Closest photo to <i> <%=theTitle%> </i>. 
  <br><%=theDetails%> 
 </p> 
 

 <a id="cbembedlink" href="http://maps.google.com/maps?cbp=12,0,,0,5&cbll=<%=theLat%>,<%=theLong%>&ll=<%=theLat%>,<%=theLong%>&layer=c" style="color:#00FF00;text-align:left">View Larger Map</a>

 </FONT> 

</body>

</html>
 	