<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
  <title>Bird's Eye View (Microsoft Virtual Earth) </title>

  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <script type="text/javascript" src="http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=6.1"></script>

  <script type="text/javascript">

     <%@ Language=VBScript %>

    //Get variables sent by user
    <% Set theTitle =request.querystring("title") %> 
    <% Set theDetails =request.querystring("details") %> 
    <% Set theLat =request.querystring("lat") %> 
    <% Set theLong =request.querystring("long") %> 
 
    //Other variables     
    var obMap = null;      
    var title = "<%=theTitle%>";
    var details = "<%=theDetails%>";
    var myLat = <%=theLat%>;
    var myLong = <%=theLong%>;

    var myLatLong = new VELatLong(myLat,myLong);
  
    function GetMap()
      {
	//Add map
	  obMap = new VEMap('birdMap');
	  obMap.LoadMap(myLatLong, 1, VEMapStyle.BirdseyeHybrid);
	
	 //Add pushpin to the map
		var pin = new VEShape(VEShapeType.Pushpin,myLatLong);
		pin.SetTitle(title);
		pin.SetDescription("<br>" + details);
		obMap.AddShape(pin);
      }
      
  
  </script>

</head>

<body onload="GetMap();" bgcolor="black">
	<div id='birdMap' style="position:relative;"></div>		
</body>

</html>
 	