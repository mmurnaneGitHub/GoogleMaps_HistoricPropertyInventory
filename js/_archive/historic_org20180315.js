//ArcGIS Server Link for Google Maps API V3
//Help: http://static.neoreef.com/common/js/libs/google/maps/v3/arcgislink/docs/examples.html
//Help Examples: https://htmlpreview.github.io/?https://raw.githubusercontent.com/googlemaps/v3-utility-library/master/arcgislink/docs/examples.html
//Help Examples: http://static.neoreef.com/common/js/libs/google/maps/v3/arcgislink/docs/examples.html
//Class References: https://htmlpreview.github.io/?https://raw.githubusercontent.com/googlemaps/v3-utility-library/master/arcgislink/docs/reference.html
//Code from here: https://github.com/googlemaps/v3-utility-library/tree/master/arcgislink/examples 

var map, geocoder, qLayer, qParams, svc, svc2;
var markers = [];  //Array for map markers
var icon;  //Selete markers - add icon (marker) details on init
var buffers = [];  //Buffer array
var bufferDistance = null; //updated from addmatch.htm - opened from buffer tool

function init() {
 google.maps.LatLng.prototype.latRadians = function(){ return (Math.PI * this.lat()) / 180; }  //Extend v3 to have same commands from v3 (latRadians)
 google.maps.LatLng.prototype.lngRadians = function(){ return (Math.PI * this.lng()) / 180; } //Extend v3 to have same commands from v3 (lngRadians)
 var mapOptions = {
    zoom: 11,
    center: new google.maps.LatLng(47.24000, -122.45000), //  47.2280, -122.4413
    mapTypeId: google.maps.MapTypeId.TERRAIN,
    draggableCursor: 'pointer', // every pixel is clickable.
    streetViewControl: true 
  };
  map = new google.maps.Map(document.getElementById("map_canvas"), mapOptions);
  var url = 'https://gis.cityoftacoma.org/arcgis/rest/services/PDS/H_Inventory/MapServer';  //Historic layers (dynamic layer) - base64 imageData instead of href link 
  var dynamap = new gmaps.ags.MapOverlay(url, { opacity: 0.75 });
      //START HERE - NULL VALUE, NOT ADDING TO MAP BELOW
      //dynamap.setMap(map); //add dynamic layer to map

      google.maps.event.addListener(map, 'zoom_changed', function() {
        var zoom_level = map.getZoom();  //change background map with zoom level
        if (zoom_level > 16) {
              map.setMapTypeId(google.maps.MapTypeId.HYBRID);
        } else if (zoom_level > 13) {
              map.setMapTypeId(google.maps.MapTypeId.ROADMAP);
        } else {
              map.setMapTypeId(google.maps.MapTypeId.TERRAIN);
        } 
      });

	    //Selected marker icon
	    icon = {
	      url: 'images/mapIcons/magenta.png',
	      size: new google.maps.Size(14, 14), 
	      origin: new google.maps.Point(0, 0),
	      anchor: new google.maps.Point(7, 7)
	    };

      geocoder = new google.maps.Geocoder();  //Geocode with Google Maps Geocoder
      qLayer = new gmaps.ags.Layer(url + '/0');  //Query layer
      qParams = {returnGeometry: true, outFields: ["ID"]};  //Query parameters
      svc = new gmaps.ags.MapService(url);  //Identify web service
	  svc2 = new gmaps.ags.GeometryService("https://gis.cityoftacoma.org/arcgis/rest/services/Utilities/Geometry/GeometryServer");  //Geometry web service (buffer)
      google.maps.event.addListener(map, 'click', identify); //Run Identify (or buffer) on map click

     //identify proxy page to use if the toJson payload to the geometry service is greater than 2000 characters.
      //If this null or not available the buffer operation will not work.  Otherwise it will do a http post to the proxy.
      //**Need to create application for proxy directory in IIS 
      gmaps.ags.Config.proxyUrl = "proxy/proxy.ashx";   //Proxy - Version 1.1.2 - https://github.com/Esri/resource-proxy/releases
      gmaps.ags.Config.alwaysUseProxy = false;

}

function zoomTacoma() {
  map.setCenter(new google.maps.LatLng(47.24000, -122.45000));
  map.setZoom(11);
}

function findAddress(address1) {
  deleteMarkers(); //cleanup previous address
  var address = address1 + ", Tacoma, WA";  //for beter geocoding results - https://developers.google.com/maps/documentation/javascript/examples/geocoding-simple
      geocoder.geocode({'address': address}, function(results, status) {
        if (status === 'OK') {
          map.setCenter(results[0].geometry.location);
          map.setZoom(17);  //zoom to level just above oblique
          var marker = new google.maps.Marker({
            map: map,
            title: address1,
            position: results[0].geometry.location
          });
          markers.push(marker);  //Add marker to the array.
          identify(results[0].geometry.location)//Identify historic info at address location
        } else {
          alert('Geocode was not successful for the following reason: ' + status);
        }
      });
}

function deleteMarkers() {
    for (var i = 0; i < markers.length; i++) {
        markers[i].setMap(null);  //Loop through all the markers and remove
    }
    markers = [];  //reset array
};

function identify(evt) {
  if (evt.hasOwnProperty("latLng")){  //true for any click, false for geocoding
    theLocation = evt.latLng  //map click location
    deleteMarkers(); //remove any existing markers (geocoding too fast, need a defer, for now remove only during map clicks)
  } else {
    theLocation = evt;  //geocoded location
  }

	if (newBuffer){ //Buffer tool active: buffer page selected and Select tool active (not Identify or Pan)
	    doBuffer(theLocation);
	} else {  //Pan or Identify tool active (even if buffer page is visible)

	  svc.identify({
	    'geometry': theLocation,
	    'tolerance': 4,  //tolerance needed for points & around selection markers
	    'layerIds': [0],
	    'layerOption': 'all',
	    'bounds': map.getBounds(),
	    'width': map.getDiv().offsetWidth,
	    'height': map.getDiv().offsetHeight
	  }, function(results, err) {
	    if (err) {
	      alert(err.message + err.details.join('\n'));
	    } else {
	      addResults(results);  
	    }
	  });

	}

}

function doBuffer(point) {
    getBoundaries(point, bufferDistance); //zoom map to buffer - distance updated from addmatch.htm - page opened from buffer tool
    var overlayOptions = {
        strokeColor: "#43AAFE", strokeWeight: 4, strokeOpacity: 0.55,
        fillColor: "#327C5A", fillOpacity: 0.35, clickable:false
    };
	for (var i = 0; i < buffers.length; i++) {
	  buffers[i].setMap(null);  //Cleanup any existing buffer
	}
	buffers = [];  //reset array
	var params = {  //buffer parameters
	  geometries: [point],
	  bufferSpatialReference: 2927,  //spatial reference of inventory points
	  distances: [bufferDistance],  //updated from addmatch.htm - page opened from buffer tool
	  unit: 9035,  //esriSRUnit_SurveyMile - US survey mile
	  unionResults: true, 
	  overlayOptions: overlayOptions
	};

    svc2.buffer(params, function(results, err) {
      if (!err) {
        var g;
        for (var i = 0; i < results.geometries.length; i++) {  
          for (var j = 0; j < results.geometries[i].length; j++) {
            g = results.geometries[i][j];
            g.setMap(map); //add buffer to map
            buffers.push(g);
          }
        }
        qParams.geometry = buffers;  //update query geometry (buffer) 
        qLayer.query(qParams, processResultSet2);  //Query for records inside buffer

      } else {
        alert(err.message + err.details.join(','));
      }
    });
}

function getBoundaries(centrePt, radius) {
 var hypotenuse = Math.sqrt(2 * radius * radius);
 var sw = getDestLatLng(centrePt, 225, hypotenuse);
 var ne = getDestLatLng(centrePt, 45, hypotenuse);
 map.fitBounds(new google.maps.LatLngBounds(sw, ne));  //fit map to new bounds
}

function getDestLatLng(latLng, bearing, distance) {
 var EARTH_RADIUS = 3963; //in miles 
 var lat1 = latLng.latRadians();
 var lng1 = latLng.lngRadians();
 var brng = bearing*Math.PI/180;
 var dDivR = distance/EARTH_RADIUS;
 var lat2 = Math.asin( Math.sin(lat1)*Math.cos(dDivR) + Math.cos(lat1)*Math.sin(dDivR)*Math.cos(brng) );
 var lng2 = lng1 + Math.atan2(Math.sin(brng)*Math.sin(dDivR)*Math.cos(lat1), Math.cos(dDivR)-Math.sin(lat1)*Math.sin(lat2));
 return new google.maps.LatLng(lat2/ Math.PI * 180, lng2/ Math.PI * 180);
}

function addResults(response) {
  // aggregate the result per map service layer
  var idResults = response.results;
  var layers = { "0": [] };
  var thePropText = ""; 

  for (var i = 0; i < idResults.length; i++) {
    var result = idResults[i];
    layers[result.layerId].push(result);
  }

  for (var layerId in layers) {  // get field values for each map service layer
    var results = layers[layerId];
    var count = results.length;
    switch(layerId) {
      case "0":
        if (count == 0) {
          thePropText = "Sorry, no inventory found.";
           break;
        }

        var ID_Builder = [];  //array for IDs
        for (var i = 0; i < count; i++) { 
          if (i == 0) {
              ID_Builder.push(results[i].feature.attributes['ID']);
          } else {
              ID_Builder.push("," + results[i].feature.attributes['ID']);
          }
        }
        var theIDs = ID_Builder.join("");
 
        break;
     }
  }

  if (thePropText == "Sorry, no inventory found.") {
      parent.TextFrame.document.open();
      parent.TextFrame.document.writeln('<html><HEAD><link href="css/master.css" rel="stylesheet"></HEAD><center><b><br>Sorry, no inventory found.</html>');
  } else {
      parent.TextFrame.location='scripts/summary.asp?ID=(' + theIDs + ')';  //send query to summary page
  }
}

function executeQuery(ID) {	
    deleteMarkers();  //Clean existing markers off map
    var params = {  //Query parameters
      returnGeometry: false,
      where: "ID IN (" + ID + ")",
      outFields: ["ID","LAT","LONG"]
    };
    qLayer.query(params, processResultSet);  //run query
}

function processResultSet(fset) {
  //Identify results
  var bounds = new google.maps.LatLngBounds();
  for (var i = 0; i < fset.features.length; i++) {
    label = fset.features[i].attributes["ID"].toString();  //convert id number to string
    id = fset.features[i].attributes["ID"];
    var point = new google.maps.LatLng(fset.features[i].attributes["LAT"], fset.features[i].attributes["LONG"]);
        bounds.extend(point);
    var marker = createMarker(point,label);  // create the marker
        markers.push(marker); //add createMarker results to marker array - add to map when loop done
  } 

  //Set map extent from bounds of selection set
    map.fitBounds(bounds);
      	if (map.getZoom() > 17) {
      		map.setZoom(17);  // Limit zoom level
      	}
  //Send APN query to TextFrame if one parcel & not just zooming
  if (fset.features.length==1) {
     window.open('scripts/summary.asp?ID=(' + label + ')&map=' + point);     
  }
  
}

function processResultSet2(fset) {
	//Buffer query results
	var features = fset.features;
	 if (features.length>0){
	    var ID_Builder = [];  //array for IDs
	    for (var i = 0; i < features.length; i++) { 
	      if (i == 0) {
	          ID_Builder.push(features[i].attributes['ID']);
	      } else {
	          ID_Builder.push("," + features[i].attributes['ID']);
	      }
	    }
	    var theIDs = ID_Builder.join("");
	        postURL('/website/HistoricMap/scripts/summary.asp?ID=(' + theIDs + ')', false); //only one variable (false), post (instead of get) the url to get around the 2048 character limit
	 } else {
	      parent.TextFrame.document.open();
	      parent.TextFrame.document.writeln('<html><HEAD><link href="css/master.css" rel="stylesheet"></HEAD><center><b><br>Sorry, no inventory found.</html>');
	 }
};

function postURL(url, multipart) {
/**
 * Takes a URL and goes to it using the POST method.
 * @param {string} url  The URL with the GET parameters to go to.
 * @param {boolean=} multipart  Indicates that the data will be sent using the
 *     multipart enctype.
 */
 //http://cwestblog.com/2012/11/21/javascript-go-to-url-using-post-variable/
  var form = document.createElement("FORM");
  form.method = "POST";
  if(multipart) {
    form.enctype = "multipart/form-data";
  }
  form.style.display = "none";
  document.body.appendChild(form);
  form.action = url.replace(/\?(.*)/, function(_, urlArgs) {
    urlArgs.replace(/\+/g, " ").replace(/([^&=]+)=([^&=]*)/g, function(input, key, value) {
      input = document.createElement("INPUT");
      input.type = "hidden";
      input.name = decodeURIComponent(key);
      input.value = decodeURIComponent(value);
      form.appendChild(input);
      parent.TextFrame.document.body.appendChild(form); //put into correct frame (otherwise ignore this line)
    });
    return "";
  });
  form.submit();
}

function createMarker(point,name) {
  var marker = new google.maps.Marker({
    position: point,
    map: map,
    icon: icon,
    title: name,
    zIndex: 99
  });
  return marker;
}

function getReady(buffer) {
  //Called from sumary.asp & addmatch.htm
  deleteMarkers();  //Clean existing markers off map
	for (var i = 0; i < buffers.length; i++) {
	  buffers[i].setMap(null);  //Cleanup any existing buffer
	}
	buffers = [];  //reset array

  //reset buffer flag
  if (buffer=="true"){
      newBuffer="true";  //buffer tool selected
  } else {
      newBuffer="";  //pan or identify tool selected
  }
  
}

window.onload = init;  //initialize the map