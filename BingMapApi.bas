Attribute VB_Name = "BingMapApi"
' REF: https://syntaxbytetutorials.com/excel-function-to-calculate-distance-using-google-maps-api-with-vba/
' Get key from key vault
Private Const apiKey As String = "{API Key}"
Function TravelTime(origin, destination)
    Dim parsed As Dictionary
    Set parsed = GetTravelRoute(origin, destination)
    ' get travel time in minutes
    TravelTime = parsed("resourceSets")(1)("resources")(1)("travelDuration") / 60
End Function

Function TravelDistance(origin, destination)
    Dim parsed As Dictionary
    Set parsed = GetTravelRoute(origin, destination)
    TravelDistance = parsed("resourceSets")(1)("resources")(1)("travelDistance")
End Function

Function GetTravelRoute(origin, destination)
    Dim strUrl As String
    ' REF https://docs.microsoft.com/en-us/bingmaps/rest-services/routes/calculate-a-route
    ' GET http://dev.virtualearth.net/REST/v1/Routes?wayPoint.1={wayPoint1}&waypoint.2={waypoint2}&maxSolutions={maxSolutions}&distanceUnit={distanceUnit}&key={BingMapsKey}
    strUrl = "http://dev.virtualearth.net/REST/v1/Routes?maxSolutions=1&distanceUnit=Mile&wayPoint.1=" & origin & "&waypoint.2=" & destination & "&key=" & apiKey
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    With httpReq
        .Open "GET", strUrl, False
        .Send
    End With
    Dim response As String
    response = httpReq.ResponseText
    Dim parsed As Dictionary
    Set parsed = JsonConverter.ParseJson(response)
    Set GetTravelRoute = parsed
End Function

Function GeoCoordinates(state, city, address)
    Dim strUrl As String
    ' REF https://docs.microsoft.com/en-us/bingmaps/rest-services/locations/find-a-location-by-address
    ' GET http://dev.virtualearth.net/REST/v1/Locations/US/{adminDistrict}/{locality}/{addressLine}?includeNeighborhood={includeNeighborhood}&include={includeValue}&maxResults={maxResults}&key={BingMapsKey}
    strUrl = "http://dev.virtualearth.net/REST/v1/Locations/US/" + state + "/" + city + "/" + address + "?maxResults=1&key=" & apiKey
    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    With httpReq
        .Open "GET", strUrl, False
        .Send
    End With
    Dim response As String
    response = httpReq.ResponseText
    Dim parsed As Dictionary
    Set parsed = JsonConverter.ParseJson(response)
    Dim latitude As String
    Dim longitude As String
    Dim point As String
    latitude = parsed("resourceSets")(1)("resources")(1)("point")("coordinates")(1)
    longitude = parsed("resourceSets")(1)("resources")(1)("point")("coordinates")(2)
    point = latitude + "," + longitude
    GeoCoordinates = point
End Function
