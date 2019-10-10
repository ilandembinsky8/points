

<!-- #include file = "secure.inc" -->
<!-- #include file = "aspJSON1.17.asp" -->
<%
Set oJSON = New aspJSON
oJSON.loadJSON(request.form("json"))
r.open "[points_nadlan].[dbo].[nadlan]",strconn,2,3
r.addnew
r("area")=""&oJSON.data("area")&""
r("gush_helka")=""&oJSON.data("gush_helka")&""
r("iska_date")=""&oJSON.data("iska_date")&""
r("CITY")=""&oJSON.data("CITY")&""
r("STREET")=""&oJSON.data("STREET")&""
r("NUMBER")=oJSON.data("NUMBER")
r("ENTERANCE")=""&oJSON.data("ENTERANCE")&""
r("appartment")=oJSON.data("APPARTMENT")
r("TPRICENIS")=oJSON.data("TPRICENIS")
r("TPRICEDOLLAR")=oJSON.data("TPRICEDOLLAR")
r("EPRICENIS")=oJSON.data("EPRICENIS")
r("EPRICEDOLLAR")=oJSON.data("EPRICEDOLLAR")
r("ARNONASURFACE")=oJSON.data("ARNONASURFACE")
r("ROOMS")=oJSON.data("ROOMS")
r("FLOOR")=oJSON.data("FLOOR")
r("STOREHOUSE")=oJSON.data("STOREHOUSE")
r("BUILDYEAR")=oJSON.data("BUILDYEAR")
r("FLOORS")=oJSON.data("FLOORS")
r("YARD")=oJSON.data("YARD")
r("PRICEPERROOM")=oJSON.data("PRICEPERROOM")
r("PARKING")=""&oJSON.data("PARKING")&""
r("GALLERY")=oJSON.data("GALLERY")
r("ELEVATOR")=""&oJSON.data("ELEVATOR")&""
r("DEALTYPE")=""&oJSON.data("DEALTYPE")&""
r("BUILDINGFUNC")=""&oJSON.data("BUILDINGFUNC")&""
r("UNITFUNC")=""&oJSON.data("UNITFUNC")&""
r("SHUMAPARTS")=""&oJSON.data("SHUMAPARTS")&""
r("GUSHMOFA")=oJSON.data("GUSHMOFA")
r("TABA")=""&oJSON.data("TABA")&""
r("MAHUT")=""&oJSON.data("MAHUT")&""
r.update
%>