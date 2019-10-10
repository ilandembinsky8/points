<!-- #include file = "secure.inc" -->
    <%
    sql = "select top 50 * from [points_nadlan].[dbo].[nadlan]"
    if request("order")<>"" then sql = sql & " order by "&request("order")
    r.open sql,strconn,1,3
    Response.write "<table cellspacing=0 cellpadding=4><tr style='background-color:#EAEAEA'>"
    For each item in r.Fields
     response.write "<th><a href='showData.asp?table="&theTable&"&order="&item.Name&"'>" & item.Name & "</a></th>"
    Next

    Response.write "</tr>"
    i=1
    While Not r.EOF
     if i mod 2 = 0 then 
        clr="#FAFAFA"
     else
        clr="#DADADA"
     end if
     response.write "<tr style='background-color:"&clr&"'>"
     For each item in r.Fields
      Response.write "<td>" & r(item.Name) & "</td>"
     Next
     response.write "</tr>"
     r.MoveNext
     i=i+1
    Wend
    response.write "</table>"
    r.close
    %>