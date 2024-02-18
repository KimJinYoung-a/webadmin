<%@ language=vbscript %>
<% option explicit %>
<%
Response.CharSet = "euc-kr"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
dim lecid : lecid=RequestCheckvar(request("lecid"),32)
dim selectBoxName : selectBoxName=RequestCheckvar(request("selectBoxName"),100)
dim retVal
dim sqlStr, simplepushid

sqlStr = " select regidx,deviceid,appkey,regdate,lastupdate,appver,isusing,pushyn  from [db_academy].[dbo].[tbl_app_regInfo]  "
sqlStr = sqlStr & " where userid in ('fingertest01','fingertest02','fingertest03','fingertest04','fingertest05','thefingers01') "
sqlStr = sqlStr & " and userid='"&lecid&"'"
sqlStr = sqlStr & " order by lastUpdate desc"


rsACADEMYget.Open sqlStr,dbACADEMYget,1

if  not rsACADEMYget.EOF  then
   do until rsACADEMYget.EOF
       simplepushid = rsACADEMYget("deviceid")
       if (LEN(simplepushid)>64) then
            simplepushid = LEFT(simplepushid,28)&"......"&RIGHT(simplepushid,28)
       end if
       retVal = retVal&"<option value='"&rsACADEMYget("appkey")&"|"&rsACADEMYget("deviceid")&"' >["&rsACADEMYget("appkey")&"]["&rsACADEMYget("appver")&"] - ["&rsACADEMYget("lastupdate")&"]"&simplepushid&" </option>"
       rsACADEMYget.MoveNext
   loop
end if
rsACADEMYget.close

if (retVal<>"") then
    retVal="<select class='select' name='"&selectBoxName&"' >"&retVal&"</select>"
end if
  
   
response.write(retVal)
   

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->