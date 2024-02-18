<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx, BGColor, Evt_Type, Evt_Code, Evt_Title, Evt_Discount, Evt_Subcopy
Dim Item1, Item2, Item3, StartDate, EndDate, DispOrder, Isusing, mode
	
	
	idx = requestCheckVar(request("idx"),10)
	BGColor = requestCheckVar(request("BGColor"),10)
	Evt_Type = requestCheckVar(request("Evt_Type"),10)
	Evt_Code = requestCheckVar(request("Evt_Code"),10)
	Evt_Title = requestCheckVar(request("Evt_Title"),128)
	Evt_Discount = requestCheckVar(request("Evt_Discount"),10)
	Evt_Subcopy = requestCheckVar(request("Evt_Subcopy"),128)
	Item1 = requestCheckVar(request("Item1"),10)
	Item2 = requestCheckVar(request("Item2"),10)
	Item3 = requestCheckVar(request("Item3"),10)
	StartDate = requestCheckVar(request("StartDate"),10)
	EndDate = requestCheckVar(request("EndDate"),10)
	DispOrder = requestCheckVar(request("DispOrder"),3)
	Isusing = requestCheckVar(request("Isusing"),1)

	if idx="" then idx=0
	If idx=0 Then
	mode = "add"
	Else
	mode = "edit"
	End If

dim sqlStr


if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_main_enjoy_event]" + VbCrlf
    sqlStr = sqlStr + " (BGColor, Evt_Type, Evt_Code, Evt_Title, Evt_Subcopy, Evt_Discount, Item1, Item2, Item3, StartDate, EndDate, DispOrder, Isusing, RegUser)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + CStr(BGColor) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Type + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Code + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Title + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Subcopy + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Discount + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Item1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Item2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Item3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + StartDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + EndDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_main_enjoy_event]') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_main_enjoy_event]" + VbCrlf
   sqlStr = sqlStr + " set BGColor='" + CStr(BGColor) + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Type='" + Evt_Type + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Code='" + Evt_Code + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Title='" + Evt_Title + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Discount='" + Evt_Discount + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Subcopy='" + Evt_Subcopy + "'" + VbCrlf
   sqlStr = sqlStr + " ,Item1='" + Item1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Item2='" + Item2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Item3='" + Item3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,StartDate='" + StartDate + "'" + VbCrlf
   sqlStr = sqlStr + " ,EndDate='" + EndDate + "'" + VbCrlf
   sqlStr = sqlStr + " ,DispOrder='" + DispOrder + "'" + VbCrlf
   sqlStr = sqlStr + " ,Isusing='" + Isusing + "'" + VbCrlf
   sqlStr = sqlStr + " ,LastUser='" + session("ssBctCname") + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + cstr(idx)
   dbget.Execute sqlStr
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->