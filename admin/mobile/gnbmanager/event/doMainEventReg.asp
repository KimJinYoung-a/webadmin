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
Dim idx, Evt_Code, Evt_Title, Evt_Discount, Evt_Coupon, Evt_Subcopy
Dim StartDate, EndDate, DispOrder, Isusing, mode
	
	
	idx = requestCheckVar(request("idx"),10)
	Evt_Code = requestCheckVar(request("Evt_Code"),10)
	Evt_Title = requestCheckVar(request("Evt_Title"),128)
	Evt_Discount = requestCheckVar(request("Evt_Discount"),10)
	Evt_Coupon = requestCheckVar(request("Evt_Coupon"),10)
	Evt_Subcopy = requestCheckVar(request("Evt_Subcopy"),128)
	
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

dim sqlStr, Evt_Img

If Evt_Code <> "" Then
	sqlStr = "select top 1 evt_mo_listbanner from [db_event].[dbo].[tbl_event_display]"
	sqlStr = sqlStr + " where evt_code='" + CStr(Evt_Code) + "'"
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF Or rsget.BOF) Then
		Evt_Img =  rsget("evt_mo_listbanner")
	End If
	rsget.Close
End If

if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_mobile_gnb_main_event]" + VbCrlf
    sqlStr = sqlStr + " (Evt_Code, Evt_Title, Evt_Subcopy, Evt_Discount, Evt_Coupon, Evt_Img, StartDate, EndDate, DispOrder, Isusing, RegUser)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + Evt_Code + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Title + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Subcopy + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Discount + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Coupon + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Img + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + StartDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + EndDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " )"
	'Response.write sqlStr
	'Response.end
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_mobile_gnb_main_event]') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_mobile_gnb_main_event]" + VbCrlf
   sqlStr = sqlStr + " set Evt_Code='" + Evt_Code + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Title='" + Evt_Title + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Discount='" + Evt_Discount + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Coupon='" + Evt_Coupon + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Subcopy='" + Evt_Subcopy + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Img='" + Evt_Img + "'" + VbCrlf
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