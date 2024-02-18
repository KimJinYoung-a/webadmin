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
Dim idx, MainCopy1, MainCopy2, Evt_Code1, Evt_Title1, Evt_Discount1, Evt_Subcopy1
Dim Evt_Code2, Evt_Title2, Evt_Discount2, Evt_Subcopy2, Evt_Code3, Evt_Title3, Evt_Discount3, Evt_Subcopy3
Dim StartDate, EndDate, DispOrder, Isusing, mode, Evt_Coupon1, Evt_Coupon2, Evt_Coupon3
	
	
	idx = requestCheckVar(request("idx"),10)
	MainCopy1 = requestCheckVar(request("MainCopy1"),128)
	MainCopy2 = requestCheckVar(request("MainCopy2"),128)
	Evt_Code1 = requestCheckVar(request("Evt_Code1"),10)
	Evt_Title1 = requestCheckVar(request("Evt_Title1"),128)
	Evt_Discount1 = requestCheckVar(request("Evt_Discount1"),10)
	Evt_Coupon1 = requestCheckVar(request("Evt_Coupon1"),8)
	Evt_Subcopy1 = requestCheckVar(request("Evt_Subcopy1"),128)

	Evt_Code2 = requestCheckVar(request("Evt_Code2"),10)
	Evt_Title2 = requestCheckVar(request("Evt_Title2"),128)
	Evt_Discount2 = requestCheckVar(request("Evt_Discount2"),10)
	Evt_Coupon2 = requestCheckVar(request("Evt_Coupon2"),8)
	Evt_Subcopy2 = requestCheckVar(request("Evt_Subcopy2"),128)
	
	Evt_Code3 = requestCheckVar(request("Evt_Code3"),10)
	Evt_Title3 = requestCheckVar(request("Evt_Title3"),128)
	Evt_Discount3 = requestCheckVar(request("Evt_Discount3"),10)
	Evt_Coupon3 = requestCheckVar(request("Evt_Coupon3"),8)
	Evt_Subcopy3 = requestCheckVar(request("Evt_Subcopy3"),128)

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

dim sqlStr, Evt_Img1, Evt_Img2, Evt_Img3

If Evt_Code1 <> "" Then
	sqlStr = "select top 1 evt_mo_listbanner from [db_event].[dbo].[tbl_event_display]"
	sqlStr = sqlStr + " where evt_code='" + CStr(Evt_Code1) + "'"
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF Or rsget.BOF) Then
		Evt_Img1 =  rsget("evt_mo_listbanner")
	End If
	rsget.Close
End If
If Evt_Code2 <> "" Then
	sqlStr = "select top 1 evt_mo_listbanner from [db_event].[dbo].[tbl_event_display]"
	sqlStr = sqlStr + " where evt_code='" + CStr(Evt_Code2) + "'"
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF Or rsget.BOF) Then
		Evt_Img2 =  rsget("evt_mo_listbanner")
	End If
	rsget.Close
End If
If Evt_Code3 <> "" Then
	sqlStr = "select top 1 evt_mo_listbanner from [db_event].[dbo].[tbl_event_display]"
	sqlStr = sqlStr + " where evt_code='" + CStr(Evt_Code3) + "'"
	rsget.Open sqlStr,dbget,1
	If Not(rsget.EOF Or rsget.BOF) Then
		Evt_Img3 =  rsget("evt_mo_listbanner")
	End If
	rsget.Close
End If

if (mode = "add") then

    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_main_gather_event]" + VbCrlf
    sqlStr = sqlStr + " (MainCopy1, MainCopy2, Evt_Code1, Evt_Title1, Evt_Subcopy1, Evt_Discount1, Evt_Coupon1, Evt_Img1, Evt_Code2, Evt_Title2, Evt_Subcopy2, Evt_Discount2, Evt_Coupon2, Evt_Img2, Evt_Code3, Evt_Title3, Evt_Subcopy3, Evt_Discount3, Evt_Coupon3, Evt_Img3, StartDate, EndDate, DispOrder, Isusing, RegUser)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + MainCopy1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + MainCopy2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Code1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Title1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Subcopy1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Discount1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Coupon1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Img1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Code2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Title2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Subcopy2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Discount2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Coupon2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Img2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Code3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Title3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Subcopy3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Evt_Discount3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Coupon3 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + Evt_Img3 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + StartDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + EndDate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + DispOrder + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + Isusing + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" +  session("ssBctCname") + "'" + VbCrlf
    sqlStr = sqlStr + " )"
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_main_gather_event]') as idx"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_main_gather_event]" + VbCrlf
   sqlStr = sqlStr + " set MainCopy1='" + MainCopy1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,MainCopy2='" + MainCopy2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Code1='" + Evt_Code1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Title1='" + Evt_Title1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Discount1='" + Evt_Discount1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Coupon1='" + Evt_Coupon1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Subcopy1='" + Evt_Subcopy1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Img1='" + Evt_Img1 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Code2='" + Evt_Code2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Title2='" + Evt_Title2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Discount2='" + Evt_Discount2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Coupon2='" + Evt_Coupon2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Subcopy2='" + Evt_Subcopy2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Img2='" + Evt_Img2 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Code3='" + Evt_Code3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Title3='" + Evt_Title3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Discount3='" + Evt_Discount3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Coupon3='" + Evt_Coupon3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Subcopy3='" + Evt_Subcopy3 + "'" + VbCrlf
   sqlStr = sqlStr + " ,Evt_Img3='" + Evt_Img3 + "'" + VbCrlf
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