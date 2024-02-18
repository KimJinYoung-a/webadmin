<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : addbrandproc.asp
' Description : 브랜드 선택 등록
' History : 2021.02.15 정태훈
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim k, sqlStr, i, brandid
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)
Dim mode : mode = requestCheckVar(Request.Form("mode"),5)
Dim stype : stype = requestCheckVar(Request.Form("stype"),1)
Dim upback : upback = requestCheckVar(Request.Form("upback"),1)
Dim reUrl : reUrl = Request.ServerVariables("HTTP_REFERER")
Dim GroupItemCheck : GroupItemCheck = requestCheckVar(Request.Form("GroupItemCheck"),1)
dim idxarr, idarr, arritems


if Request.Form("idarr") <> "" then
	if checkNotValidHTML(Request.Form("idarr")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if mode="del" Then
	dbget.beginTrans
        For k=1 To request.form("chkIdx").count
            idxarr = request.form("chkIdx")(k)
			sqlStr = " update [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] set isusing='N' WHERE idx in (" & idxarr & ")"
			dbget.execute sqlStr
            IF Err.Number <> 0 THEN
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            END IF
        Next
	dbget.CommitTrans
	Response.write "<script>alert('삭제 되었습니다.');location.replace('popBrandList.asp?idx="&idx&"');</script>"
	response.End
ElseIf mode="idarr" Then
	dbget.beginTrans
        idarr = replace(Trim(Request.Form("idarr"))," ","")
        arritems = split(idarr,",")

		For k=0 To ubound(arritems)

			sqlStr = "IF Not Exists(SELECT IDX FROM [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] WHERE isusing='Y' and brandid='" & arritems(k) & "' and masteridx=" & idx & ")" & vbcrlf
			sqlStr = sqlStr + "BEGIN " & vbcrlf
			sqlStr = sqlStr + "     INSERT INTO [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] (masteridx, brandid)" & vbcrlf
			sqlStr = sqlStr + "     VALUES (" & idx & ", '" & arritems(k) &"')" & vbcrlf
			sqlStr = sqlStr + "END "
			dbget.execute sqlStr
            'Response.write sqlStr
			'Response.end
            IF Err.Number <> 0 THEN
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            END IF
		Next
	dbget.CommitTrans
    Response.write "<script>alert('등록 되었습니다.');opener.fnreload();self.close();</script>"
	response.End 
Else
	dbget.beginTrans
		For k=1 To request.form("chkIdx").count
            brandid = request.form("chkIdx")(k)
			sqlStr = "IF Not Exists(SELECT IDX FROM [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] WHERE isusing='Y' and brandid='" & brandid & "' and masteridx=" & idx & ")" & vbcrlf
			sqlStr = sqlStr + "BEGIN " & vbcrlf
			sqlStr = sqlStr + "     INSERT INTO [db_sitemaster].[dbo].[tbl_brand_link_banner_brand_list] (masteridx, brandid)" & vbcrlf
			sqlStr = sqlStr + "     VALUES (" & idx & ", '" & brandid &"')" & vbcrlf
			sqlStr = sqlStr + "END "
			dbget.execute sqlStr
            'Response.write sqlStr
			'Response.end
            IF Err.Number <> 0 THEN
                dbget.RollBackTrans 
                Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
                response.End 
            END IF
		Next
	dbget.CommitTrans
    Response.write "<script>alert('등록 되었습니다.');opener.fnreload();self.close();</script>"
	response.End 
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->