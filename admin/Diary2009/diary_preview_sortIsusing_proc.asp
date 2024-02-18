<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  다이어리 어드민 프리뷰 이미지등록 처리 페이지
' History : 2014.08.04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/Diary2009/classes/DiaryCls.asp"-->
<%
Dim cnt, i, sqlStr, sIdx, mode
Dim sIsUsing, sSortNo , idx
	
	mode = Request.form("mode")
	idx = Request.form("idx")
		
if mode="sortisusingedit" then

	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("isusing"&sIdx)

		sqlStr = sqlStr & " UPDATE db_diary2010.dbo.tbl_diary_previewImg SET "  & VBCRLF
		sqlStr = sqlStr & " isusing = '"&sIsUsing&"'" & VBCRLF
		sqlStr = sqlStr & " ,sortnum = '"&sSortNo&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE idx =" & sIdx &";" & VBCRLF
	Next

	if sqlStr <> "" then 
		dbget.execute sqlStr
	end if 
		
	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	location.replace('/admin/diary2009/PreviewImg.asp?idx="&idx&"');"
	response.write "</script>"

elseif mode = "mdpicksortingedit" then 

	for i=1 to request.form("chkIdx").count
		sIdx = request.form("chkIdx")(i)
		sSortNo = request.form("sort"&sIdx)
		sIsUsing = request.form("isusing"&sIdx)

		sqlStr = sqlStr & " UPDATE db_diary2010.dbo.tbl_diaryMaster SET "  & VBCRLF
		sqlStr = sqlStr & " mdpick = '"&sIsUsing&"'" & VBCRLF
		sqlStr = sqlStr & " ,mdpicksort = '"&sSortNo&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE diaryid =" & sIdx &";" & VBCRLF
	Next

	if sqlStr <> "" then 
		dbget.execute sqlStr
	end if 
		
	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	location.replace('/admin/diary2009/diary_mdpicksort.asp');"
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->