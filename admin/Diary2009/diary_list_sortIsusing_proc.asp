<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  다이어리 리스트 어드민 전시번호,사용여부 처리 페이지
' History : 2015.09.14 유태욱 생성
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
Dim msg
Dim tmpSort, tmpIsusing
Dim cnt, i, sqlStr, idx, mode
Dim detailidxarr, isusingarr, sortnoarr
''	idx	= Request("idx")
	mode = Request("mode")
	sortnoarr 	= Request("sortnoarr")
	isusingarr = Request("isusingarr")
	detailidxarr = Request("detailidxarr")

dbget.beginTrans
if mode="sortisusingedit" then

	'선택이미지 파악
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	sortnoarr	=  split(sortnoarr,",")
	isusingarr	=  split(isusingarr,",")


	For i = 0 to cnt
		tmpSort = sortnoarr(i)
		tmpIsusing = isusingarr(i)
		
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_diary2010.dbo.tbl_DiaryMaster SET " & VBCRLF
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,mdpicksort = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE diaryID =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	msg = "저장 되었습니다"

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diary2009/"
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
'	response.write "<script language='javascript'>"
'	response.write "	alert('저장되었습니다');"
'	response.write "	location.replace('/admin/diary2009/index.asp);"
'	response.write "</script>"
else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->