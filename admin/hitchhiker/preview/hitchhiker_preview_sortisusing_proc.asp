<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  히치하이커 어드민 프리뷰 Iframe이미지등록 처리 페이지
' History : 2014.08.04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_previewCls.asp"-->
<%
Dim tmpSort, tmpIsusing
Dim cnt, i, sqlStr, idx, mode
Dim detailidxarr, isusingarr, sortnoarr, device
	idx	= Request("idx")
	mode = Request("mode")
	sortnoarr 	= Request("sortnoarr")
	isusingarr = Request("isusingarr")
	detailidxarr = Request("detailidxarr")
	device		= Request("device")

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
		sqlStr = sqlStr & " UPDATE db_sitemaster.dbo.tbl_hitchhiker_preview_Detail SET " & VBCRLF
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,sortnum = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE detailidx =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	if device = "W" then
		response.write "	location.replace('/admin/hitchhiker/preview/iframe_hitchhiker_preview.asp?idx="&idx&"');"
	else
		response.write "	location.replace('/admin/hitchhiker/preview/iframe_hitchhiker_preview_M.asp?idx="&idx&"');"
	end if
	response.write "</script>"
else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->