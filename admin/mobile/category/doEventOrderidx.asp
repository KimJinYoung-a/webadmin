<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  모바일 카테고리 메인 이벤트 순서 수정
' History : 2020.12.02 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// 변수 선언 및 전달값 저장
Dim idxarr, orderidxarr, cnt, sqlStr, ix

	idxarr = request.form("idxarr")
	orderidxarr	= request.form("orderidxarr")
    if idxarr <> "" then
        idxarr = split(idxarr,",")
        cnt = ubound(idxarr)
        orderidxarr = split(orderidxarr,",")
        for ix=0 to cnt	
            sqlStr = "UPDATE [db_sitemaster].[dbo].tbl_display_catemain_ex" & VbCrlf
            sqlStr = sqlStr & " SET view_order = " & Cstr(orderidxarr(ix)) & VbCrlf
            sqlStr = sqlStr &	"	WHERE idx=" & Cstr(idxarr(ix))
            dbget.execute sqlStr
        next
    end if
response.write "<script>parent.location.reload();</script>"
response.write "<script>alert('저장되었습니다.');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->