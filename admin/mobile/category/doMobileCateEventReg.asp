<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  모바일 카테고리 메인 이벤트 작성/수정
' History : 2020.12.02 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// 변수 선언 및 전달값 저장
Dim mode, mainimg, moreimg
Dim view_yn, idx, catecode
Dim linkurl, sqlStr
Dim adminid, orderidx
Dim startdate, enddate
Dim itemid1, itemid2, evt_code
Dim makerid, maincopy, subcopy, menupos

	mode = request.form("mode")
	idx	= request.form("idx")
	startdate = request.form("start_date")& " 00:00:00"
	enddate	= request.form("end_date")& " 23:59:59"
	view_yn	= request.form("view_yn")
	orderidx = request.form("orderidx")

	evt_code = request.form("evt_code")
    catecode = request.form("catecode")

    if idx="0" or idx="" then
        mode="add"
    else
        mode="edit"
    end if

if mode="" then
    Call Alert_return("not valid code.")
    dbget.Close: Response.End   
end if

'/신규등록
if mode="add" then
    sqlStr = " insert into db_sitemaster.[dbo].[tbl_display_catemain_ex] " & VbCrlf
    sqlStr = sqlStr & " (catecode, evt_code, reguserid, regdate, view_order, start_date, end_date, view_yn) " & VbCrlf
    sqlStr = sqlStr & " values(" & VbCrlf
    sqlStr = sqlStr & " '" & catecode & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & evt_code & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & session("ssBctId") & "'" & VbCrlf
    sqlStr = sqlStr & " , getdate()" & VbCrlf
    sqlStr = sqlStr & " ,'" & orderidx & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & enddate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & view_yn & "'" & VbCrlf
    sqlStr = sqlStr & " )"
    dbget.Execute sqlStr
else
   sqlStr = "update db_sitemaster.[dbo].[tbl_display_catemain_ex] " & VbCrlf
   sqlStr = sqlStr & " set catecode='" & catecode & "'" & VbCrlf
   sqlStr = sqlStr & " , evt_code ='" & evt_code & "'" & VbCrlf
   sqlStr = sqlStr & " , view_order='" & orderidx & "'" & VbCrlf
   sqlStr = sqlStr & " , start_date='" & startdate & "'" & VbCrlf
   sqlStr = sqlStr & " , end_date='" & enddate & "'" & VbCrlf
   sqlStr = sqlStr & " , view_yn='" & view_yn & "'" & VbCrlf
   sqlStr = sqlStr & " where idx='" & Cstr(idx) & "'"
   dbget.Execute sqlStr
end if

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>opener.window.location.reload();</script>"
response.write "<script>alert('저장되었습니다.');self.close();</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->