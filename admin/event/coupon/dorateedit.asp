<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
dim mode, sqlStr
dim idx, rate, evt_code, coupon

mode = request("mode")
idx = request("idx")
evt_code = request("evt_code")
coupon = request("coupon")
rate = request("rate")

'// 모드에 따른 분기
Select Case mode
	Case "edit"
		sqlStr = "Update [db_event].[dbo].[tbl_event_random_coupon] " &_
				" 	Set rate ='" & rate & "'" &_
				" 	, coupon ='" & coupon & "'" &_
				" Where idx =" & idx
		dbget.Execute(sqlStr)
	Case "add"
		sqlStr = "Insert Into [db_event].[dbo].[tbl_event_random_coupon] " &_
					" ( evt_code, rate, coupon" &_
					" ) values "&_
					" (" & evt_code &_
					" ," & rate &_
					" ," & coupon &_
					")"		
		dbget.Execute(sqlStr)
	Case "delete"
		sqlStr = "delete from [db_event].[dbo].[tbl_event_random_coupon] " &_
				" Where idx =" & idx
		dbget.Execute(sqlStr)
End Select
%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
    location.replace('/admin/event/coupon/random_coupon.asp?evt_code=<%=evt_code%>')
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
