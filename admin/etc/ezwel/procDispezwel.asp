<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, joongBok, categbn, tDepthCode
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "cate")  Then
	response.write "<script>alert('잘못된 경로입니다');window.close();</script>"
	response.end
End If

Dim catecode, depthCode
catecode	= Request("catecode")
depthCode	= requestCheckvar(Request("depthcode"),10)
joongBok = False
If (mode = "saveCate") Then
	If catecode = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        '중복 확인
		sqlStr = "Delete From db_etcmall.dbo.tbl_ezwel_dispcate_mapping Where catecode= '"& catecode &"' " & VbCrlf
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_ezwel_dispcate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (depthCode, catecode, lastUpdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
		sqlStr = sqlStr & ", '"& catecode &"', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "Delete From db_etcmall.dbo.tbl_ezwel_dispcate_mapping Where catecode= '"& catecode &"' " & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->