<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, joongBok, categbn, siteNo
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "cate")  Then
	response.write "<script>alert('잘못된 경로입니다');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, DispCtgId
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
DispCtgId= requestCheckvar(Request("DispCtgId"),20)
siteNo	= requestCheckvar(Request("siteNo"),4)
joongBok = False
If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        ' '중복 확인
        ' If categbn = "cate" Then
	    '     sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] "  & VbCrlf
		' 	sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
		' 	rsget.Open sqlStr,dbget,1
		' 	If rsget("cnt") > 0 Then
		' 	     joongBok = True
		' 	End If
		' 	rsget.Close
		' End If

		' If joongBok = False Then
		' 	sqlStr = ""
		' 	sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_DispCate_mapping]  " & VbCrlf
		' 	sqlStr = sqlStr & " (dispCtgId, tenCateLarge, tenCateMid, tenCateSmall, siteNo, lastUpdate)" & VbCrlf
		' 	sqlStr = sqlStr & " VALUES('" & DispCtgId & "' "  & VbCrlf
		' 	sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', '"&siteNo&"', getdate()) "
		' 	dbget.execute sqlStr
		' Else
		'     iErrMsg = "이미 매핑된 카테고리는  추가할 수 없습니다."
		' End If
		sqlStr = "DELETE FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_DispCate_mapping]  " & VbCrlf
		sqlStr = sqlStr & " (dispCtgId, tenCateLarge, tenCateMid, tenCateSmall, siteNo, lastUpdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & DispCtgId & "' "  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', '"&siteNo&"', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and dispCtgId='" & DispCtgId & "'"
		sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
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
<!-- #include virtual="/lib/db/dbclose.asp" -->