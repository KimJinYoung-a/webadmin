<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg
mode = Request("mode")

'// 상품번호/옵션번호를 받는다 //
Dim itemid, safeCertGbnCd, safeCertOrgCd, safeCertModelNm, safeCertNo, safeCertDt 
itemid				= request("itemid")
safeCertGbnCd		= html2db(request("safeCertGbnCd"))
safeCertOrgCd		= html2db(request("safeCertOrgCd"))
safeCertModelNm		= html2db(request("safeCertModelNm"))
safeCertNo			= html2db(request("safeCertNo"))
safeCertDt			= request("safeCertDt")

'// 모드별 분기
Select Case mode
	Case "I"
		'신규등록
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_gsshop_safecode " & VbCrlf
		sqlStr = sqlStr & " (itemid, safeCertGbnCd, safeCertOrgCd, safeCertModelNm, safeCertNo, safeCertDt)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & itemid & "'"  & VbCrlf
		sqlStr = sqlStr & ", '" & safeCertGbnCd & "','" & safeCertOrgCd & "','" & safeCertModelNm & "', '"& safeCertNo &"','"& safeCertDt &"') "
		dbget.execute sqlStr

	Case "U"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_gsshop_safecode SET " & VbCrlf
		sqlStr = sqlStr & " safeCertGbnCd = '"&safeCertGbnCd&"' " & VbCrlf
		sqlStr = sqlStr & " ,safeCertOrgCd = '"&safeCertOrgCd&"' " & VbCrlf
		sqlStr = sqlStr & " ,safeCertModelNm = '"&safeCertModelNm&"' " & VbCrlf
		sqlStr = sqlStr & " ,safeCertNo = '"&safeCertNo&"' " & VbCrlf
		sqlStr = sqlStr & " ,safeCertDt = '"&safeCertDt&"' " & VbCrlf
		sqlStr = sqlStr & " WHERE itemid='" & itemid & "'" & VbCrlf
		dbget.execute(sqlStr)
End Select

%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.self.location.reload();
parent.window.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->