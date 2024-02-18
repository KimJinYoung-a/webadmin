<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 안정인증정보 일괄변경 Excel 다운로드 + 상품목록
' Hieditor : 2015.05.26 허진원 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'/// 문서형태를 Ms-Excel로 지정 ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=텐바이텐_상품안전인증대상_상품_" & Date() & ".xls"
%>
<%
dim eCode, Sdate, Edate, limitLevel
dim strSql

	'// DB에서 목록접수
	strSql = "select i.itemid, i.itemname, isNull(c.safetyYn,'N') as safetyYn, isnull(c.safetyDiv,0) as safetyDiv, c.safetyNum " &_
			"from db_item.dbo.tbl_item as i " &_
			"	join db_item.dbo.tbl_item_Contents as c " &_
			"		on i.itemid=c.itemid " &_
			"where i.makerid='" & session("ssBctId") & "' " &_
			"	and (isnull(c.safetyYn,'N')='N' " &_
			"		or (c.safetyYn='Y' and safetyDiv<>'10') " &_
			"	) " &_
			"	and i.isusing='Y' "

		rsget.Open strSql, dbget, 1
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr style="background-color:#66FFFF;">
	<td colspan="5">※ 내용을 4번째 줄부터 작성해주세요 (노란배경은 필수 입력). 작성 후 다른이름으로 저장 > "Excel 97-2003통합 문서"로 저장 뒤 업로드해주세요.</td>
</tr>
<tr style="background-color:#FFCCCC; display:none;">
	<td>code</td>
	<td>name</td>
	<td>sYn</td>
	<td>Div</td>
	<td>sNum</td>
</tr>
<tr style="background-color:#D8D8D8; color:#5A5A5A;">
	<td align="center">상품코드</td>
	<td align="center">상품명</td>
	<td align="center" colspan="2">안전인증 대상 여부</td>
	<td align="center">국가통합인증:KC마크 번호</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td align="center" style="background-color:#FFFF00;"><%=rsget("itemid")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=rsget("itemname")%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("safetyYn")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=getSaftyDivName(rsget("safetyDiv"))%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("safetyNum")%></td>
</tr>
<%
		rsget.MoveNext
		loop
	else
%>
<%	end if %>
</table>
</body>
</html>
<%
 function getSaftyDivName(sdiv)
 	Select Case cStr(sdiv)
 		Case "10"
 			getSaftyDivName = "10:국가통합인증(KC마크)"
 		Case "20"
 			getSaftyDivName = "20:전기용품 안전인증"
 		Case "30"
 			getSaftyDivName = "30:KPS 안전인증 표시"
 		Case "40"
 			getSaftyDivName = "40:KPS 자율안전 확인 표시"
 		Case "50"
 			getSaftyDivName = "50:KPS 어린이 보호포장 표시"
 		Case Else
 			getSaftyDivName = ""
 	End Select
 end Function
 rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
