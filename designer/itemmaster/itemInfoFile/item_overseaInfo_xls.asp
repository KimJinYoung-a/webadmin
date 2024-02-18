<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품해외배송정보 일괄변경 Excel 업로드
' Hieditor : 2016.06.03 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'/// 문서형태를 Ms-Excel로 지정 ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=텐바이텐_해외배송_상품_" & Date() & ".xls"
%>
<%
dim eCode, Sdate, Edate, limitLevel
dim strSql

	'// DB에서 목록접수
	strSql = "select i.itemid, i.itemname, isNull(i.deliverOverseas,'N') as deliverOverseas,isNull(i.itemWeight,0) as itemWeight " &_
			"from db_item.dbo.tbl_item as i "&_
			"where i.makerid='" & session("ssBctId") & "' " &_ 
			"	and i.isusing='Y' "

		rsget.Open strSql, dbget, 1
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr style="background-color:#66FFFF;">
	<td colspan="4">※ 내용을 4번째 줄부터 작성해주세요 (노란배경은 필수 입력). 작성 후 다른이름으로 저장 > "Excel 97-2003통합 문서"로 저장 뒤 업로드해주세요.<font color="red"><B>무게는 숫자만</b></font> 입력해주세요(단위,콤마 입력불가능)</td>
</tr>
<tr style="background-color:#FFCCCC; display:none;">
	<td>code</td>
	<td>name</td>
	<td>sYn</td>
	<td>iW</td> 
</tr>
<tr style="background-color:#D8D8D8; color:#5A5A5A;">
	<td align="center">상품코드</td>
	<td align="center">상품명</td>
	<td align="center">해외배송여부</td>
	<td align="center">상품무게(g)</td>
</tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
		do Until rsget.EOF
%>
<tr>
	<td align="center" style="background-color:#FFFF00;"><%=rsget("itemid")%></td>
	<td align="center" style='mso-number-format:"\@"'><%=rsget("itemname")%></td>
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("deliverOverseas")%></td> 
	<td align="center" style='background-color:#FFFF00; mso-number-format:"\@"'><%=rsget("itemWeight")%></td>
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
 rsget.close
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
