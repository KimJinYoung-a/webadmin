<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : 관리카테고리 프로세스
' History : 2017.03.10 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim fromDate, toDate, mode, sqlstr, arrlist, bufStr, i, tendb
	fromDate = requestcheckvar(request("fromDate"),10)
	toDate = requestcheckvar(request("toDate"),10)
	mode = requestcheckvar(request("mode"),32)

IF application("Svr_Info")="Dev" THEN
	tendb = "tendb."
end IF

'전시카테고리 전체 리스트
if mode="categorylist" then
	sqlstr = "select top 10000" & vbcrlf
	sqlstr = sqlstr & " l.code_large+m.code_mid+s.code_small as catecd, l.code_large AS cdlarge, m.code_mid AS cdmid, s.code_small AS cdsmall" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(l.code_nm,char(9),''),char(10),''),char(13),'') as nmlarge" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm,char(9),''),char(10),''),char(13),'') as nmmid" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm,char(9),''),char(10),''),char(13),'') as nmsmall" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_eng,char(9),''),char(10),''),char(13),'') as nmmid_eng" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_eng,char(9),''),char(10),''),char(13),'') as nmsmall_eng" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_cn_gan,char(9),''),char(10),''),char(13),'') as mid_nm_cn_gan" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_cn_bun,char(9),''),char(10),''),char(13),'') as mid_nm_cn_bun" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_cn_gan,char(9),''),char(10),''),char(13),'') as small_nm_cn_gan" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_cn_bun,char(9),''),char(10),''),char(13),'') as small_nm_cn_bun" & vbcrlf
	sqlstr = sqlstr & " FROM "& tendb &"db_item.[dbo].tbl_Cate_large l" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_mid m" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = m.code_large" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_small s" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = s.code_large" & vbcrlf
	sqlstr = sqlstr & " 	AND m.code_mid = s.code_mid" & vbcrlf
	sqlstr = sqlstr & " order by cdlarge asc, cdsmall asc, nmlarge asc" & vbcrlf

	'response.write sqlstr & "<br>"
	db3_rsget.open sqlstr,db3_dbget,1
	If Not db3_rsget.Eof Then
		arrlist = db3_rsget.getrows()
	End If
	db3_rsget.close

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=categorydownload.xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<td class='txt'>

	<table>
	<tr>
	<td>카테고리코드</td>
	<td>대카테코드</td>
	<td>중카테코드</td>
	<td>소카테코드</td>
	<td>대카테명</td>
	<td>중카테명</td>
	<td>소카테명</td>
	<td>중카테영문명</td>
	<td>소카테영문명</td>
	<td>중카테간자체</td>
	<td>중카테번자체</td>
	<td>소카테간자체</td>
	<td>소카테번자체</td>
	</tr>

	<%
	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
	%>	
		<tr>
			<td class='txt'><%= arrlist(0,i) %></td>
			<td class='txt'><%= arrlist(1,i) %></td>
			<td class='txt'><%= arrlist(2,i) %></td>
			<td class='txt'><%= arrlist(3,i) %></td>
			<td><%= arrlist(4,i) %></td>
			<td><%= arrlist(5,i) %></td>
			<td><%= arrlist(6,i) %></td>
			<td><%= arrlist(7,i) %></td>
			<td><%= arrlist(8,i) %></td>
			<td><%= arrlist(9,i) %></td>
			<td><%= arrlist(10,i) %></td>
			<td><%= arrlist(11,i) %></td>
			<td><%= arrlist(12,i) %></td>
		</tr>
	<%
		next
	end if
	%>
	</table>
</html>
<%
else
	response.write "잘못된 경로 입니다."
	session.codePage = 949
	dbget.close() : response.end
end if
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->