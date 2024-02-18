<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->

<%
Dim i, strSql, mode, keywordidx, keywordcodegubun, keywordname, sortno, isusing, keywordtype
	mode = requestCheckVar(request("mode"),32)
	keywordidx = requestCheckVar(getNumeric(request("keywordidx")),10)
	keywordcodegubun = Request("keywordcodegubun")
	keywordname = Request("keywordname")	
	sortno = requestCheckVar(getNumeric(request("sortno")),10)
	isusing = requestCheckVar(request("isusing"),1)
	keywordtype = Request("keywordtype")

if mode="keywordedit" then
	If keywordname = "" or sortno = "" or isusing = "" Then
		Response.Write "<script type='text/javascript'>alert('입력값이 없습니다.'); history.go(-1);</script>"
		dbget.close() : Response.End
	End IF

	if keywordname <> "" and not(isnull(keywordname)) then
		keywordname = ReplaceBracket(keywordname)
	end If

	strSql = "IF EXISTS(select keywordidx from db_board.dbo.tbl_gift_keyword where keywordtype = '" & trim(keywordtype) & "' and keywordidx = '" & trim(keywordidx) & "')" & vbCrLf
	strSql = strSql & "BEGIN " & vbCrLf
	strSql = strSql & "		UPDATE db_board.dbo.tbl_gift_keyword SET " & vbCrLf
	strSql = strSql & "		keywordname = '" & html2db(trim(keywordname)) & "'" & vbCrLf
	strSql = strSql & "		,sortno = '" & trim(sortno) & "'" & vbCrLf
	strSql = strSql & "		,isusing = '" & trim(isusing) & "'" & vbCrLf
	strSql = strSql & "		where keywordtype = '" & trim(keywordtype) & "' and keywordidx = '" & trim(keywordidx) & "'" & vbCrLf
	strSql = strSql & "END " & vbCrLf
	strSql = strSql & "ELSE " & vbCrLf
	strSql = strSql & "BEGIN " & vbCrLf
	strSql = strSql & "		INSERT INTO db_board.dbo.tbl_gift_keyword(keywordtype, keywordname, sortno, isusing, regdate)" & vbCrLf
	strSql = strSql & "		VALUES('" & trim(keywordtype) & "', '" & html2db(trim(keywordname)) & "'" & vbCrLf
	strSql = strSql & "		, '" & trim(sortno) & "', '" & trim(isusing) & "',getdate()) " & vbCrLf
	strSql = strSql & "END"
	
	'response.write strSql & "<br>"
	dbget.execute strSql

else
	Response.Write "<script type='text/javascript'>alert('정상적인 경로가 아닙니다.'); history.go(-1);</script>"
	dbget.close() : Response.End
end if
%>

<script type='text/javascript'>
	alert('저장되었습니다.');
	location.href = "/admin/sitemaster/gift/day/keyword/giftday_keyword.asp?keywordcodegubun=<%=keywordcodegubun%>&menupos=<%= menupos %>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->