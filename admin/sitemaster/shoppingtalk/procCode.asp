<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : GIFT TALK 코드 관리
' Hieditor : 강준구 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/shoppingtalk/classes/shoppingtalkCls.asp" -->

<%
	Dim vQuery, i, vDepth, vCode, vKeyword, vCodeName, vSortNo, vUseYN
	vDepth = Request("depth")
	vCode = Request("NewCode")
	vKeyword = Request("NewKeyword1")
	vCodeName = Request("NewCodename")
	vSortNo = Request("NewSort")
	vUseYN = Request("NewUseyn")
	
if vCodeName <> "" and not(isnull(vCodeName)) then
	vCodeName = ReplaceBracket(vCodeName)
end If

	vQuery = vQuery & "IF EXISTS(select code from [db_board].[dbo].[tbl_shopping_talk_keywordcode] where code = '" & vCode & "')" & vbCrLf
	vQuery = vQuery & "BEGIN " & vbCrLf
	vQuery = vQuery & "		UPDATE [db_board].[dbo].[tbl_shopping_talk_keywordcode] SET " & vbCrLf
	vQuery = vQuery & "			codename = '" & vCodeName & "', " & vbCrLf
	vQuery = vQuery & "			sortno = '" & vSortNo & "', " & vbCrLf
	vQuery = vQuery & "			useyn = '" & vUseYN & "' " & vbCrLf
	vQuery = vQuery & "		WHERE code = '" & vCode & "' " & vbCrLf
	vQuery = vQuery & "END " & vbCrLf
	vQuery = vQuery & "ELSE " & vbCrLf
	vQuery = vQuery & "BEGIN " & vbCrLf
	vQuery = vQuery & "		INSERT INTO [db_board].[dbo].[tbl_shopping_talk_keywordcode](depth, code, codename, sortno, useyn) " & vbCrLf
	vQuery = vQuery & "		VALUES('" & vDepth & "', '" & vCode & "', '" & vCodeName & "', '" & vSortNo & "', '" & vUseYN & "') " & vbCrLf
	vQuery = vQuery & "END " & vbCrLf
	dbget.execute vQuery
%>

<script type='text/javascript'>
location.href = "PopManageCode.asp?keyword1=<%=vKeyword%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->