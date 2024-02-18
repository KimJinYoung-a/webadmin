<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 등록 제외 단어
' Hieditor : 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim arrRows, oOutMall, i
Dim strSql, mode, keyword, ukeyword, idx
mode		= request("mode")
keyword		= request("keyword")
ukeyword	= request("ukeyword")
idx			= request("idx")

If mode <> "" Then
	strSql = ""
	If mode = "I" Then
		strSql = strSql & "	If NOT EXISTS(SELECT * FROM db_etcmall.dbo.tbl_outmall_not_in_keywords WHERE keyword = '" & html2db(keyword) & "') "
		strSql = strSql & "	BEGIN "
		strSql = strSql & "		INSERT INTO db_etcmall.dbo.tbl_outmall_not_in_keywords (keyword, regdate) values ('" & html2db(keyword) & "', getdate()) "
		strSql = strSql & "	END "
	ElseIf mode = "U" Then
		strSql = ""
		strSql = strSql & "	UPDATE db_etcmall.dbo.tbl_outmall_not_in_keywords "
		strSql = strSql & "	SET keyword = '"& html2db(ukeyword) &"' "
		strSql = strSql & "	WHERE idx = '"& idx &"' "
	ElseIf mode = "D" Then
		strSql = strSql & "	DELETE FROM db_etcmall.dbo.tbl_outmall_not_in_keywords WHERE idx = '"& idx &"';"
		strSql = strSql & "	DELETE FROM db_etcmall.[dbo].[tbl_outmall_not_in_keywords_mallid] WHERE midx = '"& idx &"';"
	End If
	dbget.execute strSql
	Response.Write "<script>parent.location.reload();</script>"
	Response.End
End If

SET oOutMall = new cOutmall
	arrRows = oOutMall.fnNotInKeywordList
SET oOutMall = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function fnkeywordProc(mode, val){
	if(mode == 'U'){
		$("#idx").val(val);
		$("#ukeyword").val($("#keyword_"+val+"").val());
	}else if(mode == 'D'){
		$("#idx").val(val);
	}
	$("#mode").val(mode);
	document.frm.target = "xLink";
	document.frm.submit();
}
function ajaxOutmall(v) {
	$.ajax({
		url: "actOutmalllist.asp?idx="+v,
		cache: false,
		async: false,
		success: function(message) {
			$("#btn_"+v).hide();
			$("#outmalllist_"+v).empty().html(message);
		},
		error: function(){
			alert(message);
		}
	});
}
function ajaxOutmall222(v) {
	$("#btn_"+v).show();
	$("#outmalllist_"+v).empty().html();
}
</script>
<form name="frm" method="post" action="notinkeyword.asp" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" id="mode" name="mode">
<input type="hidden" id="idx" name="idx" value="">
<input type="hidden" id="ukeyword" name="ukeyword" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="50" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>제외단어</td>
	<td bgcolor="FFFFFF" align="left">
		<input type="text" name="keyword" value="" size="50" class="text">
		<input type="button" class="button" value="저장" onclick="fnkeywordProc('I', '');">
	</td>
</tr>
</table>
</form>
<br /><br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>idx</td>
	<td>등록 제외 단어</td>
	<td>제휴몰</td>
	<td width="200">등록일</td>
</tr>
<% If isArray(arrRows) Then %>
<% For i = 0 To Ubound(arrRows, 2) %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= arrRows(0, i) %></td>
	<td>
		<input type="text" id="keyword_<%= arrRows(0, i) %>" name="keyword" value="<%= arrRows(1, i) %>" size="50" class="text">
		<input type="button" class="button" value="수정" onclick="fnkeywordProc('U', '<%= arrRows(0, i) %>');" style=color:blue;font-weight:bold>
		<input type="button" class="button" value="삭제" onclick="fnkeywordProc('D', '<%= arrRows(0, i) %>');" style=color:red;font-weight:bold>
	</td>
	<td>
		<div id="outmalllist_<%= arrRows(0, i) %>"></div>
		<input type="button" class="button" value="보기" id="btn_<%= arrRows(0, i) %>" onclick="ajaxOutmall('<%= arrRows(0, i) %>');">
	</td>
	<td><%= arrRows(2, i) %></td>
</tr>
<% Next %>
<% Else %>
<tr height="50" bgcolor="FFFFFF">
	<td colspan="4" align="center">
		데이터가 없습니다
	</td>
</tr>
<% End If %>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100%"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->