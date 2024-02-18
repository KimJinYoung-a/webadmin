<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim midx
midx = request("midx")
If NOT isNumeric(midx) Then
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');window.close();</script>"
	dbget.close()	:	response.End
End If
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function delCateItem(v)
{
	$("#delIdx").val(v);
	document.frm.target = "xLink";
	document.frm.submit();
}

function popCateSelect(){
	$.ajax({
		url: "/admin/etc/ssg/act_CategorySelect.asp",

		cache: false,
		success: function(message) {
			$("#lyrCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}
</script>

<form name="frm" method="post" action="procSsgMargin.asp" onsubmit="return false;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" id="mode" name="mode" value="cateDetail">
<input type="hidden" id="delIdx" name="delIdx" value="">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td><%= getCategory(midx) %></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popCateSelect();"></td>
		</tr>
		</table>
		<div id="lyrCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
