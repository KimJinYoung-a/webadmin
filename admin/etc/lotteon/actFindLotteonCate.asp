<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotteon/lotteonCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim oLotteon, i, page, srcKwd
page		= requestCheckVar(request("page"),10)
srcKwd		= Trim(requestCheckVar(request("srcKwd"),50))

If page = ""	Then page = 1
'// 목록 접수
Set oLotteon = new CLotteon
	oLotteon.FPageSize = 1000
	oLotteon.FCurrPage = page
	oLotteon.FRectSearchName = srcKwd
	oLotteon.getLotteonCateList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function chkThis(v){
	$(".dps").hide();
	$("#dp_"+v).show();
	$('input[name="disp_cat_id"]').removeAttr('checked');
	$.ajax({
		url: "actFindLotteonDispCate.asp?std_cat_id="+v,
		cache: false,
		async: false,
		success: function(message) {
			$("#disp_"+v).empty().html(message);
		},
		error: function(){
			alert(message);
		}
	});
}
</script>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="right"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oLotteon.FtotalCount%></strong>&nbsp;&nbsp;</td>
</tr>
</table>
<form name="resultFrm" target="xLink">
<input type="hidden" name="cdl" value="">
<input type="hidden" name="cdm" value="">
<input type="hidden" name="cds" value="">
<input type="hidden" name="mode" value="saveCate">

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td></td>
	<td>표준카테고리</td>
	<td>표준카테고리명</td>
</tr>
<% If oLotteon.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oLotteon.FresultCount - 1
%>
<tr align="center" height="25"  title="카테고리 선택" bgcolor="#FFFFFF">
    <td>
		<input type="radio" class="radio" name="std_cat_id" value="<%= oLotteon.FItemList(i).FStd_cat_id %>" onclick="chkThis(this.value);" />
	</td>
	<td><%= oLotteon.FItemList(i).FStd_cat_id %></td>
	<td><%= oLotteon.FItemList(i).FStd_cat_nm %></td>
</tr>
<tr class="dps" id="dp_<%= oLotteon.FItemList(i).FStd_cat_id %>" style="display:none;">
	<td bgcolor="#FFFFFF"></td>
	<td colspan="2" bgcolor="#F2F2F2">
		<div id="disp_<%= oLotteon.FItemList(i).FStd_cat_id %>" ></div>
	</td>
</tr>
<%
		Next
	End If
%>
</table>
</form>
<% Set oLotteon = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
