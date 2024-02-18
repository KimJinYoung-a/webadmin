<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%

dim itemgubunarr, itemidadd, itemoptionarr, mode, title

itemgubunarr = request("itemgubunarr")
itemidadd	= request("itemidadd")
itemoptionarr = request("itemoptionarr")
mode = requestCheckvar(request("mode"),32)

title = "옵션별 랙코드 일괄수정"
if mode <> "modiopt" then
	mode = "modiitem"
	title = "상품 랙코드 일괄수정"
end if

''itemgubunarr = split(itemgubunarr,"|")
''itemidadd	= split(itemidadd,"|")
''itemoptionarr = split(itemoptionarr,"|")

%>

<script language='javascript'>

function jsModiRackCode(frm){
	var confirmMsg = '일괄수정 하시겠습니까?';
    var itemrackcode = frm.itemrackcode.value;

	if ((itemrackcode.length>0) && (itemrackcode.length != 4) && (itemrackcode.length != 8)) {
		alert('랙코드는 4 또는 8 자리로 고정되어있습니다.');
		frm.itemrackcode.focus();
		return;
	}

	var ret = confirm(confirmMsg);
	if(ret){
		frm.submit();
	}
}

window.onload = function() {
	document.frm.itemrackcode.focus();
}

</script>

<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="30">
			<td style="border-bottom:1px solid #BABABA" colspan="2">
				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b><%= title %></b>
			</td>
		</tr>
		<form name="frm" method=post action="popMultiRackCode_process.asp">
		<input type="hidden" name="mode" value="<%= mode %>">
		<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
		<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
		<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
		<tr height="65">
			<td width="60" bgcolor="<%= adminColor("tabletop") %>">랙코드</td>
			<td align="left">
				<input type="text" name="itemrackcode" value="" size="6" maxlength="8" class="text"> (4 or 8자리 Fix)
			</td>
		</tr>
		</form>
		<tr height="30">
			<td align="center" colspan="2" style="border-top:1px solid #BABABA">
				<input type="button" class="button" value=" 수 정 " onclick="jsModiRackCode(frm);">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
