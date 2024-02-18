<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 정산
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim topidx, shopid
	topidx = requestCheckVar(request("topidx"),10)
	shopid = requestCheckVar(request("shopid"),32)

dim yyyy,mm, premonth, makerid
premonth = dateadd("m",-1,now())
yyyy = Left(CStr(premonth),4)
mm = Mid(CStr(premonth),6,2)
%>
<script language='javascript'>
function AddValue(frm){
	if (frm.makerid.value.length<1){
		alert('브랜드를 선택하세요.');
		frm.makerid.focus()
		return;
	}
	frm.shopid.disabled=false;
	frm.submit();
}
</script>
<table width="760" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<form name=frm method=post action="meaipchulgojungsan_process.asp">
<input type=hidden name="mode" value="etcsubadd">
<input type=hidden name="topidx" value="<%= topidx %>">
<tr>
	<td bgcolor="#DDDDFF" width=160>매장</td>
	<td bgcolor="#FFFFFF" ><% drawSelectBoxOffShopAll "shopid", shopid %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=160>정산대상월</td>
	<td bgcolor="#FFFFFF" ><% DrawYMBox yyyy,mm %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=160>브랜드</td>
	<td bgcolor="#FFFFFF" ><% drawSelectBoxDesignerwithName "makerid", makerid %></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan=2 align=center>
	<input type=button value="내역추가" onclick="AddValue(frm)">
	</td>
</tr>
</form>
</table>

<script language='javascript'>
document.frm.shopid.disabled=true;
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->