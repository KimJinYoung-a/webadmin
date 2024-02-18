<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	Dim cCoinMng, vIdx, vCoin, vUseYN, vRegdate
	vIdx = Request("idx")
	
	If vIdx <> "" Then
		set cCoinMng = new ClsMomoCoin
		cCoinMng.FIdx = vIdx
		cCoinMng.FCoinMngView
		
		vCoin = cCoinMng.FOneItem.fcoin
		vUseYN = cCoinMng.FOneItem.fuseyn
		vRegdate = cCoinMng.FOneItem.fregdate
		set cCoinMng = nothing
	End If
%>
<script language="javascript">
function checkform()
{
	if(frm.coin.value == "")
	{
		alert('Coin 값을 입력하세요.');
		frm.coin.focus();
		return false;
	}
	if(isNaN(frm.coin.value))
	{
		alert('Coin 값은 숫자로만 입력하세요.');
		frm.coin.value = "";
		frm.coin.focus();
		return false;
	}
	if (!frm.useyn[0].checked && !frm.useyn[1].checked)
	{
		alert("사용여부를 선택하세요.")
		return false;
	}
}
</script>

<form name="frm" method="post" action="coin_manage_write_proc.asp" onSubmit="return checkform(this);">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% If vIdx <> "" Then %>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="70" bgcolor="<%= adminColor("gray") %>">idx</td>
		<td align="left" width="300"><%=vIdx%></td>
	</tr>
	<% End If %>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="70" bgcolor="<%= adminColor("gray") %>">Coin</td>
		<td align="left" width="300"><input type="text" name="coin" value="<%=vCoin%>" size="10"></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="70" bgcolor="<%= adminColor("gray") %>">사용여부</td>
		<td align="left" width="300">
			<input type="radio" name="useyn" value="y" <% If vUseYN = "y" Then Response.Write "checked" End If %>>Y&nbsp;&nbsp;&nbsp;
			<input type="radio" name="useyn" value="n" <% If vUseYN = "n" Then Response.Write "checked" End If %>>N
		</td>
	</tr>
	<% If vIdx <> "" Then %>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="70" bgcolor="<%= adminColor("gray") %>">등록일</td>
		<td align="left" width="300"><%=vRegdate%></td>
	</tr>
	<% End If %>
</table>
<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td align="right"><input type="submit" value="저장"></td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
