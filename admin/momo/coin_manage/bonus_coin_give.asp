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
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<%
	Dim giveuserid, givecontent, givecoin
	Dim vIdx
	vIdx = Request("idx")
	If vIdx <> "" Then
		rsget.open "select userid, gubun, coin from [db_momo].[dbo].[tbl_coin_log] where id = '" & vIdx & "'",dbget,1
		giveuserid = rsget(0)
		givecontent = rsget(1)
		givecoin = rsget(2)
		rsget.close()
	End If
%>

<script language="javascript">
function goGive()
{
	<% If vIdx = "" Then %>
		if(!(frm1.many.checked))
		{
			if(frm1.giveuserid.value == "")
			{
				alert("아이디를 입력하세요");
				frm1.giveuserid.focus();
				return false;
			}
		}
		else
		{
			if(frm1.giveuserid_m.value == "")
			{
				alert("아이디를 입력하세요");
				frm1.giveuserid_m.focus();
				return false;
			}
		}
	<% End IF %>
	if(frm1.givecontent.value == "")
	{
		alert("보너스 내역을 입력하세요");
		frm1.givecontent.focus();
		return false;
	}
	if(frm1.givecoin.value == "")
	{
		alert("지급코인을 입력하세요");
		frm1.givecoin.focus();
		return false;
	}
	if(isNaN(frm1.givecoin.value))
	{
		alert("지급코인을 숫자로만 입력하세요");
		frm1.givecoin.value = "";
		frm1.givecoin.focus();
		return false;
	}
	
	<% If vIdx <> "" Then %>
	if(confirm("" + frm1.giveuserid.value + " 님의 보너스 내역을 수정하시겠습니까?") == true)
	<% Else %>
		var tmp;
		if(!(frm1.many.checked))
		{
			tmp = "" + frm1.giveuserid.value + " 님께 "
		}
		else
		{
			tmp = ""
		}
			if(confirm("" + tmp + "" + frm1.givecoin.value + " 코인을 지급하시겠습니까?") == true)

	<% End IF %>
	{
		frm1.submit();
	} else {
		return false;
	}
}

function manyopen()
{
	if(!(frm1.many.checked))
	{
		frm1.insertgubun.value = "one";
		oneinsert.style.display = "block";
		manyinsert.style.display = "none";
	}
	else
	{
		frm1.insertgubun.value = "many";
		oneinsert.style.display = "none";
		manyinsert.style.display = "block";
	}
}
</script>

<% If vIdx = "" Then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" border="1" class="a">
<form name="frm1" method="post" action="bonus_coin_giveproc.asp">
<input type="hidden" name="insertgubun" value="one">
<tr height="60" align="center">
	<td width="120" style="padding:5 0 0 0">보너스코인지급<br><label id="many" style="cursor:pointer" onClick="manyopen()"><input type="checkbox" name="many" id="many" value="o" onClick="manyopen()">10명이상지급</label></td>
	<td align="left">
	    <div id="oneinsert">&nbsp;아이디: <input type="text" name="giveuserid" value="<%=giveuserid%>" size="10"></div>
	    &nbsp;보너스 내역: <input type="text" name="givecontent" value="<%=givecontent%>" size="40">
	    &nbsp;지급코인: <input type="text" name="givecoin" value="<%=givecoin%>" size="5">
	    <div id="manyinsert" style="display:none">&nbsp;아이디(쉼표로나눔)<br>&nbsp;<textarea name="giveuserid_m" rows="5" cols="100"></textarea></div>
	</td>
	<td width="50">
		<input type="button" class="button_s" value="지급" onClick="javascript:goGive();">
	</td>
</tr>
</form>
</table>
<% Else %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" border="1" class="a">
<form name="frm1" method="post" action="bonus_coin_giveproc.asp">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="insertgubun" value="one">
<tr height="40" align="center">
	<td width="120" style="padding:5 0 0 0"><b><font color="red">보내스내역만<br>수정가능!</font></b></td>
	<td align="left">
	    &nbsp;아이디: <input type="text" name="giveuserid" value="<%=giveuserid%>" size="10" ondragstart="return false" onselectstart="return false" readonly>
	    &nbsp;보너스 내역: <input type="text" name="givecontent" value="<%=givecontent%>" size="40">
	    &nbsp;지급코인: <input type="text" name="givecoin" value="<%=givecoin%>" size="5" ondragstart="return false" onselectstart="return false" readonly>
	</td>
	<td width="50">
		<input type="button" class="button_s" value="지급" onClick="javascript:goGive();">
	</td>
</tr>
</form>
</table>
<% End If %>
<br>※ 마이너스 코인 지급시(불량게시글자에게 페널티 등) 지급코인에 - 를 붙이면 됩니다. (예: -10)

<!-- #include virtual="/lib/db/dbclose.asp" -->