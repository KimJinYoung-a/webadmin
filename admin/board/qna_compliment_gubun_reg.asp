<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/qna_complimentcls.asp"-->
<%
dim idx, gubun,mode
idx = request("idx")
gubun = request("gubun")
mode = request("mode")

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = 1
omd.FPageSize=1
omd.FRectidx = idx
omd.GetQnaComplimentGubun

%>
<script language="JavaScript">
<!--

function AddIttems(){
	if (frmarr.code.value == ""){
		alert("�ڵ带 �־��ּ���!");
		frmarr.code.focus();
	}
	else if (frmarr.cname.value == ""){
		alert("���и��� �Է����ּ���!");
		frmarr.cname.focus();
	}
	else if (confirm('�߰��Ͻðڽ��ϱ�?')){
		frmarr.submit();
	}
}

//-->
</script>
<% if mode = "add" then %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="docomplimentgubun.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="masterid" value="02">
<tr>
	<td class="a">�ڵ�ֱ� : <input type="text" name="code" >&nbsp; ���и� : <input type="text" name="cname" ></td>
</tr>
<tr>
	<td><input type="button" value="�߰�" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% else %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="docomplimentgubun.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="code" value="<% = code %>">
<input type="hidden" name="masterid" value="02">
<tr>
	<td class="a">�ڵ�ֱ� : <input type="text" name="ccode" value="<% = omd.FItemList(0).Fcode %>" readonly>&nbsp; ���и� : <input type="text" name="cname"  value="<% = omd.FItemList(0).Fcname %>"></td>
</tr>
<tr>
	<td><input type="button" value="����" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% end if %>
<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->