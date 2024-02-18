<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/qna_prefacecls.asp"-->
<%
dim code, mode
code = request("code")
mode = request("mode")

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = 1
omd.FPageSize=1
omd.FRectidx = code
omd.FRectmasterid = "01"
omd.GetPrefaceGubunList

%>
<script language="JavaScript">
<!--

function AddIttems(){
	if (frmarr.code.value == ""){
		alert("코드를 넣어주세요!");
		frmarr.code.focus();
	}
	else if (frmarr.cname.value == ""){
		alert("구분명을 입력해주세요!");
		frmarr.cname.focus();
	}
	else if (confirm('추가하시겠습니까?')){
		frmarr.submit();
	}
}

//-->
</script>
<% if mode = "add" then %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="prefacegubun_process.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="masterid" value="01">
<tr>
	<td class="a">코드넣기 : <input type="text" name="code" >&nbsp; 구분명 : <input type="text" name="cname" ></td>
</tr>
<tr>
	<td><input type="button" value="추가" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% else %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="prefacegubun_process.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="code" value="<% = code %>">
<input type="hidden" name="masterid" value="01">
<tr>
	<td class="a">코드넣기 : <input type="text" name="ccode" value="<% = omd.FItemList(0).Fcode %>" readonly>&nbsp; 구분명 : <input type="text" name="cname"  value="<% = omd.FItemList(0).Fcname %>"></td>
</tr>
<tr>
	<td><input type="button" value="수정" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% end if %>
<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->