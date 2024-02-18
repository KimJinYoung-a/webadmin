<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : cate_insert.asp
' Discription : 모바일 catesub
' History : 2015-09-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
Dim idx , isusing , mode
Dim lp , ii
Dim gcode , gname , dcode , dname
	idx = requestCheckvar(request("idx"),16)

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

If idx <> "" then
	dim subcodeLIst
	set subcodeLIst = new GNBsubcode
	subcodeLIst.Fidx = idx
	subcodeLIst.GetOneSubCode()

	gcode				=	subcodeLIst.FOneItem.Fgnbcode
	dcode				=	subcodeLIst.FOneItem.Fdispcode
	isusing				=	subcodeLIst.FOneItem.Fisusing

	set subcodeLIst = Nothing
End If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function jsSubmit(){
	var frm = document.frm;

	if (frm.gcode.value == "")
	{
		alert("GNB를 선택 해주세요");
		frm.gcode.focus();
		return;
	}

	if (frm.dcode.value == "")
	{
		alert("전시 카테고리를 선택 해주세요");
		frm.dcode.focus();
		return;
	}

	if (confirm('저장 하시겠습니까?')){
		//frm.target = "blank";
		frm.submit();
	}
}
function jsgolist(){
	 self.close(); 
}
</script>
<table width="650" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="docateproc.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<tr bgcolor="#FFFFFF">
	<% If idx = ""  Then %>
	<td colspan="2" align="center" height="35">등록중 입니다.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">수정중 입니다.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td><%=idx%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="20%">GNB 선택</td>
    <td>
		<% Call drawSelectBoxGNB("gcode" , gcode) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="20%">전시 카테고리 선택</td>
    <td>
		<% Call drawSelectBoxDISP("dcode" , dcode) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
