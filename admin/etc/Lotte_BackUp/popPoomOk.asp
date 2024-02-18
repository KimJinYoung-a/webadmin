<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim mallid, infoDiv, itemid, mat_Name, mat_percent, mat_place
mallid		= request("mallid")
infoDiv		= request("infoDiv")
itemid		= request("itemid")

Dim strSql, cnt, mode
Dim imat_Name, imat_percent, imat_place, imaterial
mode = request("mode")

If mode = "I" Then
	imat_Name		= request("mat_Name")
	imat_percent	= request("mat_percent")
	imat_place		= request("mat_place")
	imaterial		= imat_Name&"!!^^"&imat_percent&"!!^^"&imat_place

	strSql = ""
	strSql = strSql & " INSERT INTO db_item.dbo.tbl_lotte_material(mallid, mallinfoDiv, itemid, material, regdate) VALUES " & VBCRLF
	strSql = strSql & " ('"&mallid&"', '"&infoDiv&"', '"&itemid&"', '"&imaterial&"', getdate()) "
	dbget.execute strSql
	response.write "<script>alert('등록 되었습니다');location.replace('/admin/etc/lotte/popPoomOK.asp?mallid="&mallid&"&infoDiv="&infoDiv&"&itemid="&itemid&"');</script>"
ElseIf mode = "U" Then
	imat_Name		= request("mat_Name")
	imat_percent	= request("mat_percent")
	imat_place		= request("mat_place")
	imaterial		= imat_Name&"!!^^"&imat_percent&"!!^^"&imat_place

	strSql = ""
	strSql = strSql & " UPDATE db_item.dbo.tbl_lotte_material SET " & VBCRLF
	strSql = strSql & " material = '"&imaterial&"' " & VBCRLF
	strSql = strSql & " WHERE mallid = '"&mallid&"' and mallinfoDiv = '"&infoDiv&"' and itemid = '"&itemid&"' "
	dbget.execute strSql
	response.write "<script>alert('수정 되었습니다');location.replace('/admin/etc/lotte/popPoomOK.asp?mallid="&mallid&"&infoDiv="&infoDiv&"&itemid="&itemid&"');</script>"
ElseIf mode = "" Then
	Dim material
	strSql = ""
	strSql = strSql & " SELECT TOP 1 mallid, mallinfoDiv, itemid, material, regdate " & VBCRLF
	strSql = strSql & " FROM db_item.dbo.tbl_lotte_material " & VBCRLF
	strSql = strSql & " WHERE mallid = '"&mallid&"' and mallinfoDiv = '"&infoDiv&"' and itemid = '"&itemid&"' "
	rsget.Open strSql, dbget, 1
	If Not(rsget.EOF or rsget.BOF) Then
		cnt			= 1
		material	= Split(rsget("material"),"!!^^")
		mat_Name	= material(0)
		mat_percent	= material(1)
		mat_place	= material(2)
	Else
		cnt	= 0
	End If
	rsget.close
End If
%>
<script>
function frm_check(){
	var frm = document.frm;
	if("<%=cnt%>" > 0){
		document.getElementById('mode').value = 'U';
	}else{
		document.getElementById('mode').value = 'I';
	}
	if(frm.mat_Name.value == ''){
		alert('원재료명을 입력하세요');
		frm.mat_Name.focus();
		return;
	}
	if(frm.mat_percent.value == ''){
		alert('함량을 입력하세요');
		frm.mat_percent.focus();
		return;
	}
	if(parseInt(document.frm.mat_percent.value) > 100){
		alert('함량은 100이하로 해주세요');
		document.frm.mat_percent.value = "";
		document.frm.mat_percent.focus();
		return;
	}

	if(frm.mat_place.value == ''){
		alert('원료원산지를 입력하세요');
		frm.mat_place.focus();
		return;
	}
	frm.submit();
}
function inputNumCom(){
	var keycode = event.keyCode;
	if( !((48 <= keycode && keycode <=57) || keycode == 13 || keycode == 46) ){
		alert("숫자만 입력 가능합니다.!");
		event.keyCode = 0;
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>품목 팝업</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<br><br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="popPoomOK.asp">
<input type= "hidden" name = "mode" 	value="<%=mode%>">
<input type= "hidden" name = "mallid" value="<%=mallid%>">
<input type= "hidden" name = "infoDiv" value="<%=infoDiv%>">
<input type= "hidden" name = "itemid" value="<%=itemid%>">
<col width="10%" />
<col width="%" />
<col width="10%" />
<col width="%" />
<col width="10%" />
<col width="%" />
<tr bgcolor="#FFFFFF"><td colspan="6">구성우선1</td></tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>원재료명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mat_Name" value="<%=mat_Name%>" ></td>
	<td>함량(%)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mat_percent" value="<%=mat_percent%>" maxlength="3" onkeypress="inputNumCom();" style="ime-mode:Disabled;"></td>
	<td>원료원산지</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mat_place" value="<%=mat_place%>"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td bgcolor="#FFFFFF" colspan="6">
		<input type="button" class="button" value="확인" onclick="javascript:frm_check();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->