<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
Dim mode, NOTmakerid, sqlStr, makerid
Dim nMaker, page, itemidarr, isusingarr
page		= request("page")
mode		= request("mode")
NOTmakerid	= request("NOTmakerid")
itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
makerid		= request("makerid")

If page = "" Then page = 1
If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(c.userid) as userCNT, count(m.makerid) as notmakerCNT "
	sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c "
	sqlStr = sqlStr & " LEFT JOIN db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung as m on m.makerid = c.userid "
	sqlStr = sqlStr & " WHERE c.userid = '"&NOTmakerid&"' "
	rsget.Open sqlStr,dbget,1
	If rsget("userCNT") = 0 Then
		response.write "<script language='javascript'>alert('�ٹ����ٿ� ��ϵ� �귣�尡 �ƴմϴ�.');location.href='/admin/etc/onlyLotte_Not_In_Makerid.asp?menupos="&menupos&"';</script>"
	ElseIf rsget("notmakerCNT") > 0 Then
		response.write "<script language='javascript'>alert('�̹� ��ϵ� �귣�� �Դϴ�.');location.href='/admin/etc/onlyLotte_Not_In_Makerid.asp?menupos="&menupos&"';</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung (makerid, isusing, regdate, regid) VALUES "
		sqlStr = sqlStr & " ('"&NOTmakerid&"', 'Y',getdate(), '"&session("ssBctID")&"') "
		dbget.Execute sqlStr
		response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');location.href='/admin/etc/onlyLotte_Not_In_Makerid.asp?menupos="&menupos&"';</script>"
	End If
	rsget.Close
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	
	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_temp.dbo.tbl_Lotte_not_in_makerid_By_KimJinYoung SET "
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,lastupdateId = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE makerid ='" & itemidarr(i)&"'"
		dbget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');location.href='/admin/etc/onlyLotte_Not_In_Makerid.asp?menupos="&menupos&"';</script>"
End If

SET nMaker = new CLotte
	nMaker.FCurrPage					= page
	nMaker.FPageSize					= 20
	nMaker.FRectMakerId					= makerid
    nMaker.OnlyLotteNotUpdateMakeridList
%>
<script language='javascript'>
var ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

//���� �귣�� �����ϱ�
function jsIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	frm = document.fitem;
	sValue = "";
	isusing = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}

				// ��뿩��
				if (isusing==""){
					isusing = frm.isusing[i].value;
				}else{
					isusing =isusing+","+frm.isusing[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("������ �귣�尡 �����ϴ�.");
		return;
	}
	document.frmIsusing.itemidarr.value = sValue;
	document.frmIsusing.isusingarr.value = isusing;
	document.frmIsusing.mode.value = "U";
	document.frmIsusing.submit();
}

function insert_makerid()
{
	if(document.frm.NOTmakerid.value == "")
	{
		alert("�귣��ID�� �Է��ϼ���.");
		document.frm.NOTmakerid.focus();
		return;
	}
	document.frm.mode.value = "I";
	document.frm.submit();
}

function goPage(page){
    var frm = document.goPg;
    frm.page.value = page;
	frm.submit();
}
</script>
<form name="goPg" method="get" action="onlyLotte_Not_In_Makerid.asp" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
</form>

<form name="frmsearch" method="get" action="onlyLotte_Not_In_Makerid.asp" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : LotteCom</td>
		    <td rowspan="4" width="10%"><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			�귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<form name="frmIsusing" method="post" action="onlyLotte_Not_In_Makerid.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="onlyLotte_Not_In_Makerid.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<tr>
	<td>
		���� �귣�� <input type="text" class="text" name="NOTmakerid">&nbsp;&nbsp;<input type="button" class="button" value="����" onClick="insert_makerid()">
	</td>
	<td align="right">
		<% If nMaker.fresultcount >0 then %>	
			<input class="button" type="button" id="btnEditSel" value="���ܺ귣��Y/N ����" onClick="jsIsusing();">
	    <% End If %>
	</td>
</tr>
</form>
</table>
<strong>�� : <%= nMaker.FTotalCount%> ��</strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>�귣��ID</td>
	<td>�����</td>
	<td>�����</td>
	<td>����������</td>
	<td>����������</td>
	<td>���ܺ귣��Y/N</td>
</tr>
<% If nMaker.FResultCount > 0 Then %>
<% For i = 0 To nMaker.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nMaker.FItemlist(i).FMakerid %>"></td>
	<td><%=nMaker.FItemList(i).FMakerid%></td>
	<td><%=nMaker.FItemList(i).FRegdate%></td>
	<td><%=nMaker.FItemList(i).FRegid%></td>
	<td><%=nMaker.FItemList(i).FLastupdate%></td>
	<td><%=nMaker.FItemList(i).FLastupdateId%></td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <%=Chkiif(nMaker.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
			<option value="N" <%=Chkiif(nMaker.FItemList(i).FIsusing = "N","selected","")%> >N</option>
		</select>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nMaker.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
		<% If i>nMaker.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nMaker.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		��ϵ� �귣�尡 �����ϴ�
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nMaker = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->