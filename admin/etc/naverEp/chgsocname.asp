<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim makerid, itemid
makerid	= request("makerid")
itemid	= request("itemid")

Dim mode, chgitemid, sqlStr, chgsocname, chgsocname_kor, updatesocnamearr, updatesocname_korarr
Dim nItem, page, itemidarr, isusingarr
page				= request("page")
mode				= request("mode")
chgitemid			= Trim(request("chgitemid"))
chgsocname			= Trim(request("chgsocname"))
chgsocname_kor		= Trim(request("chgsocname_kor"))
itemidarr			= request("itemidarr")
isusingarr			= request("isusingarr")
itemid				= Trim(request("itemid"))
updatesocnamearr	= Trim(request("updatesocnamearr"))
updatesocname_korarr = Trim(request("updatesocname_korarr"))

If page = "" Then page = 1
If mode = "I" Then
	If isnumeric(chgitemid) = "False" Then
		response.write "<script language='javascript'>alert('�ٹ����ٿ� ��ϵ� ��ǰ�� �ƴմϴ�.');location.href='/admin/etc/naverEp/chgsocname.asp?menupos="&menupos&"';</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(i.itemid) as userCNT, count(m.itemid) as notitemCNT "
		sqlStr = sqlStr & " FROM db_AppWish.[dbo].[tbl_item] as i "
		sqlStr = sqlStr & " LEFT JOIN db_outmall.dbo.[tbl_EpShop_itemid_Socname] as m on m.itemid = i.itemid AND m.mallgubun = 'naverEP' "
		sqlStr = sqlStr & " WHERE i.itemid = '"&chgitemid&"' "
		rsCTget.CursorLocation = adUseClient
		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly
		If rsCTget("userCNT") = 0 Then
			response.write "<script language='javascript'>alert('�ٹ����ٿ� ��ϵ� ��ǰ�� �ƴմϴ�.');location.href='/admin/etc/naverEp/chgsocname.asp?menupos="&menupos&"';</script>"
		ElseIf rsCTget("notitemCNT") > 0 Then
			response.write "<script language='javascript'>alert('�̹� ��ϵ� ��ǰ �Դϴ�.');location.href='/admin/etc/naverEp/chgsocname.asp?menupos="&menupos&"';</script>"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_EpShop_itemid_Socname (itemid, mallgubun, socname, socname_kor, isusing, regdate, regid) VALUES "
			sqlStr = sqlStr & " ('"&chgitemid&"', 'naverep', '"&chgsocname&"', '"&chgsocname_kor&"', 'Y' ,getdate(), '"&session("ssBctID")&"') "
			sqlStr = sqlStr & vbCRLF & "; update I SET lastupdate=getdate() from [db_AppWish].dbo.tbl_item I where I.itemid="&chgitemid&""
			dbCTget.Execute sqlStr
			response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');location.href='/admin/etc/naverEp/chgsocname.asp?menupos="&menupos&"';</script>"
		End If
		rsCTget.Close
	End If
ElseIf mode = "U" Then
	Dim cnt, tmpIsusing
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)

	updatesocnamearr = split(updatesocnamearr,",")
	updatesocname_korarr = split(updatesocname_korarr,",")
	isusingarr	=  split(isusingarr,",")
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
		sqlStr = "UPDATE db_outmall.dbo.tbl_EpShop_itemid_Socname SET "
		sqlStr = sqlStr & " socname = '"&html2db(updatesocnamearr(i))&"'"
		sqlStr = sqlStr & " ,socname_kor = '"&html2db(updatesocname_korarr(i))&"'"
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'"
		sqlStr = sqlStr & " ,lastupdate = getdate()"
		sqlStr = sqlStr & " ,updateid = '"&session("ssBctID")&"'"
		sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
		sqlStr = sqlStr & ";"&VbCRLF

		sqlStr = sqlStr & " update I SET lastupdate=getdate() "
		sqlStr = sqlStr & " from [db_AppWish].dbo.tbl_item I "
		sqlStr = sqlStr & " INNER JOIN db_outmall.dbo.tbl_EpShop_itemid_Socname T "
		sqlStr = sqlStr & " ON T.idx="&itemidarr(i)
		sqlStr = sqlStr & " and T.itemid=I.itemid"
		dbCTget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');location.href='/admin/etc/naverEp/chgsocname.asp?menupos="&menupos&"';</script>"
End If

SET nItem = new epShop
	nItem.FCurrPage			= page
	nItem.FPageSize			= 20
	nItem.FRectItemid		= itemid
	nItem.FRectMakerid		= makerid
    nItem.EpshopChgItemidSocnameList
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

function jsSearchEpItemID(v) {
	var pEPI = window.open("/admin/itemmaster/itemlist.asp?itemid="+v,"pop_EPItemID","width=1400,height=600,scrollbars=yes,resizable=yes");
	pEPI.focus();
}
//���� ��ǰ �����ϱ�
function jsIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	var chg3depthnm;
	frm = document.fitem;
	sValue = "";
	isusing = "";
	updatesocname = "";
	updatesocname_kor = "";
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

				// updatesocname
				if (updatesocname==""){
					updatesocname = frm.updatesocname[i].value;
				}else{
					updatesocname =updatesocname+","+frm.updatesocname[i].value;
				}

				// updatesocname_kor
				if (updatesocname_kor==""){
					updatesocname_kor = frm.updatesocname_kor[i].value;
				}else{
					updatesocname_kor =updatesocname_kor+","+frm.updatesocname_kor[i].value;
				}

			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			isusing =  frm.isusing.value;
			updatesocname=  frm.updatesocname.value;
			updatesocname_kor=  frm.updatesocname_kor.value;
		}
	}
	if(chkSel<=0) {
		alert("������ ��ǰ�� �����ϴ�.");
		return;
	}
	document.frmIsusing.itemidarr.value = sValue;
	document.frmIsusing.isusingarr.value = isusing;
	document.frmIsusing.updatesocnamearr.value = updatesocname;
	document.frmIsusing.updatesocname_korarr.value = updatesocname_kor;
	document.frmIsusing.mode.value = "U";
	document.frmIsusing.submit();
}

function goPage(page){
    var frm = document.frmsearch;
    frm.page.value=page;
	frm.submit();
}
function frm_submit(){
	if(document.frm.chgitemid.value == ""){
		alert("��ǰ�ڵ带 �Է��ϼ���.");
		document.frm.chgitemid.focus();
		return;
	}
	if(document.frm.chgsocname_kor.value == ""){
		alert("�귣���(KR)�� �Է��ϼ���.");
		document.frm.chgsocname_kor.focus();
		return;
	}
	if(document.frm.chgsocname.value == ""){
		alert("�귣���(EN)�� �Է��ϼ���.");
		document.frm.chgsocname.focus();
		return;
	}
	document.frm.mode.value = "I";
	document.frm.submit();
}
function popUploadExcel(){
	var popwin;
	popwin = window.open("/admin/etc/naverEp/popUploadBrandname.asp", "popup_Brandname", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<!-- �˻� ���� -->
<form name="frmsearch" method="get" action="chgsocname.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : ���̹�EP</td>
		    <td rowspan="4" width="10%"><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td>
				�귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20">
				<input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;&nbsp;
				��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%=itemid%>" size="10">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<p /><table width="100%"><tr bgcolor="#000000"><td height="10"></td></tr></table><p />
<form name="frmIsusing" method="post" action="chgsocname.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="updatesocnamearr" value="">
	<input type="hidden" name="updatesocname_korarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>

<!-- �Է� ���� -->
<form name="frm" method="POST" action="chgsocname.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td width="10%" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td align="left">
		<input type="text" class="text" name="chgitemid">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td width="10%" bgcolor="<%= adminColor("tabletop") %>">�귣���(KR)</td>
	<td align="left">
		<input type="text" class="text" size="50" name="chgsocname_kor">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td width="10%" bgcolor="<%= adminColor("tabletop") %>">�귣���(EN)</td>
	<td align="left">
		<input type="text" class="text" size="50" name="chgsocname">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td colspan="2">
		<input type="button" class="button" value="����" onclick="frm_submit();">
		<input type="button" class="button" value="�������" onclick="popUploadExcel();" />
	</td>
</tr>
</table>
</form>
<!-- �Է� �� -->
<p /><table width="100%"><tr bgcolor="#000000"><td height="10"></td></tr></table><p />

<% If nItem.fresultcount >0 then %>
	<input class="button" type="button" id="btnEditSel" value="�귣���/���YN ����" onClick="jsIsusing();"><br />
<% End If %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
    <td>������</td>
	<td>�귣��ID</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣���(KR)</td>
	<td>�귣���(EN)</td>
	<td>���Y/N</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FIdx %>"></td>
	<td><%=nItem.FItemList(i).FMallgubun%></td>
	<td><%=nItem.FItemList(i).FMakerid%></td>
	<td><%=nItem.FItemList(i).FItemid%></td>
	<td>
		<input type="text" name="updatesocname" value="<%=nItem.FItemList(i).FSocname%>">
	</td>
	<td>
		<input type="text" name="updatesocname_kor" value="<%= nItem.FItemList(i).FSocname_kor %>">
	</td>
	<td>
		<select name="isusing" class="select">
			<option value="Y" <%=Chkiif(nItem.FItemList(i).FIsusing = "Y","selected","")%> >Y</option>
			<option value="N" <%=Chkiif(nItem.FItemList(i).FIsusing = "N","selected","")%> >N</option>
		</select>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nItem.StartScrollPage To nItem.FScrollCount + nItem.StartScrollPage - 1 %>
		<% If i>nItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		��ϵ� ��ǰ�� �����ϴ�
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->