<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  [App]�۰���>>�Ϻ��÷���ǰ��� �󼼻�ǰ
' History : 2013.12.17 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/appColorCls.asp" -->
<%
Dim yyyymmdd, page, i, olist
page		= request("page")
yyyymmdd	= request("yyyymmdd")

If page = "" Then page = 1
	
SET olist = new AppColorList
	olist.FCurrPage		= page
	olist.FPageSize		= 50
	olist.Frectyyyymmdd	= yyyymmdd
	olist.sbDailyColoritemlist
%>
<script language="javascript">

var ichk;
ichk = 1;
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

//����,��뿩�� ����
function jsSortIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	frm = document.fitem;
	sValue = "";
	sortNo = "";
	isusing = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;

			if (frm.isusing[i].value ==''){
				alert('��뿩�θ� �����ϼ���.');
				frm.isusing[i].focus();
				return;
			}			
			if(!IsDigit(frm.sortNo[i].value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sortNo[i].focus();
				return;
			}
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
				// ���ļ���
				if (sortNo==""){
					sortNo = frm.sortNo[i].value;
				}else{
					sortNo =sortNo+","+frm.sortNo[i].value;
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
			if(!IsDigit(frm.sortNo.value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sortNo.focus();
				return;
			}
			sortNo 	=  frm.sortNo.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("������ ��ǰ�� �����ϴ�.");
		return;
	}
	document.frmSortIsusing.detailitemarr.value = sValue;
	document.frmSortIsusing.sortnoarr.value = sortNo;
	document.frmSortIsusing.isusingarr.value = isusing;
	document.frmSortIsusing.submit();
}

function jsIsusingChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			frm.isusing[i].value=selv;
		}
	}else{
		frm.isusing.value=selv;
	}
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function jsImgView(sImgUrl){
	var wImgView;

	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function pop_collectionitemreg(yyyymmdd){
	var pop_collectionitemreg = window.open('/admin/appmanage/pop_appDailyColorItemlist.asp?yyyymmdd='+yyyymmdd+'','pop_collectionitemreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_collectionitemreg.focus();
}

</script>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="yyyymmdd" value="<%=yyyymmdd%>">
</form>
<form name="frmSortIsusing" method="post" action="/admin/appmanage/appDailyColor_process.asp" style="margin:0px;">
	<input type="hidden" name="detailitemarr" value="">
	<input type="hidden" name="sortnoarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="yyyymmdd" value="<%=yyyymmdd%>">
	<input type="hidden" name="mode" value="sortisusingedit">
</form>
<tr>
	<td align="left">
	<% If olist.FResultCount > 0 Then %>
		<input class="button" type="button" id="btnEditSel" value="����/��뿩�� ����" onClick="jsSortIsusing();">
		&nbsp;&nbsp;
		�س������&��뿩�θ� �����Ͻ� �Ŀ� ��ư�� �����ּž� ���� �� �ݿ��� �Ϸ�˴ϴ�.		
	<% End If %>
	</td>
	<td align="right">
		<input type="button" name="btnBan" value="��ǰ���" onClick="pop_collectionitemreg('<%= yyyymmdd %>')" class="button">
	</td>
</tr>

</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=olist.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= olist.FTotalPage %></b>
		&nbsp;&nbsp;
		<font color="red"><b>���� 40�� ��ǰ���� ��� ���� �մϴ�.</b></font>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>�̹���</td>
	<td>��ǰ�ڵ�</td>
	<td>�Ǹſ���</td>
	<td>�������</td>
	<td>
		��뿩��
		<% If olist.FResultCount > 0 Then %>
		<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
		<% End If %>
	</td>
</tr>
<% If olist.FResultCount > 0 Then %>
<% For i = 0 to olist.fresultcount -1 %>
<% if olist.FColorList(i).FIsusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF" align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1" align="center">
<% end if %>
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FColorList(i).FItemid %>"></td>
	<td>
		<img src="<%= olist.FColorList(i).FimageSmall %>" width="50" height="50" onClick="jsImgView('<%=olist.FColorList(i).FimageSmall%>')" style="cursor:pointer" >
	</td>	
	<td><%= olist.FColorList(i).fitemid %></td>
	<td><%= olist.FColorList(i).fsellyn %></td>
	<td><input type="text" size="2" maxlength="2" name="sortNo" value="<%=olist.FColorList(i).FSortNo%>" class="text"></td>
	<td>
		<input type="hidden" value="<%=olist.FColorList(i).FIsusing%>" name="orgisusing">
		<% drawSelectBoxUsingYN "isusing", olist.FColorList(i).FIsusing %>
	</td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If olist.HasPreScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
			<% If (i > olist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(olist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="olist_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If olist.HasNextScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr height="80" bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</form>
</table>
<% 
SET olist = nothing 
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->