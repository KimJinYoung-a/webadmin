<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �Կ� ��û ���������
' History : 2012.03.13 ������ ����
'			2015.07.28 �ѿ�� ����(����� ���� �ִ� �κ� ��񿡼� ������. ����� ����. ��� ����&�߰�. ������ �űԷ� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim i, cdl, makerid, arrFileList, MDid, cdl_disp

Dim cPhotoreq, isUpdateDate
set cPhotoreq = new Photoreq
	cPhotoreq.fnReqno

	isUpdateDate = CDate("2016-12-19")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>

function request_modi(){
	var qqq;
	var frm = document.itemreg;
	var chk = 0;

	if (frm.request_no.value==''){
		alert('�Կ���û ��ȣ�� �Է��ϼ���.');
		frm.request_no.focus();
		return;
	}

	for(var i=0; i<frm.req_gubun.length; i++) {
		if(frm.req_gubun[i].checked){
			qqq = frm.req_gubun[i].id;
			chk++;
		}
	}
	if(confirm("����� ������ �ҷ����ðڽ��ϱ�?") == true) {
		location.href= 'request_modi.asp?req_no='+document.getElementById('request_no').value+'&gub='+qqq+'&sMode2=I&menupos=<%= menupos %>';
	} else {
		return false;
	}
}
function jsChkSubj(chk){

	if(chk=='5') {
		document.getElementById('detail').style.display = "block";
	} else {
		document.getElementById('detail').style.display = "none";
	}
}

function jsDefaultOpt(){
	if ($("#dopt").attr("checked")){
		$("#defaultOptTr").show();
	}else{
		$("#defaultOptTr").hide();
	}
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function fileupload()
{
	window.open('request_popupload2.asp','worker','width=420,height=200,scrollbars=yes');
}
function onlyNumber(){
	if((event.keyCode<48)||(event.keyCode>57))
		event.returnValue=false;
}
function form_check(){
	var frm = document.itemreg;
	var chk = 0;
	for(var i=0; i<frm.req_gubun.length; i++) {
		if(frm.req_gubun[i].checked) chk++;
	}
	if(chk == "0"){
		alert("�Կ� ������ �����ϼ���");
		return false;
	}

	if(frm.req_use.value == 0) {
		alert("�Կ��뵵 ������ �����ϼ���");
		frm.req_use.focus();
		return false;
	}
	if(frm.req_use.value == 1) {
		if(frm.req_use_detail.value == 0) {
			alert("�⺻ ���������� �����ϼ���");
			frm.req_use_detail.focus();
			return false;
		}
	}
	if (frm.prd_name.value == ""){
		alert("��ǰ���� �Է��ϼ���.");
		frm.prd_name.focus();
		return;
	}
//	if (frm.prd_type.value == ""){
//		alert("��ǰ���� �Է��ϼ���.");
//		frm.prd_type.focus();
//		return;
//	}
	if (frm.prd_type2.value == ""){
		alert("�� ��ǰ������ �Է��ϼ���.");
		frm.prd_type2.focus();
		return;
	}
	if(frm.prd_type2.value!=''){
		if (!IsDouble(frm.prd_type2.value)){
			alert('�� ��ǰ������ ���ڸ� �����մϴ�.');
			frm.prd_type2.focus();
			return;
		}
	}
	if (frm.import_level.value == "0"){
		alert("�߿䵵�� �Է��ϼ���.");
		frm.import_level.focus();
		return;
	}
	<% if not(session("ssAdminLsn")<="3") then %>
		if (frm.import_level.value == "4" || frm.import_level.value == "5"){
			alert('�߿䵵 S�� A�� ����,��Ʈ��,���Ӹ� ���� ���� �մϴ�.');
			return;
		}
	<% end if %>
	if(frm.req_department.value == "") {
		alert("��û�μ��� �����ϼ���");
		frm.req_department.focus();
		return false;
	}
	if(frm.req_cdl_disp.value == "") {
		alert("ī�װ��� �����ϼ���");
		frm.req_cdl_disp.focus();
		return false;
	}
//	if(frm.MDid.value == "00") {
//		alert("���MD�� �����ϼ���");
//		frm.MDid.focus();
//		return false;
//	}

	if (frm.req_etc1.value == ""){
		alert("��ǰ Ư¡ �� �ֿ� ���� ������ �Է��ϼ���.");
		frm.req_etc1.focus();
		return;
	}

	if(frm.itemid.value!=''){
		if (!IsNumbers(frm.itemid.value)){
			alert('��ǰ�ڵ带 ��Ȯ�ϰ� �Է��� �ּ���.');
			frm.itemid.focus();
			return;
		}
	}

	frm.action = "request_proc.asp";
	frm.submit();
}

function popdepartment(){
	var popwin = window.open('popdepartmentselect.asp','addreg','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function downphotoitemlist_sample(){
	var popwin = window.open('http://imgstatic.10x10.co.kr/offshop/sample/photo/photoitemlist_sample.xlsx','exceldown','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ����ǰ �߰� �˾�
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/photo_req/pop_itemAddInfo.asp", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<form name="itemreg" method="post">
<input type="hidden" name="mode" value="I">
<input type = "hidden" name = "req_no" value = "<%=cPhotoreq.Freq_no + 1%>">

<!-- 1.�Ϲ����� ��� �� ����-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="5">1.�Ϲ�����</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� ���� *</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="req_gubun" id="1" value="�ű�" onClick="document.getElementById('lyRequre').style.display='none';">�ű�</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="2" value="�߰��Կ�" onClick="document.getElementById('lyRequre').style.display='block';">�߰��Կ�</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="3" value="���Կ�" onClick="document.getElementById('lyRequre').style.display='block';">���Կ�</label>
		&nbsp;<label><input type="radio" name="req_gubun" id="4" value="�߰� ���� ��û��" onClick="document.getElementById('lyRequre').style.display='block';">�߰� ���� ��û��</label>
	</td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="4">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0" id="lyRequre" style="display:none;">
		<tr>
			<td align="left">
				�Կ���û no. <input type="text" name="request_no" value="" id="request_no" size="10" class="text">
				<input type="button" value="Ȯ��" class="button" onclick="request_modi();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ��뵵 ���� *</td>
	<td bgcolor="#FFFFFF" ><% call DrawPicGubun("req_use", "doc_status", "1") %></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<div id = "detail" bgcolor="#FFFFFF" style="display:none;" ><% call DrawPicGubun("req_use_detail", "doc_status_detail", "1") %></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ��(��ȹ����) *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_name" size="60" maxlength="30" class="text"></td>
</tr>
<input type="hidden" name="prd_type" size="60" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_type" size="60" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�� ��ǰ���� *</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" style="IME-MODE:disabled;" name="prd_type2" size="10" maxlength="10" class="text" onkeypress="onlyNumber();">&nbsp;(15�� �̻� �� ������ ���� �ʼ�)
		<!--<select name="prd_type2">
			<% for i = 1 to 10 %> 
			<option value="<%= i %>"><%= i %></option>
			<% next %>
		</select>
		(������ 10�� �̻��� ��� ��û���� ������ �÷��ּ���)-->
	</td>
</tr>
<input type="hidden" name="prd_price" size="10" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�)</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_price" size="10" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߿䵵 </td>
	<td bgcolor="#FFFFFF" colspan="4">
		<select name="import_level" class="select">
			<option value="0">--�߿䵵 ����--</option>

			<% if session("ssAdminLsn")<="3" then %>
				<option value="5">S</option>
				<option value="4">A</option>
			<% end if %>

			<option value="3">B</option>
			<option value="2">C</option>
			<option value="1">D</option>
		</select>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��û�μ�/ī�װ� *</td>
	<td bgcolor="#FFFFFF" >
		<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1">
		<tr>
			<td align="left">
				<input type="hidden" name="req_department" value="">
				<input type="hidden" name="MDid" value="">
				<div name="divdepartmentname" id="divdepartmentname"></div>
				<div name="divMDidname" id="divMDidname"></div>
				<input type="button" onclick="popdepartment();" value="�μ��˻�" class="button" >
				<!--<select name="req_department" class="select">
					<option value="">--�μ� ����--</option>
					<option value="MD">MD</option>
					<option value="MKT">MKT</option>
					<option value="ithinkso">ithinkso</option>
					<option value="off">off</option>
					<option value="JR">������ȹ</option>
					<option value="WD">WD</option>
					<option value="CT">������</option>
				</select>-->
			</td>
		</tr>
		</table>
	</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<!--����ī�� : <% call DrawCategoryLarge("req_category", cdl) %><br>-->
		����ī�װ� : <%= fnStandardDispCateSelectBox(1,cdl_disp, "req_cdl_disp", cdl_disp, "")%>
		<% 'call DrawCategoryLarge_disp("req_cdl_disp", cdl_disp) %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
	<td bgcolor="#FFFFFF" colspan="4"><%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ�<br>(��ǥ�� �����Է°���)</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" name="itemid" size="60" maxlength="128" class="text">
		<input type="button" value="��ǰ�߰�" onclick="addnewItem();" class="button">
	</td>
</tr>
<!-- ��ü��Ͻÿ��� ������(MD�� ��ϰ���) -->
<input type="hidden" name="req_date" size="10" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��� �Կ����� </td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="req_date" size="10" onClick="jsPopCal('req_date');" style="cursor:hand;"></td>
</tr>-->
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">÷�����ϵ��</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="button" onclick="downphotoitemlist_sample();" value="�Կ���û��ǰ����Ʈ ���� �ٿ�ε�" class="button">
		<br><br><input type="button" value="���Ͼ��ε�" class="button" onclick="fileupload();">
		(�ִ�20mb���� ���ε� �����ϸ�, ������ �ֹι�ȣ Ȥ�� ��ȭ��ȣ���� ���������� �� ��� ��ȭ������ ������ ����)
		<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
<%
	IF isArray(arrFileList) THEN
		For i =0 To UBound(arrFileList,2)
%>
		<tr>
			<td>
				<input type='hidden' name='doc_file' value='<%=arrFileList(0,i)%>'>
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<a href='<%=arrFileList(0,i)%>' target='_blank'>
				<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
			</td>
		</tr>
<%
		Next
		Response.Write "<input type='hidden' name='isfile' value='o'>"
	Else
		Response.write "<tr><td></td></tr>"
	End If
%>
		</table>
	</td>
</tr>
</table>
<Br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="2">2.�Կ�����(�ߺ����� ����)</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ʿ� �Կ���</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				<% call CheckBoxUseType("doc_use_type", "", "") %>
				<!-- <br />�⺻���� ��� ����, �ĸ�, ����, ��ü��, ��Ű���� �߿� �ʿ��Ͻ� �׸��� �������ּ���. -->
			</td>
		</tr>
		<tr id="defaultOptTr" style="display:none;">
			<td>
				<input type="checkbox" name="defaultOpt" value="901">����
				<input type="checkbox" name="defaultOpt" value="902">�ĸ�
				<input type="checkbox" name="defaultOpt" value="903">����
				<input type="checkbox" name="defaultOpt" value="904">��ü��
				<input type="checkbox" name="defaultOpt" value="905">��Ű����
			</td>
		</tr>
		</table>
	</td>
</tr>

<% If isUpdateDate > Date() Then %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� �Կ� ����</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr><% call CheckBoxUseType("doc_use_concept", "", "") %></tr>
		</table>
	</td>
</tr>
<% End If %>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ Ư¡ �� �ֿ� ���� ���� *</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc1" rows="18" class="textarea" style="width:100%"></textarea></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� ��ũ �� ���� URL </td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="req_url" size="60" maxlength="200" class="text" value="http://">
	</td>
</tr>
<input type="hidden" name="req_etc2" size="10" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� �� ���ǻ���</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc2" rows="5" class="textarea" style="width:100%"></textarea></td>
</tr>-->
</table>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#FFFFFF" height="30">
		<input type = "button" class="button" value="����" onclick="form_check();">
		<input type = "button" class="button" value="���" onClick="window.location='/admin/photo_req/request_list.asp?menupos=<%=menupos%>'">
	</td>
</tr>
</table>

</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->