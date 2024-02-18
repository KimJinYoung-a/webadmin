<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �Կ� ��û ���� & �� ������
' History : 2012.03.15 ������ ����
'			2015.07.28 �ѿ�� ����(����� ���� �ִ� �κ� ��񿡼� ������. ����� ����. ��� ����&�߰�. ������ �űԷ� ����)
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
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim gub, gubnm, i, udate, k
Dim cPhotoreq, rno, arrFileList, sMode2, isUpdateDate
Dim PhotoCnt

rno = request("req_no")
gub = request("gub")
sMode2 = request("sMode2")
udate = request("udate")

set cPhotoreq = new Photoreq
	cPhotoreq.FReq_no = rno
	cPhotoreq.fnPhotoreqUpdate
	PhotoCnt = cPhotoreq.fnGetPhotoUser
	arrFileList = cPhotoreq.fnGetFileList

if cPhotoreq.FTotalCount = 0 or cPhotoreq.FTotalCount="" or isnull(cPhotoreq.FTotalCount) then
	Call Alert_move("�ش� ������ �����ϴ�","request_list.asp?menupos="&menupos)
end if

	isUpdateDate = CDate("2016-12-19 18:30:00")
If cPhotoreq.FPhotoreqList(0).FReq_use = "" Then
	Call Alert_move("�ش� ������ �����ϴ�","request_list.asp?menupos="&menupos)
End If

dim copendata
set copendata = new Photoreq
	copendata.FReq_no = rno
	copendata.fnphoto_opendata
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function jsChkSubj(chk){
	if(chk=='5') {
		document.getElementById('detail').style.display = "block";
	} else {
		document.getElementById('detail').style.display = "none";
	}
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function fileupload(){
	window.open('request_popupload2.asp','worker','width=420,height=200,scrollbars=yes');
}
function jsDefaultOpt(){
	if ($("#dopt").attr("checked")){
		$("#defaultOptTr").show();
	}else{
		$("#defaultOptTr").hide();
	}
}

// �ϼ���ũ �� ���� URL
function AutoOpenurlInsert() {
	var f = document.all;

	var rowLen = f.divopenurl.rows.length;
//	if(rowLen > 5){
//		alert('�� �̻� �ø� �� �����ϴ�.');
//		return;
//	}
	var r  = f.divopenurl.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;
	c0.innerHTML = "";

	var inHtml = "<input type='hidden' name='openidx' value=''>"
	inHtml = inHtml + "<input type='text' name='openurl' value='' size=50 maxlength=512>"
	inHtml = inHtml + " <a href='' onclick='dateclearRow(this); return false;'>����</a>";
	//alert(inHtml)
	c0.innerHTML = inHtml;

	document.itemreg.lineopenurlCnt.value = f.divopenurl.rows.length;
}

function AutoInsert() {
	// ����׷��� ��������
	var vreq_photo = $.ajax({
		type: "POST",
		contentType: "application/x-www-form-urlencoded;charset=euc-kr",
		url: "/admin/photo_req/popsearchselect.asp",
		data: "searchtype=req_photo",
		dataType: "text",
		async: false
	}).responseText;

	// ��Ÿ�ϸ���Ʈ ��������
	var vreq_Stylist = $.ajax({
		type: "GET",
		contentType: "application/x-www-form-urlencoded;charset=euc-kr",
		url: "/admin/photo_req/popsearchselect.asp",
		data: "searchtype=req_Stylist",
		dataType: "text",
		async: false
	}).responseText;

	var f = document.all;

	var rowLen = f.div1.rows.length;
	if(rowLen > 5){
		alert('�� �̻� �ø� �� �����ϴ�.');
		return;
	}
	var r  = f.div1.insertRow(rowLen++);
	var c0 = r.insertCell(0);

	var Html;
	c0.innerHTML = "";

	var inHtml = "<input type='hidden' name='tmpcnt'>"
	inHtml = inHtml + "<select class='select' name='yyyy' >"
	inHtml = inHtml + "<option value=<%= year(date()) %> selected><%= year(date()) %></option>"
	<% for i=2002 to Year(now)+1 %>
	inHtml = inHtml + "<option value=<%= CStr(i) %> ><%= CStr(i) %></option>"
	<% next %>
	inHtml = inHtml + "</select>"
	inHtml = inHtml + "<select class='select' name='mm' >"
	inHtml = inHtml + "<option value='<%= month(date()) %>' selected><%= month(date()) %></option>"
	<% for i=1 to 12 %>
	inHtml = inHtml + "<option value='<%= Format00(2,i) %>' ><%= Format00(2,i) %></option>"
	<% next %>
	inHtml = inHtml + "</select>"
	inHtml = inHtml + "<select class='select' name='dd' >"
	inHtml = inHtml + "<option value='<%= day(date()) %>' selected><%= day(date()) %></option>"
	<% for i=1 to 31 %>
	inHtml = inHtml + "<option value='<%= Format00(2,i) %>' ><%= Format00(2,i) %></option>"
	<% next %>
	inHtml = inHtml + "</select> "
	inHtml = inHtml + "<select name='req_day_start' class='select' onchange=document.getElementById('sca1').value=this.value >"
	inHtml = inHtml + "<option value=''>-����-</option><option value='8'>10:00</option><option value='9'>10:30</option><option value='10'>11:00</option>"
	inHtml = inHtml + "<option value='11'>11:30</option><option value='12'>12:00</option><option value='13'>12:30</option><option value='14'>13:00</option>"
	inHtml = inHtml + "<option value='15'>13:30</option><option value='16'>14:00</option><option value='17'>14:30</option><option value='18'>15:00</option>"
	inHtml = inHtml + "<option value='19'>15:30</option><option value='20'>16:00</option><option value='21'>16:30</option><option value='22'>17:00</option>"
	inHtml = inHtml + "<option value='23'>17:30</option><option value='24'>18:00</option>"
	inHtml = inHtml + "</select> ~ "
	inHtml = inHtml + "<select name='req_day_end' class='select' onchange=document.getElementById('sca2').value=this.value>"
	inHtml = inHtml + "<option value=''>-����-</option><option value='8'>10:00</option><option value='9'>10:30</option><option value='10'>11:00</option>"
	inHtml = inHtml + "<option value='11'>11:30</option><option value='12'>12:00</option><option value='13'>12:30</option><option value='14'>13:00</option>"
	inHtml = inHtml + "<option value='15'>13:30</option><option value='16'>14:00</option><option value='17'>14:30</option><option value='18'>15:00</option>"
	inHtml = inHtml + "<option value='19'>15:30</option><option value='20'>16:00</option><option value='21'>16:30</option><option value='22'>17:00</option>"
	inHtml = inHtml + "<option value='23'>17:30</option><option value='24'>18:00</option>"
	inHtml = inHtml + "</select>"
	inHtml = inHtml + " <input type='text' name='comment' size=25>"
	inHtml = inHtml + "&nbsp;&nbsp;" + vreq_photo + "&nbsp;&nbsp;" + vreq_Stylist
	inHtml = inHtml + " <a href='' onclick='dateclearRow(this); return false;'>����</a>";
	//alert(inHtml)
	c0.innerHTML = inHtml;

	document.itemreg.lineCnt.value = f.div1.rows.length;
}
function goURL(){
	var uu = document.getElementById('lurl').value;
	window.open(uu);
}
function pop_print(){
	window.open('request_print.asp?req_no=<%=rno%>');
}
function onlyNumber(){
	if((event.keyCode<48)||(event.keyCode>57))
		event.returnValue=false;
}
function form_check(){
	var frm = document.itemreg;

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
	if(parseInt(document.getElementById('sca1').value) >= parseInt(document.getElementById('sca2').value)){
		alert("������ �ð��� �� �� �Ǿ����ϴ�.");
		document.getElementById('sca1').value = "1";
		document.getElementById('sca2').value = "2";
		return false;
	}
	if(frm.itemid.value!=''){
		if (!IsNumbers(frm.itemid.value)){
			alert('��ǰ�ڵ带 ��Ȯ�ϰ� �Է��� �ּ���.');
			frm.itemid.focus();
			return;
		}
	}
	<% If sMode2 <> "I" and (C_ADMIN_AUTH or C_CONTENTS_part) Then %>
		if (frm.tmpcnt != undefined){
			if (frm.tmpcnt.length>1){
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.yyyy[i].value == "") {
						alert('�Կ��Ͻ� �⵵�� �Է��� �ּ���.');
						frm.yyyy[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.mm[i].value == "") {
						alert('�Կ��Ͻ� ���� �Է��� �ּ���.');
						frm.mm[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.dd[i].value == "") {
						alert('�Կ��Ͻ� ���� �Է��� �ּ���.');
						frm.dd[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.req_day_start[i].value == "") {
						alert('�Կ��Ͻ� ���� �ð��� �Է��� �ּ���.');
						frm.req_day_start[i].focus();
						return false;
					}
				}
				for(var i=0; i < frm.tmpcnt.length; i++)	{
					if(frm.req_day_end[i].value == "") {
						alert('�Կ��Ͻ� ���� �ð��� �Է��� �ּ���.');
						frm.req_day_end[i].focus();
						return false;
					}
				}
			}else{
				if(frm.yyyy.value == "") {
					alert('�Կ��Ͻ� �⵵�� �Է��� �ּ���.');
					frm.yyyy.focus();
					return false;
				}
				if(frm.mm.value == "") {
					alert('�Կ��Ͻ� ���� �Է��� �ּ���.');
					frm.mm.focus();
					return false;
				}
				if(frm.dd.value == "") {
					alert('�Կ��Ͻ� ���� �Է��� �ּ���.');
					frm.dd.focus();
					return false;
				}
				if(frm.req_day_start.value == "") {
					alert('�Կ��Ͻ� ���� �ð��� �Է��� �ּ���.');
					frm.req_day_start.focus();
					return false;
				}
				if(frm.req_day_end.value == "") {
					alert('�Կ��Ͻ� ���� �ð��� �Է��� �ּ���.');
					frm.req_day_end.focus();
					return false;
				}
			}
		}
	<% end if %>

	//frm.lineCnt.value = document.all.div1.rows.length;
	frm.action = "/admin/photo_req/request_proc.asp";
	frm.submit();
}
function filedownload(idx){
	filefrm.file_idx.value = idx;
	filefrm.submit();
}
function clearRow(tdObj) {
	if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

function viewSche(vdate){
	window.open('request_cal_day.asp?getday='+vdate);
}

function dateclearRow(tdObj) {
	if ( itemreg.lineCnt.value < 2 ){
		alert('�ּ� 1���� �������� �Է��ϼž� �մϴ�.');
		return;
	}

	if(confirm("�����Ͻ� ���� �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
		var f = document.all;
		document.itemreg.lineCnt.value = f.div1.rows.length;
	} else {
		return false;
	}
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
	popwin = window.open("/admin/photo_req/pop_itemAddInfo.asp?smode=E", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<!-- ǥ ��ܹ� ����-->
<form name="itemreg" method="post">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="mode2" value="<%=sMode2%>">
<input type="hidden" name="userFont" value="<%=cPhotoreq.FPhotoreqList(0).FFontColor%>">
<input type="hidden" name="req_no" value="<%=rno%>">
<input type="hidden" id="sca1" name = "sca1" value = "1">
<input type="hidden" id="sca2" name = "sc12" value = "2">
<input type="hidden" name="req_id" value="<%=cPhotoreq.FPhotoreqList(0).FReq_id%>">
<input type="hidden" name="lineCnt" value="<%= cPhotoreq.FResultcount %>">

<!-- 1.�Ϲ����� ��� �� ����-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="#FFFFFF" height="30" colspan="5">1.�Ϲ�����</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� ���� *</td>
	<td bgcolor="#FFFFFF">
	<%
		If gub <> "" Then
			Select Case gub
				Case "2"	gubnm = "�߰��Կ�"
				Case "3"	gubnm = "���Կ�"
				Case "4"	gubnm = "�߰� ���� ��û��"
			End Select
			response.write gubnm
			response.write "<input type='hidden' name='req_gubunS' value='"&gub&"'>"
			response.write "<input type='hidden' name='req_gubun' value='"&gubnm&"'>"
		Else
			response.write cPhotoreq.FPhotoreqList(0).FReq_gubun
			response.write "<input type='hidden' name='req_gubun' value='"&cPhotoreq.FPhotoreqList(0).FReq_gubun&"'>"
		End If
	%>
	</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<%If sMode2 <> "I" Then %>
			<%= chkIIF(cPhotoreq.FPhotoreqList(0).FLoad_req <> "","�ҷ��� ��û�� No : "&cPhotoreq.FPhotoreqList(0).FLoad_req&"","") %>
		<%End If%>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="4">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0" id="lyRequre" style="display:none;">
		<tr>
			<td align="left">
				�Կ���û no. <input type="text" name="request_no" value="" size="10" class="text">
				<input type="button" value="Ȯ��" class="button" onclick="request_modi();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ��뵵 ���� *</td>
	<td bgcolor="#FFFFFF" ><% call DrawPicGubun2("req_use", "doc_status", cPhotoreq.FPhotoreqList(0).FReq_use) %></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<div id = "detail" bgcolor="#FFFFFF" <% If cPhotoreq.FPhotoreqList(0).FReq_use_detail = "" Then response.write "style=display:none;" End If %>><% call DrawPicGubun2("req_use_detail", "doc_status_detail" ,cPhotoreq.FPhotoreqList(0).FReq_use_detail) %></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ��(��ȹ����) *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_name" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_name%>" size="64" maxlength="30" class="text"></td>
</tr>
<input type="hidden" name="prd_type" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_type%>" size="60" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� *</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_type" value="<%'=cPhotoreq.FPhotoreqList(0).FPrd_type%>" size="60" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�� ��ǰ���� *</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" style="IME-MODE:disabled;" name="prd_type2" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_type2%>" size="10" maxlength="10" class="text" onkeypress="onlyNumber();">&nbsp;(15�� �̻� �� ������ ���� �ʼ�)
		<!--<select name="prd_type2">
			<% for i = 1 to 10 %> 
			<option value="<%= i %>"><%= i %></option>
			<% next %>
		</select>
		(������ 10�� �̻��� ��� ��û���� ������ �÷��ּ���)-->
	</td>
</tr>
<input type="hidden" name="prd_price" value="<%=cPhotoreq.FPhotoreqList(0).FPrd_price%>" size="10" maxlength="128" class="text">
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ǸŰ�(�Һ��ڰ�)</td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="prd_price" value="<%'=cPhotoreq.FPhotoreqList(0).FPrd_price%>" size="10" maxlength="128" class="text"></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߿䵵 </td>
	<td bgcolor="#FFFFFF" colspan="4">
		<select name="import_level" class="select">
			<option value="0">--�߿䵵 ����--</option>

			<% if session("ssAdminLsn")<="3" then %>
				<option value="5" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="5","selected","") %>>S</option>
				<option value="4" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="4","selected","") %>>A</option>
			<% end if %>

			<option value="3" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="3","selected","") %>>B</option>
			<option value="2" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="2","selected","") %>>C</option>
			<option value="1" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FImport_level="1","selected","") %>>D</option>
		</select>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��û�μ�/ī�װ� *</td>
	<td bgcolor="#FFFFFF" >
		<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1">
		<tr>
			<td align="left">
				<input type="hidden" name="req_department" value="<%= cPhotoreq.FPhotoreqList(0).FReq_department %>">
				<input type="hidden" name="MDid" value="<%= cPhotoreq.FPhotoreqList(0).FMDid %>">
				<div name="divdepartmentname" id="divdepartmentname"><%= getDepartmentALL(cPhotoreq.FPhotoreqList(0).FReq_department) %></div>
				<div name="divMDidname" id="divMDidname"><%= cPhotoreq.FPhotoreqList(0).FMDname %></div>
				<input type="button" onclick="popdepartment();" value="�μ��˻�" class="button" >
				<!--<select name="req_department" class="select">
					<option value="">--�μ� ����--</option>
					<option value="MD" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="MD","selected","") %>>MD</option>
					<option value="MKT" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="MKT","selected","") %>>MKT</option>
					<option value="ithinkso" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="ithinkso","selected","") %>>ithinkso</option>
					<option value="off" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="off","selected","") %>>off</option>
					<option value="JR" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="JR","selected","") %>>������ȹ</option>
					<option value="WD" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="WD","selected","") %>>WD</option>
					<option value="CT" <%'= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_department="CT","selected","") %>>������</option>
				</select>-->
			</td>
		</tr>
		</table>
	</td>
	<td bgcolor="#FFFFFF">
		<!--����ī�� : <% call DrawCategoryLarge("req_category", cPhotoreq.FPhotoreqList(0).FReq_category) %><br>-->
		����ī�װ� : <%= fnStandardDispCateSelectBox(1,cPhotoreq.FPhotoreqList(0).freq_cdl_disp, "req_cdl_disp", cPhotoreq.FPhotoreqList(0).freq_cdl_disp, "")%>
		<% 'call DrawCategoryLarge_disp("req_cdl_disp", cPhotoreq.FPhotoreqList(0).freq_cdl_disp) %>
	</td>
	<td bgcolor="#FFFFFF" colspan=2>
		�Կ���û�� : <%= cPhotoreq.FPhotoreqList(0).FReq_name %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
	<td bgcolor="#FFFFFF" colspan="4"><%	drawSelectBoxDesignerWithName "makerid", cPhotoreq.FPhotoreqList(0).FMakerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ�<br>(��ǥ�� �����Է°���)</td>
	<td bgcolor="#FFFFFF" colspan="4">
		<input type="text" name="itemid" size="60" maxlength="128" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FItemid%>" >
		<input type="button" value="��ǰ�߰�" onclick="addnewItem();" class="button">
<%
if cPhotoreq.FPhotoreqList(0).FItemid <> "" then
dim oitem
set oitem = new CItem
oitem.FPageSize = 1000
oitem.FCurrPage = 1
oitem.FRectItemid = cPhotoreq.FPhotoreqList(0).FItemid
oitem.GetItemList
%>
<% if oitem.FresultCount > 0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">��ǰID</td>
	<td align="center">�̹���</td>
	<td align="center">�귣��</td>
	<td align="center">��ǰ��</td>
</tr>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
	<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
</tr>
<% next %>
</table>
<% end if %>
<%
set oitem = nothing
end if
%>
	</td>
</tr>
<!-- ��ü��Ͻÿ��� ������(MD�� ��ϰ���) -->
<input type="hidden" name="req_date" size="10" value="<%=left(cPhotoreq.FPhotoreqList(0).FReq_date,10)%>" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��� �Կ����� </td>
	<td bgcolor="#FFFFFF" colspan="4"><input type="text" name="req_date" size="10" value="<%'=left(cPhotoreq.FPhotoreqList(0).FReq_date,10)%>" onClick="jsPopCal('req_date');" style="cursor:hand;"></td>
</tr>-->
<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF">
		÷�����ϵ��
	</td>
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
				<input type='hidden' name='doc_file' value='<%=arrFileList(1,i)%>'>
				<input type='hidden' name='doc_realfile' value='<%=arrFileList(3,i)%>'>
				<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<span class="a" onClick="filedownload(<%=arrFileList(0,i)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,i),"http://",""),"/")(4)%></span>
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
<br>
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
				<% call CheckBoxUseType("doc_use_type", rno, "1") %>
			</td>
		<%
			Dim odefault, isOptExist, defaultoptArr, dOptdispYn
			set odefault = new Photoreq
				isOptExist = odefault.getdefaultOpt(rno)

				IF isArray(isOptExist) THEN
					dOptdispYn = "Y"
				Else
					dOptdispYn = "N"
				End If

				IF dOptdispYn = "Y" THEN
					For k =0 To UBound(isOptExist,2)
						defaultoptArr = defaultoptArr & isOptExist(1,k) & "," 
					Next
				Else
					defaultoptArr = ""
				End If
		%>
			<tr id="defaultOptTr" <%= Chkiif(dOptdispYn="Y", "style='display:block;'", "style='display:none;'") %>style="display:block;">
				<td>
					<input type="checkbox" name="defaultOpt" value="901" <%= Chkiif(instr(defaultoptArr, "901") > 0, "checked", "") %>>����
					<input type="checkbox" name="defaultOpt" value="902" <%= Chkiif(instr(defaultoptArr, "902") > 0, "checked", "") %>>�ĸ�
					<input type="checkbox" name="defaultOpt" value="903" <%= Chkiif(instr(defaultoptArr, "903") > 0, "checked", "") %>>����
					<input type="checkbox" name="defaultOpt" value="904" <%= Chkiif(instr(defaultoptArr, "904") > 0, "checked", "") %>>��ü��
					<input type="checkbox" name="defaultOpt" value="905" <%= Chkiif(instr(defaultoptArr, "905") > 0, "checked", "") %>>��Ű����
				</td>
			</tr>
		<%
				SET odefault = nothing
		%>
		</tr>
		</table>
	</td>
</tr>
<% If isUpdateDate >= cPhotoreq.FPhotoreqList(0).FReq_regdate Then %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� �Կ� ����</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr><% call CheckBoxUseType("doc_use_concept", rno, "2") %></tr>
		</table>
	</td>
</tr>
<% End If %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ Ư¡ �� �ֿ� ���� ���� *</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc1" rows="18" class="textarea" style="width:100%"><%=cPhotoreq.FPhotoreqList(0).FReq_etc1%></textarea></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� ��ũ �� ���� URL </td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="req_url" size="60" maxlength="200" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FReq_url%>" id="lurl">
		<!--<input type="button" class="button" value="�ٷΰ���" onclick="goURL();">-->
	</td>
</tr>
<input type="hidden" name="req_etc2" size="10" value="<%=cPhotoreq.FPhotoreqList(0).FReq_etc2%>" >
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� �� ���ǻ���</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="req_etc2" rows="5" class="textarea" style="width:100%"><%=cPhotoreq.FPhotoreqList(0).FReq_etc2%></textarea></td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�Խñ� ��뿩��</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="use_yn" value="Y" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FUse_yn="Y","checked","") %>>Y
		<input type="radio" name="use_yn" value="N" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FUse_yn="N","checked","") %>>N
	</td>
</tr>
</table>

<% If sMode2 <> "I" and (C_ADMIN_AUTH or C_CONTENTS_part) Then %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="4">3. �����Ȳ <input type="button" value="�Կ� ������ ����" class="button" onclick="viewSche('<%=Left(now(),10)%>');"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select name="req_status" class="select">
				<option value="0">--������¼���--</option>
				<option value="4" <%= chkIIF(gub = "4","selected","") %> <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "4","selected","") %> >�߰����� ��û</option>
				<option value="1" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "1","selected","") %> >�Կ������� ����</option>
				<option value="2" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "2","selected","") %> >�Կ���</option>
				<option value="3" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "3","selected","") %> >�Կ��Ϸ�</option>
				<option value="9" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FReq_status = "9","selected","") %> >��������</option>
			</select>
			<input type="checkbox" name="fontColor" value="R" <%= chkIIF(cPhotoreq.FPhotoreqList(0).FFontColor = "R","checked","") %>>������� �ؽ�Ʈ ���� �ٲ�(��û�ڿ��� �߰����� ��û �˸��ÿ��� üũ�ϼ���!)
		</td>
	</tr>

	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� Ȯ�� �Ͻ�</td>
		<td bgcolor="#FFFFFF" width="70%">
			<div id="lyRequre1">
				<table id="div1" class="a">
				<%
				dim vstart_date, vend_date

				For i = 0 to cPhotoreq.FResultcount -1
				
				if cPhotoreq.FPhotoreqList(i).FStart_date="" or isnull(cPhotoreq.FPhotoreqList(i).FStart_date) then
					vstart_date=date()
				else
					vstart_date=cPhotoreq.FPhotoreqList(i).FStart_date
				end if
				%>
					<tr>
						<td>
							<input type="hidden" name="tmpcnt">
							<% DrawOneDateBoxdynamic "yyyy", year(vstart_date), "mm", month(vstart_date), "dd", day(vstart_date), "", "", "", "" %>
							<select name="req_day_start" class="select" onchange="document.getElementById('sca1').value=this.value">
								<option value="">-����-</option>
								<option value="8" <% if cstr("10:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>10:00</option>
								<option value="9" <% if cstr("10:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>10:30</option>
								<option value="10" <% if cstr("11:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>11:00</option>
								<option value="11" <% if cstr("11:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>11:30</option>
								<option value="12" <% if cstr("12:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>12:00</option>
								<option value="13" <% if cstr("12:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>12:30</option>
								<option value="14" <% if cstr("13:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>13:00</option>
								<option value="15" <% if cstr("13:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>13:30</option>
								<option value="16" <% if cstr("14:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>14:00</option>
								<option value="17" <% if cstr("14:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>14:30</option>
								<option value="18" <% if cstr("15:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>15:00</option>
								<option value="19" <% if cstr("15:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>15:30</option>
								<option value="20" <% if cstr("16:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>16:00</option>
								<option value="21" <% if cstr("16:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>16:30</option>
								<option value="22" <% if cstr("17:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>17:00</option>
								<option value="23" <% if cstr("17:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>17:30</option>
								<option value="24" <% if cstr("18:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FStart_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FStart_date)) ) then response.write " selected" %>>18:00</option>
							</select>
							~
							<select name="req_day_end" class="select" onchange="document.getElementById('sca2').value=this.value">
								<option value="">-����-</option>
								<option value="8" <% if cstr("10:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>10:00</option>
								<option value="9" <% if cstr("10:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>10:30</option>
								<option value="10" <% if cstr("11:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>11:00</option>
								<option value="11" <% if cstr("11:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>11:30</option>
								<option value="12" <% if cstr("12:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>12:00</option>
								<option value="13" <% if cstr("12:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>12:30</option>
								<option value="14" <% if cstr("13:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>13:00</option>
								<option value="15" <% if cstr("13:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>13:30</option>
								<option value="16" <% if cstr("14:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>14:00</option>
								<option value="17" <% if cstr("14:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>14:30</option>
								<option value="18" <% if cstr("15:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>15:00</option>
								<option value="19" <% if cstr("15:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>15:30</option>
								<option value="20" <% if cstr("16:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>16:00</option>
								<option value="21" <% if cstr("16:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>16:30</option>
								<option value="22" <% if cstr("17:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>17:00</option>
								<option value="23" <% if cstr("17:30")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>17:30</option>
								<option value="24" <% if cstr("18:00")= cstr( Format00(2,hour(cPhotoreq.FPhotoreqList(i).FEnd_date)) & ":" & Format00(2,minute(cPhotoreq.FPhotoreqList(i).FEnd_date)) ) then response.write " selected" %>>18:00</option>
							</select>
							<input type='text' name='comment' value='<%= cPhotoreq.FPhotoreqList(i).fcomment %>' size=25>
							&nbsp;<% call SelectUser("1", "req_photo", ""&cPhotoreq.FPhotoreqList(i).FReq_photo&"") %>
							&nbsp;<% call SelectUser("2", "req_Stylist", ""&cPhotoreq.FPhotoreqList(i).FReq_stylist&"") %>
							<a href='' onclick='dateclearRow(this); return false;'>����</a>
						</td>
					</tr>
				<%
				Next
				%>
				</table>
			</div>
		</td>
		<td bgcolor="#FFFFFF"><input type="button" value="�Կ��Ͻ��߰�" onClick="AutoInsert()" class="button"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�ڸ�Ʈ �ۼ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="req_comment" size="60" maxlength="128" class="text" value="<%=cPhotoreq.FPhotoreqList(0).FReq_comment%>"></td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">SMS ����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="checkbox" name="req_SMS" value="Y">���MD���� SMS���� (MD���� �� �� ���, �Կ���û�ڿ��� ����)
		</td>
	</tr>
	</table>

<% else %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="4">3. �����Ȳ</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">����</td>
		<td bgcolor="#FFFFFF" colspan="3">
	<%
		Select Case cPhotoreq.FPhotoreqList(0).FReq_status
			Case "1"	response.write "�Կ������� ����"
			Case "2"	response.write "�Կ���"
			Case "3"	response.write "�Կ��Ϸ�"
			Case "4"	response.write "�߰����� ��û"
			Case "10"	response.write "��������"
		End Select
	%>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�Կ� Ȯ�� �Ͻ�</td>
		<td bgcolor="#FFFFFF" colspan=3>
				<table class="a">
	<%
		For i = 0 to cPhotoreq.FResultcount -1
	%>
				<tr>
					<td>
						<font color="BLUE">���� : <%=cPhotoreq.FPhotoreqList(i).FStart_date%></font> ~  <font color="RED">���� : <%=cPhotoreq.FPhotoreqList(i).FEnd_date%></font>
						&nbsp;���� : <% call SelectUser2("1", ""&cPhotoreq.FPhotoreqList(i).FReq_photo&"") %>
						&nbsp;��Ÿ�� : <% call SelectUser2("2", ""&cPhotoreq.FPhotoreqList(i).FReq_stylist&"") %>
					</td>
				</tr>
	<%
		Next
	%>
				</table>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�ڸ�Ʈ �ۼ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><%=cPhotoreq.FPhotoreqList(0).FReq_comment%></td>
	</tr>
	</table>
<% End If %>

<%
' ���°��� �Կ��Ϸ� �̰ų� �������� �ϰ��
if cPhotoreq.FPhotoreqList(0).FReq_status="3" or cPhotoreq.FPhotoreqList(0).FReq_status="9" then
%>
	<Br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left">
		<td bgcolor="#FFFFFF" height="30" colspan="3">4. ������������</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">�ϼ���ũ �� ���� URL</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="lineopenurlCnt" value="0">
			<div id="lyopenurl">
				<table id="divopenurl" class="a">
				<%
				For i = 0 to copendata.FResultcount -1
				%>
					<tr>
						<td>
							<input type='hidden' name='openidx' value='<%= copendata.FPhotoreqList(i).fopenidx %>'>
							<input type='text' name='openurl' value='<%= copendata.FPhotoreqList(i).fopenurl %>' size=50 maxlength=512>
							<a href='' onclick='dateclearRow(this); return false;'>����</a>
						</td>
					</tr>
				<%
				Next
				%>
				</table>
			</div>
		</td>
		<td bgcolor="#FFFFFF"><input type="button" value="�ϼ���ũ�߰�" onClick="AutoOpenurlInsert()" class="button"></td>
	</tr>
	</table>
<% end if %>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="#FFFFFF" height="30">
		<input type="button" class="button" value=" �� �� " onclick="form_check();">
		<input type="button" class="button" value=" �� �� " onClick="window.location='/admin/photo_req/request_list.asp?menupos=<%=menupos%>'">
		<input type="button" class="button" value=" �� �� " onClick="pop_print();">		
	</td>
</tr>
</table>

</form>
<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/photo_req/photo_req_download2.asp" target="fileiframe">
<input type="hidden" name="brd_sn" value="<%=rno%>">
<input type="hidden" name="file_idx" value="">
</form>
<%
set cPhotoreq=nothing
set copendata = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->