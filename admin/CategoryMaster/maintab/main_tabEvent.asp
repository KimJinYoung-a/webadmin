<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� ī�װ�md�� �̵�/ �߰�/����
'	Description : ���������� �ǰ���
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_tabitem_cls.asp"-->
<%
dim page, cdl, isusing
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	
	if page="" then page=1


dim oip
	set oip = New Cmain_tabEvent_list
	oip.FCurrPage = page
	oip.FPageSize=20
	oip.FRectCDL = cdl
	oip.FRectIsUsing = isusing
	oip.Getmain_tabEvent

dim i
%>
<script language='javascript'>

	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
	
			if (frm.name.indexOf('frmBuyPrc')!= -1) {
	
				pass = ((pass)||(frm.cksel.checked));
			}
	
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
	
	function delitems(upfrm){
		if (!CheckSelected()){
			alert('�����̺�Ʈ�� �����ϴ�.');
			return;
		}
	
		var ret = confirm('���� �̺�Ʈ�� �����Ͻðڽ��ϱ�?');
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
					}
				}
			}
			upfrm.mode.value="del";
			upfrm.submit();
	
		}
	}
	
	function changeUsing(upfrm){
		if (!CheckSelected()){
			alert('��ǰ�� ������ �ּ���.');
			return;
		}
	
		if (upfrm.allusing.value=='Y'){
			var ret = confirm('�����Ͻ� ��ǰ�� ����� ���� �����մϴ�.');
		} else {
			var ret = confirm('�����Ͻ� ��ǰ�� ������ ���� �����մϴ�.');
		}
	
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
					}
				}
			}
			upfrm.cdl.value = Listfrm.cdl.value;
			upfrm.mode.value="isUsingValue";
			upfrm.submit();
	
		}
	}
	
	// ��������
	function changeSort(upfrm) {
		if (!CheckSelected()){
			alert('��ǰ�� ������ �ּ���.');
			return;
		}
		var ret = confirm('�����Ͻ� ��ǰ�� ������ �����Ͻ� ��ȣ�� �����Ͻðڽ��ϱ�?');
	
		if (ret){
			var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
						upfrm.sortNo.value = upfrm.sortNo.value + frm.sortNo.value + "," ;
					}
				}
			}
			upfrm.cdl.value = Listfrm.cdl.value;
			upfrm.mode.value="ChangeSort";
			upfrm.submit();
	
		}
	}
	
	function AddIttems(){
		var ret = confirm(arrFrm.evt_code.value + '�̺�Ʈ�� �߰��Ͻðڽ��ϱ�?');
		if (ret){
			arrFrm.evt_code.value = arrFrm.evt_code.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}
	
	function AddIttems2(){
		if (document.Listfrm.cdl.value == ""){
			alert("ī�װ��� ������ �ּ���!");
			document.Listfrm.cdl.focus();
		}
		else if (document.arrFrm.evt_codearr.value == ""){
			alert("�̺�Ʈ��ȣ��  �����ּ���!");
			document.arrFrm.evt_codearr.focus();
		}
		else if (confirm(arrFrm.evt_codearr.value + '�̺�Ʈ�� �߰��Ͻðڽ��ϱ�?')){
			arrFrm.evt_code.value = arrFrm.evt_codearr.value;
			arrFrm.cdl.value = Listfrm.cdl.value;
			arrFrm.mode.value="add";
			arrFrm.submit();
		}
	}

	function RefreshCaMDChoiceRec(){
		if (document.Listfrm.cdl.value == ""){
			alert("ī�װ��� ������ �ּ���!");
			document.Listfrm.cdl.focus();
		}
		 else{
				  var popwin = window.open('','refreshFrm','');
				  popwin.focus();
				  refreshFrm.target = "refreshFrm";
				  refreshFrm.cdl.value = document.Listfrm.cdl.value;
				  refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_tabEvent.asp";
				  refreshFrm.submit();
		 }
	}

	// ī�װ� ����� ���
	function changecontent(){}

</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="cdl">
</form>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		�Ǽ��� :
		<select name='cdl' class="select">
			<option value="">�����ϼ���</option>
			<option value=1 <% if cdl = "1" then response.write " selected" %>>������/���ǽ�</option>
			<option value=2 <% if cdl = "2" then response.write " selected" %>>Ű��Ʈ/���</option>
			<option value=3 <% if cdl = "3" then response.write " selected" %>>����</option>
			<option value=4 <% if cdl = "4" then response.write " selected" %>>�м�</option>
			<option value=5 <% if cdl = "5" then response.write " selected" %>>���̺�/Ű��</option>
			<option value=6 <% if cdl = "6" then response.write " selected" %>>����ä��</option>
		</select>
		&nbsp;/&nbsp;
		������� :
		<select name="isusing" onchange="this.form.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>��ü</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>���</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>������</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td colspan="2">
				<img src="/images/icon_reload.gif" onClick="RefreshCaMDChoiceRec()" style="cursor:pointer" align="absmiddle" alt="html�����">
				����Ʈ�� ����
			</td>
		</tr>
		</form>
		<form name="arrFrm" method="post" action="domaintabEvent.asp">
		<input type="hidden" name="cdl">
		<input type="hidden" name="mode">
		<input type="hidden" name="evt_code">
		<input type="hidden" name="sortNo">
		<tr>
			<td>
				<input type="button" value="�����̺�Ʈ ����" onClick="delitems(arrFrm)" class="button"> /
				<select name="allusing"  class="select">
					<option value="Y">���� ��� -> Y </option>
					<option value="N">���� ��� -> N </option>
				</select><input type="button" value="����" class="button" onclick="changeUsing(arrFrm);"> /
				<input type="button" value="��������" class="button" onclick="changeSort(arrFrm);">
			</td>
			<td colspan="2" align="right">
				<input type="text" name="evt_codearr" value="" size="80" class="input">
				<input type="button" value="�̺�Ʈ �߰�" onclick="AddIttems2()" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;�˻��� ��ǰ�� : <%=oip.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">ī�װ���</td>
	<td align="center">�̺�Ʈ�ڵ�</td>
	<td align="center">Image</td>
	<td align="center">�̺�Ʈ��</td>
	<td align="center">�̺�Ʈ�Ⱓ</td>
	<td align="center">����</td>
	<td align="center">�������</td>
	<td align="center">����</td>
</tr>
<% for i=0 to oip.FResultCount-1 %>
<form name="frmBuyPrc<%=i%>" method="post" action="" >
<input type="hidden" name="evt_code" value="<%= oip.FItemList(i).Fevt_code %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<% if  oip.FItemList(i).Fcdl = 1 then
		response.write "������/���ǽ�"
		elseif oip.FItemList(i).Fcdl = 2 then
		response.write "Ű��Ʈ/���"
		elseif 	oip.FItemList(i).Fcdl = 3 then
		response.write "����"
		elseif 	oip.FItemList(i).Fcdl = 4 then
		response.write "�м�"
		elseif 	oip.FItemList(i).Fcdl = 5 then
		response.write "���̺�/Ű��"
		elseif 	oip.FItemList(i).Fcdl = 6 then
		response.write "����ä��"
		end if
		%>
	</td>
	<td align="center"><%= oip.FItemList(i).Fevt_code %></td>
	<td align="center"><img src="<%= oip.FItemList(i).Fevt_bannerimg %>" height="100" ></td>
	<td align="center"><%= oip.FItemList(i).Fevt_name %></td>
	<td align="center"><%= left(oip.FItemList(i).Fevt_startdate,10) & "<br>~ " & left(oip.FItemList(i).Fevt_enddate,10) %></td>
	<td align="center"><input type="text" name="sortNo" value="<%= oip.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
	<td align="center"><%= oip.FItemList(i).Fisusing %></td>
	<td align="center">
	<%
		IF oip.FItemList(i).Fevt_state="7" AND datediff("d",oip.FItemList(i).Fevt_state,date()) >= 0 and datediff("d",oip.FItemList(i).Fevt_enddate,date()) <=0 THEN
			Response.write "����"
		ELSEIF oip.FItemList(i).Fevt_state="7" AND datediff("d",oip.FItemList(i).Fevt_enddate,date()) > 0 THEN
			Response.write "����"
		else
			Response.write oip.FItemList(i).Fevt_statedesc
		end if
	%>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oip.HasPreScroll then %>
		<a href="?page=<%= oip.StarScrollPage-1 %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oip.StarScrollPage to oip.FScrollCount + oip.StarScrollPage - 1 %>
		<% if i>oip.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oip.HasNextScroll then %>
		<a href="?page=<%= i %>&cdl=<%=cdl%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oip = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
