<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �Ŵ�����
' History : ������ ����
'			2021.10.19 �ѿ�� ����(�����α� ����)
'			2022.09.08 ������ ����(isms�ɻ�� ���� ���ٱ��� üũ �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
IF application("Svr_Info")<>"Dev" THEN
	if Not(C_privacyadminuser) or Not(isVPNConnect) then
			response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [���ٱ���:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
			response.end
	end if
end if

Dim pid
	pid = requestCheckvar(Request("pid"),10)

%>
<script type="text/javascript">
<!--
	// ���� ���� �˾�
	function popAuthSelect()
	{
		window.open("pop_Menu_auth.asp", "popMenuAuth","width=700,height=400,scrollbars=no");
	}

	// �˾����� ���ñ��� �߰�
	function addAuthItem(psn,pnm,lsn,lnm)
	{
		var lenRow = tbl_auth.rows.length;

		// ������ ���� �ߺ� ��Ʈ ���� �˻�
		if(lenRow>1)	{
			for(l=0;l<document.all.part_sn.length;l++)	{
				if(document.all.part_sn[l].value==psn) {
					alert("�̹� ������ ������ �μ��Դϴ�.\n���� �μ��� �����ϰ� ������ �ٽ� �������ּ���.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.part_sn.value==psn) {
					alert("�̹� ������ ������ �μ��Դϴ�.\n���� �μ��� �����ϰ� ������ �ٽ� �������ּ���.");
					return;
				}
			}
		}

		// ���߰�
		var oRow = tbl_auth.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};

		// ���߰� (�μ�,���,������ư)
		var oCell1 = oRow.insertCell(0);
		var oCell2 = oRow.insertCell(1);
		var oCell3 = oRow.insertCell(2);

		oCell1.innerHTML = pnm + "<input type='hidden' name='part_sn' value='" + psn + "'>";
		oCell2.innerHTML = lnm + "<input type='hidden' name='level_sn' value='" + lsn + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle>";
	}

	// ���ñ��� ����
	function delAuthItem()
	{
		if(confirm("������ ������ �����Ͻðڽ��ϱ�?"))
			tbl_auth.deleteRow(tbl_auth.clickedRowIndex);
	}

	// ���˻� �� ����
	function submitForm()
	{
		var form = document.frm;

		if(!form.viewIdx.value||!IsDigit(form.viewIdx.value))
		{
			alert("ǥ�ü����� ������ �Է����ֽʽÿ�.");
			form.viewIdx.focus();
			return;
		}
		if(!form.menuname.value)
		{
			alert("�޴����� �Է����ֽʽÿ�.");
			form.menuname.focus();
			return;
		}
		if(!form.parentid.value)
		{
			alert("�����޴��� �������ֽʽÿ�.\n\n�ػ����޴��� ������� ��Ʈ�޴��� �������ֽʽÿ�.");
			form.parentid.focus();
			return;
		}

//		if(tbl_auth.rows.length<=0)
//		{
//			alert("�޴��� ������ �� �ִ� ������ [�߰�]��ư�� ���� �����Ͽ��ֽʽÿ�.");
//			return;
//		}

		if(confirm("�Է��� �������� �����Ͻðڽ��ϱ�?"))
		{
			form.action="menu_process.asp";
			form.submit();
		}
		else
		{
			return;
		}
	}
//-->
</script>
<script language="javascript" src="colorbox.js"></script>
<!-- ���� ���� ���� -->
<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td bgcolor="#FFFFFF"><img src="/images/icon_star.gif" align="absmiddle"> <b>�޴� �ű� ���</b></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">ǥ�ü���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="viewIdx" size="5" value=""></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="menuname" size="40" value=""></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴���(����)</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="menuname_en" size="40" value=""></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��ũURL</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="linkurl" size="60" value="">
		<input type="checkbox" name="useSslYN" value="Y"> SSL ���
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴����</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="lv1customerYN" value="Y" >LV1(������)
		<input type="checkbox" name="lv2partnerYN" value="Y" >LV2(��Ʈ������)
		<input type="checkbox" name="lv3InternalYN" value="Y" >LV3(��������)
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">ǥ�û���</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="prvColor" readonly style="background-color:'#000000';width:21px;height:21px;border:1px solid #606060;cursor:pointer;" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)">
		<input type="text" class="text_ro" name="menucolor" size="7" maxlength="7" value="" readonly onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)" style="cursor:pointer">
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�����޴�</td>
	<td bgcolor="#FFFFFF"><%=printRootMenuOption("parentid",pid, "NoAction")%></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��������</td>
	<td bgcolor="#FFFFFF">
		<table class=a>
		<tr>
			<td><%=getPartLevelInfo(0,"modi")%></td>
			<td valign="bottom"><input type="button" class="button" value="�߰�" onClick="popAuthSelect()"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="isUsing">
			<option value="Y" selected>���</option>
			<option value="N">����</option>
		</select>
	</td>
</tr>
<tr>
    <td bgcolor="#E6E6E6" align="center">(��������)</td>
    <td bgcolor="#EEEEEE">
        <% DrawAuthBox "divcd","2" %>
        (��ü, ���޻�, ����, ���� /admin/ ������ �ƴѰ�.)
    </td>
</tr>
<tr height="25">
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absmiddle"></a> &nbsp;
				<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absmiddle"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<!-- ���� ���� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
