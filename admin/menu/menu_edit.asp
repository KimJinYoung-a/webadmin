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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<%
IF application("Svr_Info")<>"Dev" THEN
	if Not(C_privacyadminuser) or Not(isVPNConnect) then
			response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [���ٱ���:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
			response.end
	end if
end if

Dim page, SearchKey, SearchString, midd, pid, i
	midd = requestCheckvar(Request("mid"),10)
	pid = requestCheckvar(Request("pid"),10)
	page = requestCheckvar(Request("page"),10)
	SearchKey = requestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")

if page="" then page=1
i=0

	'// ���� ����
	dim oMenu, lp
	Set oMenu = new CMenuList

	oMenu.FRectMid = midd

	oMenu.GetMenuCont

dim olog
Set olog = new CTenByTenMember
	olog.FPagesize = 50
	olog.FCurrPage = 1
	olog.frectmenuid=midd
	olog.getpartner_menu_log()
%>
<script type="text/javascript">
<!--
	// ���� ���� �˾�
	function popAuthSelect()
	{
		window.open("pop_Menu_auth.asp?mid=<%=midd%>", "popMenuAuth","width=700,height=400,scrollbars=no");
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

//���ѽ°� ���� ��� ����

		if(form.parentid.value=='0')
		{
			if(confirm("�����޴��� ������� ������ �°��Ͻðڽ��ϱ�?\n\n[Ȯ��]:��, [���]:�ƴϿ�"))
				form.childYN.value="Y";
			else
				form.childYN.value="N";
		}

     //   form.childYN.value="N";

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
<form name="frm" method="POST" action="" style="margin:0px;">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mid" value="<%=midd%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="childYN" value="N">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td bgcolor="#FFFFFF"><img src="/images/icon_star.gif" align="absmiddle"> <b>�޴� �󼼺���/����</b></td>
		</tr>
		</table>
	</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="4" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �޴��� �����ϴ�.</td>
</tr>
<%
	else
%>
<tr>
	<td width="100" bgcolor="#E6E6E6" align="center">�Ϸù�ȣ</td>
	<td bgcolor="#FFFFFF"><b><%=oMenu.FitemList(1).Fmenu_id%></b></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">ǥ�ü���</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="viewIdx" size="5" value="<%=oMenu.FitemList(1).Fmenu_viewIdx%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴���</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="menuname" size="40" value="<%=oMenu.FitemList(1).Fmenu_name%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴���(����)</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="menuname_en" size="40" value="<%=oMenu.FitemList(1).Fmenu_name_En%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��ũURL</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class='text' name="linkurl" size="60" value="<%=oMenu.FitemList(1).Fmenu_linkurl%>">
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴����</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="lv1customerYN" value="Y" <% if oMenu.FitemList(1).Flv1customerYN = "Y" then %>checked<% end if %> >LV1(������)
		<input type="checkbox" name="lv2partnerYN" value="Y" <% if oMenu.FitemList(1).Flv2partnerYN = "Y" then %>checked<% end if %> >LV2(��Ʈ������)
		<input type="checkbox" name="lv3InternalYN" value="Y" <% if oMenu.FitemList(1).Flv3InternalYN = "Y" then %>checked<% end if %> >LV3(��������)							
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�޴�����</td>
	<td bgcolor="#FFFFFF">
		<input type="checkbox" name="useSslYN" value="Y" <% if (oMenu.FitemList(1).Fmenu_useSslYN = "Y") then %>checked<% end if %> > SSL ���
		&nbsp;
		<input type="checkbox" name="saveLog" value="1" <% if (oMenu.FitemList(1).Fmenu_saveLog = "1") then %>checked<% end if %> > ���ӷα� ����
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">ǥ�û���</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="prvColor" readonly style="background-color:'<%=oMenu.FitemList(1).Fmenu_color%>';width:21px;height:21px;border:1px solid #606060;cursor:pointer;" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)">
		<input type="text" class='text_ro' name="menucolor" size="7" maxlength="7" value="<%=oMenu.FitemList(1).Fmenu_color%>" readonly onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)" style="cursor:pointer">
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�����޴�</td>
	<td bgcolor="#FFFFFF"><%=printRootMenuOption("parentid",oMenu.FitemList(1).Fmenu_parentid, "NoAction")%></td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" align="center">��������</td>
	<td bgcolor="#FFFFFF">
		<table class=a>
		<tr>
			<td><%=getPartLevelInfo(oMenu.FitemList(1).Fmenu_id,"modi")%></td>
			<td valign="bottom"><input type="button" class='button' value="�߰�" onClick="popAuthSelect()"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="isUsing">
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		<script language="javascript">frm.isUsing.value='<%=oMenu.FitemList(1).Fmenu_isUsing%>';</script>
	</td>
</tr>
<tr>
    <td bgcolor="#E6E6E6" align="center">(��������)</td>
    <td bgcolor="#EEEEEE">
        <% DrawAuthBox "divcd",oMenu.FitemList(1).Fmenu_divcd %>
        (��ü, ���޻�, ����, ���� /admin/ ������ �ƴѰ�.)
    </td>
</tr>
<%
	end if
%>
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

<% if olog.FResultCount>0 then %>
	<br>
	<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=olog.FtotalCount%></b>
			&nbsp;&nbsp;�� �ֱ� 50�Ǹ� ǥ�� �˴ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=60>�α׹�ȣ</td>
		<td>���泻��</td>
		<td width=100>������</td>
	</tr>
	<% for i=0 to olog.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= olog.FitemList(i).fidx %></td>
		<td align="left"><%= olog.FitemList(i).flogmsg %></td>
		<td>
			<%= olog.FitemList(i).fadminid %>
			<Br><%= left(olog.FitemList(i).fregdate,10) %>
			<Br><%= mid(olog.FitemList(i).fregdate,12,22) %>
		</td>
	</tr>
	<% next %>
	</table>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
