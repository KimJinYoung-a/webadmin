<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �Ŵ�����
' History : ������ ����
'			2021.10.19 �ѿ�� ����(�����α� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim page, SearchKey, SearchString, pid, strUse, useSslYN, criticinfo, saveLog
dim part_sn, level_sn, lv1customerYN, lv2partnerYN, lv3InternalYN
	pid     = RequestCheckvar(Request("pid"),10)
	page    = RequestCheckvar(Request("page"),10)
	SearchKey = RequestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	strUse =   RequestCheckvar(Request("strUse"),10)
	useSslYN = RequestCheckvar(Request("useSslYN"),10)
	criticinfo = RequestCheckvar(Request("criticinfo"),10)
	saveLog = RequestCheckvar(Request("saveLog"),10)
	part_sn = RequestCheckvar(Request("part_sn"),10)
	level_sn = RequestCheckvar(Request("level_sn"),10)
	lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
	lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
	lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
	
	if page="" then	page=1
	if pid="" then pid=0
	if strUse="" then strUse="Y"


	'// ���� ����
	dim oMenu, lp
	Set oMenu = new CMenuList

	oMenu.FPagesize = 100
	oMenu.FCurrPage = page
	oMenu.FRectsearchKey = searchKey
	oMenu.FRectsearchString = searchString
	oMenu.FRectPid = pid
	oMenu.FRectisUsing = strUse
	oMenu.FRectuseSslYN=useSslYN
	oMenu.FRectcriticinfo=criticinfo
	oMenu.FRectSaveLog = saveLog
	oMenu.FRectlv1customerYN = lv1customerYN
	oMenu.FRectlv2partnerYN = lv2partnerYN
	oMenu.FRectlv3InternalYN = lv3InternalYN
	oMenu.FRectPart_sn = part_sn
	oMenu.FRectLevel_sn = level_sn
	oMenu.GetMenuListNew
%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="menu_list.asp";
		document.frm.submit();
	}

	// �����޴� �̵�
	function goChild(pid)
	{
		document.frm.pid.value=pid;
		document.frm.action="menu_list.asp";
		document.frm.submit();
	}


	// �޴� ������(����) ������ �̵�
	function goEdit(mid)
	{
	    //document.frm.mid.value=mid;
		//document.frm.page.value='<%= page %>';
		//document.frm.action="menu_edit.asp";
		//document.frm.submit();

	    var popwin=window.open('menu_edit.asp?mid='+mid,'popmenu_edit','width=1200,height=800,scrollbars=yes,resizable=yes');
	    popwin.focus();

	}

	// �űԵ�� �������� �̵�
	function goAddItem()  {
		self.location="menu_add.asp?menupos=<%=menupos%>&pid=<%=pid%>";
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" action="menu_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mid" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�����޴� <%=printRootMenuOption("pid",pid, "Action")%>
			&nbsp;
			���Ѻз� :
			<%= printPartOption("part_sn", part_sn) %>
			&nbsp;
			���ѵ�� :
			<%= printLevelOption("level_sn", level_sn) %>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�˻�">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		��뿩�� :
		<select class="select" name="strUse">
			<option value="all">��ü</option>
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		&nbsp;
		�˻� :
		<select class="select" name="SearchKey">
			<option value="">::����::</option>
			<option value="id">�޴���ȣ</option>
			<option value="menuname">�޴���</option>
		</select>
		<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">

		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.strUse.value="<%=strUse%>";
		</script>

		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�޴���� :
			<%' Call DrawSelectBoxCriticInfoMenu("criticinfo", criticinfo) %>
			<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(������)
			<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(��Ʈ������)
			<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(��������)			
			&nbsp;
			SSL ���� :
			<select class="select" name="useSslYN">
				<option value="">��ü</option>
				<option value="Y" <%=CHKIIF(useSslYN="Y","selected","")%> >SSL ���</option>
				<option value="N" <%=CHKIIF(useSslYN="N","selected","")%> >SSL ������</option>
			</select>
			&nbsp;
			���ӷα� ���� :
			<select class="select" name="saveLog">
				<option value="">��ü</option>
				<option value="1" <%=CHKIIF(saveLog="1","selected","")%> >����</option>
				<option value="0" <%=CHKIIF(saveLog="0","selected","")%> >�������</option>
			</select>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% if pid<>0 then %>
	<td><input type="button" class="button" value="�޴���Ʈ" onClick="goChild(0)"></td>
	<% end if %>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onClick="goAddItem()">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%=oMenu.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMenu.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">ID</td>
	<td>�����޴�</td>
	<td>�����޴�</td>
	<td>��������</td>
	<td>��ũ</td>
	<td>����</td>
	<td width="30">����</td>
	<td width="30">���</td>
	<td width="30">LV1<br>��<br>����</td>
	<td width="40">LV2<br>��Ʈ��<br>����</td>
	<td width="30">LV3<br>����<br>����</td>
	<!--td width="30">SSL</td-->
	<!--td width="30">�α�<br>����</td-->
	<td width="80">����</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="16" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �޴��� �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oMenu.FResultCount - 1
%>
<tr align="center" bgcolor="<% if oMenu.FitemList(lp).Fmenu_isUsing="Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><%=oMenu.FitemList(lp).Fmenu_id%></td>
	<td align="left">
		&nbsp;
		<%
		response.Write "<a href='javascript:goChild(" & oMenu.FitemList(lp).Fmenu_id & ")'>" & oMenu.FitemList(lp).Fmenu_name_parent & "</a>"

		if Not(isNull(oMenu.FitemList(lp).Fmenu_cnt)) then
			response.Write "<span style='color:#AA5555;font-size:10px'> [" & oMenu.FitemList(lp).Fmenu_cnt & "]</span>"
		end if
		%>
	</td>
	<td align="left">
		&nbsp;
		<%
		response.Write oMenu.FitemList(lp).Fmenu_name
		%>
	</td>
	<td><%=oMenu.FitemList(lp).getOldMenuDivStr%></td>
	<td align="left"><%=oMenu.FitemList(lp).Fmenu_linkurl%></td>
	<td align="left"><%=getPartLevelInfo(oMenu.FitemList(lp).Fmenu_id, "list")%></td>
	<td><%=oMenu.FitemList(lp).Fmenu_viewIdx%></td>
	<td><%=oMenu.FitemList(lp).Fmenu_isUsing%></td>
	<!--td><%'GetCriticInfoMenuLevelName(oMenu.FitemList(lp).Fmenu_criticinfo) %></td-->
	<td><%=oMenu.FitemList(lp).Flv1customerYN%></td>
	<td><%=oMenu.FitemList(lp).Flv2partnerYN%></td>
	<td><%=oMenu.FitemList(lp).Flv3InternalYN%></td>
	<!--td><%'oMenu.FitemList(lp).Fmenu_useSslYN%></td-->
	<!--td><%'oMenu.FitemList(lp).Fmenu_saveLog%></td-->
	<td><input type="button" value="����" class="button" onClick="goEdit(<%=oMenu.FitemList(lp).Fmenu_id%>)"></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
	<!-- ������ ���� -->
	<%
		if oMenu.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oMenu.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oMenu.StartScrollPage to oMenu.FScrollCount + oMenu.StartScrollPage - 1

			if lp>oMenu.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oMenu.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
