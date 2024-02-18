<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim page, SearchKey, SearchString, pid, strUse, useSslYN, criticinfo
dim part_sn, level_sn, lv1customerYN, lv2partnerYN, lv3InternalYN

function getLevelStr(level_sn)
    getLevelStr = ""

	if IsNull(level_sn) then
		exit function
	end if

	''select case level_sn
	''	case 1
	''		getLevelStr = "������"
	''	case 2
	''		getLevelStr = "������"
	''	case 3
	''		getLevelStr = "��Ʈ����"
	''	case 4
	''		getLevelStr = "��Ʈ������"
	''	case 5
	''		getLevelStr = "��Ʈ�ӽ���"
	''	case 6
	''		getLevelStr = "��������"
	''	case 7
	''		getLevelStr = "��������"
	''	case 9
	''		getLevelStr = "����������ȸ"
	''	case else
	''		getLevelStr = "ERR(" & level_sn & ")"
	''end select

	select case level_sn
		case 1
			getLevelStr = "A"
		case 2
			getLevelStr = "B"
		case 3
			getLevelStr = "C"
		case 4
			getLevelStr = "D"
		case 5
			getLevelStr = "E"
		case 6
			getLevelStr = "F"
		case 7
			getLevelStr = "G"
		case 9
			getLevelStr = "H"
		case else
			getLevelStr = "ERR(" & level_sn & ")"
	end select
end function

	pid     = RequestCheckvar(Request("pid"),10)
	page    = RequestCheckvar(Request("page"),10)
	SearchKey = RequestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	strUse =   RequestCheckvar(Request("strUse"),10)
	useSslYN = RequestCheckvar(Request("useSslYN"),10)
	criticinfo = RequestCheckvar(Request("criticinfo"),10)
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

	oMenu.FPagesize = 300
	oMenu.FCurrPage = page
	oMenu.FRectsearchKey = searchKey
	oMenu.FRectsearchString = searchString
	oMenu.FRectPid = pid
	oMenu.FRectisUsing = strUse
	oMenu.FRectuseSslYN=useSslYN
	oMenu.FRectcriticinfo=criticinfo
	oMenu.FRectlv1customerYN = lv1customerYN
	oMenu.FRectlv2partnerYN = lv2partnerYN
	oMenu.FRectlv3InternalYN = lv3InternalYN
	oMenu.FRectPart_sn = part_sn
	oMenu.FRectLevel_sn = level_sn
	oMenu.GetMenuPrivList
%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="menu_priv_list.asp";
		document.frm.submit();
	}

	// �����޴� �̵�
	function goChild(pid)
	{
		document.frm.pid.value=pid;
		document.frm.action="menu_priv_list.asp";
		document.frm.submit();
	}


	// �޴� ������(����) ������ �̵�
	function goEdit(mid)
	{
	    //document.frm.mid.value=mid;
		//document.frm.page.value='<%= page %>';
		//document.frm.action="menu_edit.asp";
		//document.frm.submit();

	    var popwin=window.open('menu_edit.asp?mid='+mid,'popmenu_edit','width=900,height=700,scrollbars=yes,resizable=yes');
	    popwin.focus();

	}

	// �űԵ�� �������� �̵�
	function goAddItem()  {
		self.location="menu_add.asp?menupos=<%=menupos%>&pid=<%=pid%>";
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" action="menu_priv_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mid" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�����޴� <%=printRootMenuOption("pid",pid, "Action")%>
			&nbsp;
			���Ѻз� :
			<%= printPartOption("part_sn", part_sn) %>
			&nbsp;
			���ѵ�� :
			<%= printLevelOption("level_sn", level_sn) %>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�˻�">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		��뿩��
		<select class="select" name="strUse">
			<option value="all">��ü</option>
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		/ �˻�
		<select class="select" name="SearchKey">
			<option value="">::����::</option>
			<option value="id">�޴���ȣ</option>
			<option value="menuname">�޴���</option>
		</select>
		<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">
		/ SSL ����
		<select class="select" name="useSslYN">
			<option value="">��ü</option>
			<option value="Y" <%=CHKIIF(useSslYN="Y","selected","")%> >SSL ���</option>
			<option value="N" <%=CHKIIF(useSslYN="N","selected","")%> >SSL ������</option>
		</select>
		/
		�޴���� :
		<%' Call DrawSelectBoxCriticInfoMenu("criticinfo", criticinfo) %>
		<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(������)
		<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(��Ʈ������)
		<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(��������)
		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.strUse.value="<%=strUse%>";
		</script>

		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% if pid<>0 then %>
	<td><input type="button" class="button" value="�޴���Ʈ" onClick="goChild(0)"></td>
	<% end if %>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

* A : ������ ���� / B : ������ ���� / C : ��Ʈ���� ���� / D : ��Ʈ������ ���� / E : ��Ʈ�ӽ��� ���� / F : �������� ���� / G : �������� ���� / H : ��������ȸ����

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="50">
		�˻���� : <b><%=oMenu.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMenu.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">ID</td>
	<td>�����޴�</td>
	<td>�����޴�</td>
	<td width="80">��������</td>
	<td>�μ���ü</td>
	<td>�����ȭ</td>
	<td>������</td>
	<td>��������</td>
	<td>�¶���MD�</td>
	<td>�¶���MD����</td>
	<td>�¶���WD</td>
	<td>������</td>
	<td>�������κ���</td>
	<td>��������������</td>
	<td>���ȹ</td>
	<td>�ý���</td>
	<td>����</td>
	<td>CS</td>
	<td>�繫ȸ��</td>
	<td>�λ��ѹ�</td>
	<td>�����</td>
	<td>�߰�01</td>
	<td>�߰�02</td>
	<td>��Ÿ</td>
	<td width="50">����</td>
	<td width="50">���</td>
	<!--td width="50">SSL</td-->
	<!--td width="100">�޴����</td-->
	<td width="30">LV1<br>��<br>����</td>
	<td width="40">LV2<br>��Ʈ��<br>����</td>
	<td width="30">LV3<br>����<br>����</td>
	<td>����</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="50" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �޴��� �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oMenu.FResultCount - 1
%>
<% if (oMenu.FitemList(lp).Fmenu_isUsing = "Y") then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
<% else %>
<tr align="center" bgcolor="<%= adminColor("gray") %>" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="<%= adminColor("gray") %>";>
<% end if %>
	<td><%=oMenu.FitemList(lp).Fmenu_id%></td>
	<td align="left">
		<%
		response.Write "<a href='javascript:goChild(" & oMenu.FitemList(lp).Fmenu_id & ")'>" & oMenu.FitemList(lp).Fmenu_name_parent & "</a>"
		%>
	</td>
	<td align="left">
		<%
		response.Write oMenu.FitemList(lp).Fmenu_name
		%>
	</td>
	<td><%= oMenu.FitemList(lp).getOldMenuDivStr %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn1) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn16) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn14) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn22) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn11) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn21) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn12) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn23) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn13) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn24) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn30) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn7) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn9) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn10) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn8) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn20) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn17) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn33) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn25) %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_part_sn_etc %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_viewIdx %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_isUsing %></td>
	<!--td><%'oMenu.FitemList(lp).Fmenu_useSslYN %></td-->
	<!--td><%'GetCriticInfoMenuLevelName(oMenu.FitemList(lp).Fmenu_criticinfo) %></td-->
	<td><%= oMenu.FItemList(lp).Flv1customerYN %></td>
	<td><%= oMenu.FItemList(lp).Flv2partnerYN %></td>
	<td><%= oMenu.FItemList(lp).Flv3InternalYN %></td>
	<td><input type="button" value="����" class="button" onClick="goEdit(<%=oMenu.FitemList(lp).Fmenu_id%>)"></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="50" align="center">
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
