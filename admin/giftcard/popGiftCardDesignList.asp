<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	dim oGiftcard, i, page
	dim cardItemid, groupDiv

	cardItemid	= request("cardid")
	groupDiv	= request("groupDiv")
	page		= request("page")
	if page="" then page=1

	'// ��� ����
	Set oGiftcard = new cGiftCard
	oGiftcard.FRectCardItemid=cardItemid
	oGiftcard.FRectGroupDiv=groupDiv
	oGiftcard.FPageSize = 10
	oGiftcard.FCurrPage = page
	oGiftcard.fGiftcard_DesignList
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">
<!--
	// ������ ���/����
	function editDesignInfo(cardid,dgnin) {
		if(!dgnin) dgnin="";
		self.location.href="popEditGiftCardDesign.asp?cardid="+cardid+"&designid="+dgnin;
	}

	// ������ �̵�
	function goPage(pg) {
		self.location.href="?cardid=<%=cardItemid%>&groupDiv=<%=groupDiv%>&page="+pg;
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>����Ʈī�� ������ ���</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;"><input type="button" class="button" value="+�űԵ��" onclick="editDesignInfo(<%=cardItemid%>)"></td>
</tr>
</table>
<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> ��ǰ�ڵ� : <strong><%=cardItemId%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>��ȣ</td>
	<td>�׷�</td>
	<td>�̹���</td>
	<td>�����θ�</td>
	<td>���</td>
</tr>
<% if oGiftcard.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="5" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	else
		for i=0 to oGiftcard.FresultCount-1
%>
<tr align="center" height="25" bgcolor="<%=chkIIF(oGiftcard.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oGiftcard.FItemList(i).FdesignId %></td>
	<td><%= oGiftcard.FItemList(i).fgetDesignGrpName %></td>
	<td><a href="javascript:editDesignInfo(<%=cardItemid%>,<%= oGiftcard.FItemList(i).FdesignId %>)"><img src="<%= oGiftcard.FItemList(i).FMMSThumb %>" border="0" width="50" height="50"></a></td>
	<td><a href="javascript:editDesignInfo(<%=cardItemid%>,<%= oGiftcard.FItemList(i).FdesignId %>)"><%= oGiftcard.FItemList(i).FcardDesignName %></a></td>
	<td><%= chkIIF(oGiftcard.FItemList(i).FisUsing="Y","���","����") %></td>
</tr>
<%
		next
	end if
%>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% if oGiftcard.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGiftcard.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oGiftcard.StartScrollPage to oGiftcard.FScrollCount + oGiftcard.StartScrollPage - 1 %>
			<% if i>oGiftcard.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oGiftcard.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</p>
<% Set oGiftcard = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->