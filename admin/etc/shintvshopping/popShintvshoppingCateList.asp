<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim oShintvshopping, i, page, isMapping, srcDiv, srcKwd, orderby
Dim cateAllNm, cateAllkey
Dim Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm, Depth5Nm
Dim Depth1Key, Depth2Key, Depth3Key, Depth4Key, Depth5Key

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"
If orderby = ""	Then orderby = "1"

'// ��� ����
Set oShintvshopping = new CShintvshopping
	oShintvshopping.FPageSize 		= 20
	oShintvshopping.FCurrPage		= page
	oShintvshopping.FRectIsMapping	= isMapping
	oShintvshopping.FRectSDiv		= srcDiv
	oShintvshopping.FRectKeyword	= srcKwd
	oShintvshopping.FRectCDL		= request("cdl")
	oShintvshopping.FRectCDM		= request("cdm")
	oShintvshopping.FRectCDS		= request("cds")
	oShintvshopping.FRectOrderby	= orderby
	oShintvshopping.getTenShintvshoppingCateList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// �˻�
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// shintvshopping ī�װ� ��Ī �˾�
	function popShintvshoppingCateMap(cdl,cdm,cds,lgroup, mgroup, sgroup, dgroup, tgroup) {
		var pCM = window.open("popShintvshoppingCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&lgroup="+lgroup+"&mgroup="+mgroup+"&sgroup="+sgroup+"&dgroup="+dgroup+"&tgroup="+tgroup,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
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
	<font color="red"><strong>Shintvshopping ī�װ� ����</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		�ٹ����� <!-- #include virtual="/common/module/categoryselectbox.asp"--><br>
		���Ĺ�� :
		<select name="orderby" class="select">
			<option value="1" <%=chkIIF(orderby="1","selected","")%>>ī�װ���</option>
			<option value="2" <%=chkIIF(orderby="2","selected","")%>>��ǰ��</option>
		</select> /
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>�̸�Ī</option>
		</select> /
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>ī�װ���</option>
		</select> /
		�˻��� :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="�˻�" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
</tr>
</table>
</form>
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
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">�ٹ����� ī�װ�</td>
	<td colspan="4">Shintvshopping ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>Shintvshopping (�ѱ�)</td>
</tr>
<% If oShintvshopping.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		Dim spCateNm
		For i = 0 to oShintvshopping.FresultCount - 1
			Depth1Key = oShintvshopping.FItemList(i).FLgroup
			Depth2Key = oShintvshopping.FItemList(i).FMgroup
			Depth3Key = oShintvshopping.FItemList(i).FSgroup
			Depth4Key = oShintvshopping.FItemList(i).FDgroup
			Depth5Key = oShintvshopping.FItemList(i).FTgroup

			Depth1Nm = oShintvshopping.FItemList(i).FLgroupName
			Depth2Nm = oShintvshopping.FItemList(i).FMgroupName
			Depth3Nm = oShintvshopping.FItemList(i).FSgroupName
			Depth4Nm = oShintvshopping.FItemList(i).FDgroupName
			Depth5Nm = oShintvshopping.FItemList(i).FTgroupName

			cateAllNm 	= Depth1Nm &" > "& Depth2Nm & " > " & Depth3Nm & " > " & Depth4Nm & " > " & Depth5Nm
			cateAllkey	= Depth1Key & Depth2Key & Depth3Key & Depth4Key & Depth5Key
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(cateAllkey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oShintvshopping.FItemList(i).FtenCateLarge & oShintvshopping.FItemList(i).FtenCateMid & oShintvshopping.FItemList(i).FtenCateSmall %></td>
	<td><%= oShintvshopping.FItemList(i).FtenCDLName %></td>
	<td><%= oShintvshopping.FItemList(i).FtenCDMName %></td>
	<td><%= oShintvshopping.FItemList(i).FtenCDSName %></td>
	<td><%= oShintvshopping.FItemList(i).FItemcnt %></td>
	<% If cateAllkey="" OR isNull(cateAllkey) Then %>
	<td colspan="3"><input type="button" class="button" value="Shintvshopping ī�� ��Ī" onClick="popShintvshoppingCateMap('<%= oShintvshopping.FItemList(i).FtenCateLarge %>','<%= oShintvshopping.FItemList(i).FtenCateMid %>','<%= oShintvshopping.FItemList(i).FtenCateSmall %>','', '', '', '', '')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popShintvshoppingCateMap('<%= oShintvshopping.FItemList(i).FtenCateLarge %>','<%= oShintvshopping.FItemList(i).FtenCateMid %>','<%= oShintvshopping.FItemList(i).FtenCateSmall %>','<%=Depth1Key%>','<%=Depth2Key%>','<%=Depth3Key%>','<%=Depth4Key%>','<%=Depth5Key%>')" style="cursor:pointer"><%= cateAllkey %></td>
	<td title="<%=cateAllNm%>" onClick="popShintvshoppingCateMap('<%= oShintvshopping.FItemList(i).FtenCateLarge %>','<%= oShintvshopping.FItemList(i).FtenCateMid %>','<%= oShintvshopping.FItemList(i).FtenCateSmall %>','<%=Depth1Key%>','<%=Depth2Key%>','<%=Depth3Key%>','<%=Depth4Key%>','<%=Depth5Key%>')" style="cursor:pointer"><%= oShintvshopping.FItemList(i).FLastDepthNm %></td>
	<td><%=cateAllNm%></td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oShintvshopping.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oShintvshopping.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oShintvshopping.StartScrollPage to oShintvshopping.FScrollCount + oShintvshopping.StartScrollPage - 1 %>
			<% If i > oShintvshopping.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oShintvshopping.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<% Set oShintvshopping = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->