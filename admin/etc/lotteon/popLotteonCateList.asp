<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotteon/lotteonCls.asp"-->
<%
Dim oLotteon, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, lastDepthName

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// ��� ����
Set oLotteon = new CLotteon
	oLotteon.FPageSize 		= 20
	oLotteon.FCurrPage		= page
	oLotteon.FRectIsMapping	= isMapping
	oLotteon.FRectSDiv		= srcDiv
	oLotteon.FRectKeyword		= srcKwd
	oLotteon.FRectCDL			= request("cdl")
	oLotteon.FRectCDM			= request("cdm")
	oLotteon.FRectCDS			= request("cds")
	oLotteon.getTenLotteonCateList
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

	// Lotteon ī�װ��� ��Ī �˾�
	function popLotteonCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popLotteonCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&std_cat_id="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>Lotteon ī�װ��� ����</strong></font></td>
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
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>�̸�Ī</option>
		</select> /
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>Lotteon �ڵ�</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>ī�װ�����</option>
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
	<td colspan="4">�ٹ����� ī�װ���</td>
	<td colspan="4">Lotteon ī�װ���</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>�ڵ�</td>
	<td>ī�װ�����</td>
	<td>Lotteon (�ѱ�)</td>
</tr>
<% If oLotteon.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		Dim spCateNm
		For i = 0 to oLotteon.FresultCount - 1
			spCateNm = ""
			cateAllNm = oLotteon.FItemList(i).FStd_cat_nm
			spCateNm = Split(cateAllNm, ">")
			If Isarray(spCateNm) Then
				If Ubound(spCateNm) > 0 Then
					lastDepthName = spCateNm(Ubound(spCateNm))
				End If
			End If
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oLotteon.FItemList(i).FStd_cat_id),"#CCCCCC","#FFFFFF") %>">
	<td><%= oLotteon.FItemList(i).FtenCateLarge & oLotteon.FItemList(i).FtenCateMid & oLotteon.FItemList(i).FtenCateSmall %></td>
	<td><%= oLotteon.FItemList(i).FtenCDLName %></td>
	<td><%= oLotteon.FItemList(i).FtenCDMName %></td>
	<td><%= oLotteon.FItemList(i).FtenCDSName %></td>
	<% If oLotteon.FItemList(i).FStd_cat_id="" OR isNull(oLotteon.FItemList(i).FStd_cat_id) Then %>
	<td colspan="3"><input type="button" class="button" value="Lotteon ī�� ��Ī" onClick="popLotteonCateMap('<%= oLotteon.FItemList(i).FtenCateLarge %>','<%= oLotteon.FItemList(i).FtenCateMid %>','<%= oLotteon.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popLotteonCateMap('<%= oLotteon.FItemList(i).FtenCateLarge %>','<%= oLotteon.FItemList(i).FtenCateMid %>','<%= oLotteon.FItemList(i).FtenCateSmall %>','<%=oLotteon.FItemList(i).FStd_cat_id%>')" style="cursor:pointer"><%= oLotteon.FItemList(i).FStd_cat_id %></td>
	<td title="<%=cateAllNm%>" onClick="popLotteonCateMap('<%= oLotteon.FItemList(i).FtenCateLarge %>','<%= oLotteon.FItemList(i).FtenCateMid %>','<%= oLotteon.FItemList(i).FtenCateSmall %>','<%=oLotteon.FItemList(i).FStd_cat_id%>')" style="cursor:pointer"><%= lastDepthName %></td>
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
		<% If oLotteon.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oLotteon.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oLotteon.StartScrollPage to oLotteon.FScrollCount + oLotteon.StartScrollPage - 1 %>
			<% If i > oLotteon.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oLotteon.HasNextScroll Then %>
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
<% Set oLotteon = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->