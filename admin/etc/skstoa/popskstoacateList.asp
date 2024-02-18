<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/skstoa/skstoaCls.asp"-->
<%
Dim oSkstoa, i, page, isMapping, srcDiv, srcKwd, orderby
Dim cateAllNm, cateAllkey
Dim Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm
Dim Depth1Key, Depth2Key, Depth3Key, Depth4Key

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"
If orderby = ""	Then orderby = "1"

'// ��� ����
Set oSkstoa = new CSkstoa
	oSkstoa.FPageSize 		= 20
	oSkstoa.FCurrPage		= page
	oSkstoa.FRectIsMapping	= isMapping
	oSkstoa.FRectSDiv		= srcDiv
	oSkstoa.FRectKeyword	= srcKwd
	oSkstoa.FRectCDL		= request("cdl")
	oSkstoa.FRectCDM		= request("cdm")
	oSkstoa.FRectCDS		= request("cds")
	oSkstoa.FRectOrderby	= orderby
	oSkstoa.getTenSkstoaCateList
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

	// skstoa ī�װ� ��Ī �˾�
	function popSkstoaCateMap(cdl,cdm,cds,lgroup, mgroup, sgroup, dgroup, tgroup) {
		var pCM = window.open("popskstoaCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&lgroup="+lgroup+"&mgroup="+mgroup+"&sgroup="+sgroup+"&dgroup="+dgroup+"&tgroup="+tgroup,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>Skstoa ī�װ� ����</strong></font></td>
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
	<td colspan="4">Skstoa ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>Skstoa (�ѱ�)</td>
</tr>
<% If oSkstoa.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		Dim spCateNm
		For i = 0 to oSkstoa.FresultCount - 1
			Depth1Key = oSkstoa.FItemList(i).FLgroup
			Depth2Key = oSkstoa.FItemList(i).FMgroup
			Depth3Key = oSkstoa.FItemList(i).FSgroup
			Depth4Key = oSkstoa.FItemList(i).FDgroup

			Depth1Nm = oSkstoa.FItemList(i).FLgroupName
			Depth2Nm = oSkstoa.FItemList(i).FMgroupName
			Depth3Nm = oSkstoa.FItemList(i).FSgroupName
			Depth4Nm = oSkstoa.FItemList(i).FDgroupName

			cateAllNm 	= Depth1Nm &" > "& Depth2Nm & " > " & Depth3Nm & " > " & Depth4Nm
			cateAllkey	= Depth1Key & Depth2Key & Depth3Key & Depth4Key
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(cateAllkey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oSkstoa.FItemList(i).FtenCateLarge & oSkstoa.FItemList(i).FtenCateMid & oSkstoa.FItemList(i).FtenCateSmall %></td>
	<td><%= oSkstoa.FItemList(i).FtenCDLName %></td>
	<td><%= oSkstoa.FItemList(i).FtenCDMName %></td>
	<td><%= oSkstoa.FItemList(i).FtenCDSName %></td>
	<td><%= oSkstoa.FItemList(i).FItemcnt %></td>
	<% If cateAllkey="" OR isNull(cateAllkey) Then %>
	<td colspan="3"><input type="button" class="button" value="Skstoa ī�� ��Ī" onClick="popSkstoaCateMap('<%= oSkstoa.FItemList(i).FtenCateLarge %>','<%= oSkstoa.FItemList(i).FtenCateMid %>','<%= oSkstoa.FItemList(i).FtenCateSmall %>','', '', '', '', '')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popSkstoaCateMap('<%= oSkstoa.FItemList(i).FtenCateLarge %>','<%= oSkstoa.FItemList(i).FtenCateMid %>','<%= oSkstoa.FItemList(i).FtenCateSmall %>','<%=Depth1Key%>','<%=Depth2Key%>','<%=Depth3Key%>','<%=Depth4Key%>')" style="cursor:pointer"><%= cateAllkey %></td>
	<td title="<%=cateAllNm%>" onClick="popSkstoaCateMap('<%= oSkstoa.FItemList(i).FtenCateLarge %>','<%= oSkstoa.FItemList(i).FtenCateMid %>','<%= oSkstoa.FItemList(i).FtenCateSmall %>','<%=Depth1Key%>','<%=Depth2Key%>','<%=Depth3Key%>','<%=Depth4Key%>')" style="cursor:pointer"><%= oSkstoa.FItemList(i).FLastDepthNm %></td>
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
		<% If oSkstoa.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oSkstoa.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oSkstoa.StartScrollPage to oSkstoa.FScrollCount + oSkstoa.StartScrollPage - 1 %>
			<% If i > oSkstoa.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oSkstoa.HasNextScroll Then %>
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
<% Set oSkstoa = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->