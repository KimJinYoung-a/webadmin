<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/wconcept/wconceptcls.asp"-->
<%
Dim oWconcept, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm, Depth5Nm

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// ��� ����
Set oWconcept = new Cwconcept
	oWconcept.FPageSize 		= 100
	oWconcept.FCurrPage			= page
	oWconcept.FRectIsMapping	= isMapping
	oWconcept.FRectSDiv			= srcDiv
	oWconcept.FRectKeyword		= srcKwd
	oWconcept.FRectCDL			= request("cdl")
	oWconcept.FRectCDM			= request("cdm")
	oWconcept.FRectCDS			= request("cds")
	oWconcept.getTenwconceptCateList
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

	// wconcept ī�װ� ��Ī �˾�
	function popwconceptCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popwconceptCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>wconcept ī�װ� ����</strong></font></td>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>wconcept �ڵ�</option>
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
	<td colspan="4">wconcept ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
</tr>
<% If oWconcept.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oWconcept.FresultCount - 1
			Depth1Nm = oWconcept.FItemList(i).FLCategoryNameEn
			Depth2Nm = oWconcept.FItemList(i).FMCategoryNameEn
			Depth3Nm = oWconcept.FItemList(i).FSCategoryNameEn
			Depth4Nm = oWconcept.FItemList(i).FDCategoryNameEn
			cateAllNm 	= Depth1Nm &" > "& Depth2Nm &" > "& Depth3Nm &" > "& Depth4Nm
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oWconcept.FItemList(i).FCategoryCode),"#CCCCCC","#FFFFFF") %>">
	<td><%= oWconcept.FItemList(i).FtenCateLarge & oWconcept.FItemList(i).FtenCateMid & oWconcept.FItemList(i).FtenCateSmall %></td>
	<td><%= oWconcept.FItemList(i).FtenCDLName %></td>
	<td><%= oWconcept.FItemList(i).FtenCDMName %></td>
	<td><%= oWconcept.FItemList(i).FtenCDSName %></td>
	<td><%= oWconcept.FItemList(i).FItemcnt %></td>
	<% If oWconcept.FItemList(i).FCategoryCode="" OR isNull(oWconcept.FItemList(i).FCategoryCode) Then %>
	<td colspan="3"><input type="button" class="button" value="wconcept ī�� ��Ī" onClick="popwconceptCateMap('<%= oWconcept.FItemList(i).FtenCateLarge %>','<%= oWconcept.FItemList(i).FtenCateMid %>','<%= oWconcept.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td onClick="popwconceptCateMap('<%= oWconcept.FItemList(i).FtenCateLarge %>','<%= oWconcept.FItemList(i).FtenCateMid %>','<%= oWconcept.FItemList(i).FtenCateSmall %>','<%=oWconcept.FItemList(i).FCategoryCode%>')" style="cursor:pointer"><%= oWconcept.FItemList(i).FCategoryCode %></td>
	<td onClick="popwconceptCateMap('<%= oWconcept.FItemList(i).FtenCateLarge %>','<%= oWconcept.FItemList(i).FtenCateMid %>','<%= oWconcept.FItemList(i).FtenCateSmall %>','<%=oWconcept.FItemList(i).FCategoryCode%>')" style="cursor:pointer"><%= cateAllNm %></td>
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
		<% If oWconcept.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oWconcept.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oWconcept.StartScrollPage to oWconcept.FScrollCount + oWconcept.StartScrollPage - 1 %>
			<% If i > oWconcept.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oWconcept.HasNextScroll Then %>
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
<% Set oWconcept = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->