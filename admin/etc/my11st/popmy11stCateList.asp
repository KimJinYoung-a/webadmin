<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/my11st/my11stcls.asp"-->
<%
Dim oMy11st, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, matchCateNm
Dim Depth1Nm, Depth2Nm, Depth3Nm

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// ��� ����
Set oMy11st = new CMy11st
	oMy11st.FPageSize 	= 20
	oMy11st.FCurrPage	= page
	oMy11st.FRectIsMapping	= isMapping
	oMy11st.FRectSDiv		= srcDiv
	oMy11st.FRectKeyword	= srcKwd
	oMy11st.FRectCDL		= request("cdl")
	oMy11st.FRectCDM		= request("cdm")
	oMy11st.FRectCDS		= request("cds")
	oMy11st.getTenmy11stCateList
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

	// 11���� ī�װ� ��Ī �˾�
	function popCjCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popmy11stCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>11���� ī�װ� ����</strong></font></td>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>11���� �ڵ�</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>10x10��ī�װ���</option>
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
	<td colspan="4">�ٹ����� ī�װ�</td>
	<td colspan="4">11���� ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>11���� (�ѱ�)</td>
</tr>
<% If oMy11st.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oMy11st.FresultCount - 1
			Depth1Nm = oMy11st.FItemList(i).FDepth1Nm
			Depth2Nm = oMy11st.FItemList(i).FDepth2Nm
			Depth3Nm = oMy11st.FItemList(i).FDepth3Nm

			If Depth3Nm = "" Then
				cateAllNm 	= Depth1Nm &" > "& Depth2Nm
				matchCateNm = Depth2Nm
			Else
				cateAllNm 	= Depth1Nm &" > "& Depth2Nm & " > " & Depth3Nm 
				matchCateNm = Depth3Nm
			End If
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oMy11st.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oMy11st.FItemList(i).FtenCateLarge & oMy11st.FItemList(i).FtenCateMid & oMy11st.FItemList(i).FtenCateSmall %></td>
	<td><%= oMy11st.FItemList(i).FtenCDLName %></td>
	<td><%= oMy11st.FItemList(i).FtenCDMName %></td>
	<td><%= oMy11st.FItemList(i).FtenCDSName %></td>
	<% If oMy11st.FItemList(i).FCateKey="" OR isNull(oMy11st.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="11���� ī�� ��Ī" onClick="popCjCateMap('<%= oMy11st.FItemList(i).FtenCateLarge %>','<%= oMy11st.FItemList(i).FtenCateMid %>','<%= oMy11st.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= oMy11st.FItemList(i).FtenCateLarge %>','<%= oMy11st.FItemList(i).FtenCateMid %>','<%= oMy11st.FItemList(i).FtenCateSmall %>','<%=oMy11st.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oMy11st.FItemList(i).FCateKey %></td>
	<td title="<%=cateAllNm%>" onClick="popCjCateMap('<%= oMy11st.FItemList(i).FtenCateLarge %>','<%= oMy11st.FItemList(i).FtenCateMid %>','<%= oMy11st.FItemList(i).FtenCateSmall %>','<%=oMy11st.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= matchCateNm %></td>
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
		<% If oMy11st.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oMy11st.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oMy11st.StartScrollPage to oMy11st.FScrollCount + oMy11st.StartScrollPage - 1 %>
			<% If i > oMy11st.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oMy11st.HasNextScroll Then %>
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
<% Set oMy11st = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->