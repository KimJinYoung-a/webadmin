<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lfmall/lfmallCls.asp"-->
<%
Dim oLfmall, i, page, isMapping, srcDiv, srcKwd, orderby
Dim cateAllNm
Dim Depth1Name, Depth2Name, Depth3Name, Depth4Name

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"

'// ��� ����
Set oLfmall = new CLfmall
	oLfmall.FPageSize 			= 20
	oLfmall.FCurrPage			= page
	oLfmall.FRectIsMapping		= isMapping
	oLfmall.FRectSDiv			= srcDiv
	oLfmall.FRectKeyword		= srcKwd
	oLfmall.FRectCDL			= request("cdl")
	oLfmall.FRectCDM			= request("cdm")
	oLfmall.FRectCDS			= request("cds")
	oLfmall.FRectOrderby		= orderby
	oLfmall.getTenLfmallCateList
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

	// Lfmall ī�װ� ��Ī �˾�
	function popLfmallCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("poplfmallCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr valign="top" bgcolor="#FFFFFF">
	<td>
		<font color="red"><strong>Lfmall ī�װ� ����</strong></font>
	</td>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>Lfmall �ڵ�</option>
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

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">�ٹ����� ī�װ�</td>
	<td colspan="4">Lfmall ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>Lfmall (�ѱ�)</td>
</tr>
<% If oLfmall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oLfmall.FresultCount - 1
			Depth1Name = oLfmall.FItemList(i).FDepth1Name
			Depth2Name = oLfmall.FItemList(i).FDepth2Name
			Depth3Name = oLfmall.FItemList(i).FDepth3Name
			Depth4Name = oLfmall.FItemList(i).FDepth4Name
			If Depth4Name = "" Then
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name
			Else
				cateAllNm 	= Depth1Name &" > "& Depth2Name & " > " & Depth3Name & " > " & Depth4Name
			End If
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oLfmall.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oLfmall.FItemList(i).FtenCateLarge & oLfmall.FItemList(i).FtenCateMid & oLfmall.FItemList(i).FtenCateSmall %></td>
	<td><%= oLfmall.FItemList(i).FtenCDLName %></td>
	<td><%= oLfmall.FItemList(i).FtenCDMName %></td>
	<td><%= oLfmall.FItemList(i).FtenCDSName %></td>
	<td><%= oLfmall.FItemList(i).FItemcnt %></td>
	<% If oLfmall.FItemList(i).FCateKey="" OR isNull(oLfmall.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="Lfmall ī�� ��Ī" onClick="popLfmallCateMap('<%= oLfmall.FItemList(i).FtenCateLarge %>','<%= oLfmall.FItemList(i).FtenCateMid %>','<%= oLfmall.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popLfmallCateMap('<%= oLfmall.FItemList(i).FtenCateLarge %>','<%= oLfmall.FItemList(i).FtenCateMid %>','<%= oLfmall.FItemList(i).FtenCateSmall %>','<%=oLfmall.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oLfmall.FItemList(i).FCateKey %></td>
	<td title="<%=cateAllNm%>" onClick="popLfmallCateMap('<%= oLfmall.FItemList(i).FtenCateLarge %>','<%= oLfmall.FItemList(i).FtenCateMid %>','<%= oLfmall.FItemList(i).FtenCateSmall %>','<%=oLfmall.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= Chkiif(oLfmall.FItemList(i).FDepth4Name="", oLfmall.FItemList(i).FDepth3Name, oLfmall.FItemList(i).FDepth4Name) %></td>
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
		<% If oLfmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oLfmall.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oLfmall.StartScrollPage to oLfmall.FScrollCount + oLfmall.StartScrollPage - 1 %>
			<% If i > oLfmall.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oLfmall.HasNextScroll Then %>
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
<% Set oLfmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->