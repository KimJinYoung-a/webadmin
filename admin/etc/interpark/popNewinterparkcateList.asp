<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/interpark/interparkcls.asp"-->
<%
Dim oInterpark, i, page, isMapping, srcDiv, srcKwd, orderby
Dim cateAllNm
Dim Depth1Name, Depth2Name, Depth3Name, Depth4Name

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"
If orderby = ""	Then orderby = "1"

'// ��� ����
Set oInterpark = new CInterpark
	oInterpark.FPageSize 			= 20
	oInterpark.FCurrPage			= page
	oInterpark.FRectIsMapping		= isMapping
	oInterpark.FRectSDiv			= srcDiv
	oInterpark.FRectKeyword		= srcKwd
	oInterpark.FRectCDL			= request("cdl")
	oInterpark.FRectCDM			= request("cdm")
	oInterpark.FRectCDS			= request("cds")
	oInterpark.FRectOrderby		= orderby
	oInterpark.getTenInterparkCateList
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

	// ������ũ ī�װ� ��Ī �˾�
	function popInterparkCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popinterparkCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr valign="top" bgcolor="#FFFFFF">
	<td>
		<font color="red"><strong>Interpark ī�װ� ����</strong></font>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>Interpark �ڵ�</option>
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
	<td colspan="4">Interpark ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>Interpark (�ѱ�)</td>
</tr>
<% If oInterpark.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oInterpark.FresultCount - 1
			cateAllNm 	= oInterpark.FItemList(i).FDispNm
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(oInterpark.FItemList(i).FCateKey),"#CCCCCC","#FFFFFF") %>">
	<td><%= oInterpark.FItemList(i).FtenCateLarge & oInterpark.FItemList(i).FtenCateMid & oInterpark.FItemList(i).FtenCateSmall %></td>
	<td><%= oInterpark.FItemList(i).FtenCDLName %></td>
	<td><%= oInterpark.FItemList(i).FtenCDMName %></td>
	<td><%= oInterpark.FItemList(i).FtenCDSName %></td>
	<td><%= oInterpark.FItemList(i).FItemcnt %></td>
	<% If oInterpark.FItemList(i).FCateKey="" OR isNull(oInterpark.FItemList(i).FCateKey) Then %>
	<td colspan="3"><input type="button" class="button" value="Interpark ī�� ��Ī" onClick="popInterparkCateMap('<%= oInterpark.FItemList(i).FtenCateLarge %>','<%= oInterpark.FItemList(i).FtenCateMid %>','<%= oInterpark.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popInterparkCateMap('<%= oInterpark.FItemList(i).FtenCateLarge %>','<%= oInterpark.FItemList(i).FtenCateMid %>','<%= oInterpark.FItemList(i).FtenCateSmall %>','<%=oInterpark.FItemList(i).FCateKey%>')" style="cursor:pointer"><%= oInterpark.FItemList(i).FCateKey %></td>
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
		<% If oInterpark.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oInterpark.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oInterpark.StartScrollPage to oInterpark.FScrollCount + oInterpark.StartScrollPage - 1 %>
			<% If i > oInterpark.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oInterpark.HasNextScroll Then %>
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
<% Set oInterpark = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->