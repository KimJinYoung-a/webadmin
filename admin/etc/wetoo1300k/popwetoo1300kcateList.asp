<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/wetoo1300k/wetoo1300kcls.asp"-->
<%
Dim cateAllNm, cateAllkey, Depth1Key, Depth2Key, Depth3Key, Depth4Key

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"
If orderby = ""	Then orderby = "1"

Dim o1300k, i, page, isMapping, srcDiv, srcKwd, orderby
'// ��� ����
Set o1300k = new C1300k
	o1300k.FPageSize 		= 50
	o1300k.FCurrPage		= page
	o1300k.FRectIsMapping	= isMapping
	o1300k.FRectSDiv		= srcDiv
	o1300k.FRectKeyword		= srcKwd
	o1300k.FRectCDL			= request("cdl")
	o1300k.FRectCDM			= request("cdm")
	o1300k.FRectCDS			= request("cds")
	o1300k.FRectOrderby		= orderby
	o1300k.getTen1300kCateList
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

	// wetoo1300k ī�װ� ��Ī �˾�
	function popwetoo1300kCateMap(cdl, cdm, cds, l, m, s, d) {
		var pCM = window.open("popwetoo1300kCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&large_category="+l+"&millde_category="+m+"&small_category="+s+"&detail_category="+d,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>wetoo1300k ī�װ� ����</strong></font></td>
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
	<td colspan="4">wetoo1300k ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>wetoo1300k (�ѱ�)</td>
</tr>
<% If o1300k.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to o1300k.FresultCount - 1
			Depth1Key = o1300k.FItemList(i).FLarge_category
			Depth2Key = o1300k.FItemList(i).FMiddle_category
			Depth3Key = o1300k.FItemList(i).FSmall_category
			Depth4Key = o1300k.FItemList(i).FDetail_category
			cateAllNm = o1300k.FItemList(i).FCategory_name
			cateAllkey = Depth1Key & Depth2Key & Depth3Key & Depth4Key
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(cateAllkey),"#CCCCCC","#FFFFFF") %>">
	<td><%= o1300k.FItemList(i).FtenCateLarge & o1300k.FItemList(i).FtenCateMid & o1300k.FItemList(i).FtenCateSmall %></td>
	<td><%= o1300k.FItemList(i).FtenCDLName %></td>
	<td><%= o1300k.FItemList(i).FtenCDMName %></td>
	<td><%= o1300k.FItemList(i).FtenCDSName %></td>
	<td><%= o1300k.FItemList(i).FItemcnt %></td>
	<% If cateAllkey="" OR isNull(cateAllkey) Then %>
	<td colspan="3"><input type="button" class="button" value="wetoo1300k ī�� ��Ī" onClick="popwetoo1300kCateMap('<%= o1300k.FItemList(i).FtenCateLarge %>','<%= o1300k.FItemList(i).FtenCateMid %>','<%= o1300k.FItemList(i).FtenCateSmall %>','', '', '', '')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popwetoo1300kCateMap('<%= o1300k.FItemList(i).FtenCateLarge %>','<%= o1300k.FItemList(i).FtenCateMid %>','<%= o1300k.FItemList(i).FtenCateSmall %>','<%= o1300k.FItemList(i).FLarge_category %>','<%= o1300k.FItemList(i).FMiddle_category %>','<%= o1300k.FItemList(i).FSmall_category %>','<%= o1300k.FItemList(i).FDetail_category %>')" style="cursor:pointer"><%= cateAllkey %></td>
	<td title="<%=cateAllNm%>" onClick="popwetoo1300kCateMap('<%= o1300k.FItemList(i).FtenCateLarge %>','<%= o1300k.FItemList(i).FtenCateMid %>','<%= o1300k.FItemList(i).FtenCateSmall %>','<%= o1300k.FItemList(i).FLarge_category %>','<%= o1300k.FItemList(i).FMiddle_category %>','<%= o1300k.FItemList(i).FSmall_category %>','<%= o1300k.FItemList(i).FDetail_category %>')" style="cursor:pointer"><%= cateAllNm %></td>
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
		<% If o1300k.HasPreScroll Then %>
		<a href="javascript:goPage('<%= o1300k.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + o1300k.StartScrollPage to o1300k.FScrollCount + o1300k.StartScrollPage - 1 %>
			<% If i > o1300k.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If o1300k.HasNextScroll Then %>
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
<% Set o1300k = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->