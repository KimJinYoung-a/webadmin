<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim oiMall, i, page, isMapping, srcDiv, srcKwd, orderby
Dim disptpcd
page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
disptpcd    = request("disptpcd")
orderby		= request("orderby")

If page = "" Then page = 1
If srcDiv = "" Then srcDiv="LCD"

'// ��� ����
Set oiMall = new CLotteiMall
	oiMall.FPageSize = 20
	oiMall.FCurrPage = page
	oiMall.FRectIsMapping = isMapping
	oiMall.FRectSDiv = srcDiv
	oiMall.FRectKeyword = srcKwd
	oiMall.FRectCDL = request("cdl")
	oiMall.FRectCDM = request("cdm")
	oiMall.FRectCDS = request("cds")
	oiMall.FRectOrderby		= orderby
	oiMall.FRectdisptpcd = disptpcd
	oiMall.getTenLotteimallCateList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	// �˻�
	function serchItem() {
		frm.page.value=1;
		frm.submit();
	}

	// �Ե����̸� ī�װ� ��Ī �˾�
	function popLotteiMallCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popLtiMallCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=500,height=300,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>�Ե����̸� ī�װ� ����</strong></font></td>
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

		���屸�� : 
	    <select name="disptpcd" class="select">
	        <option value="">����</option>
			<option value="10" <%=chkIIF(disptpcd = "10", "selected","")%>>�Ϲݸ���</option>
			<option value="12" <%=chkIIF(disptpcd = "12", "selected","")%>>��������</option>
			<option value="99" <%=chkIIF(disptpcd = "99", "selected","")%>>�ű�ī�װ�</option>
		</select> /
		
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping = "Y", "selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping = "N", "selected","")%>>�̸�Ī</option>
		</select> /
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv = "LCD", "selected","")%>>�Ե����̸� �ڵ�</option>
			<option value="CNM" <%=chkIIF(srcDiv = "CNM", "selected","")%>>ī�װ���</option>
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
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oiMall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">�ٹ����� ī�װ�</td>
	<td colspan="5">�Ե����̸� ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>����</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>LotteiMall ����(�ѱ�)</td>
</tr>
<% If oiMall.FResultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oiMall.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(oiMall.FItemList(i).FCateisusing="N","#CCCCCC","#FFFFFF") %>">
	<td><%= oiMall.FItemList(i).FtenCateLarge & oiMall.FItemList(i).FtenCateMid & oiMall.FItemList(i).FtenCateSmall %></td>
	<td><%= oiMall.FItemList(i).FtenCDLName %></td>
	<td><%= oiMall.FItemList(i).FtenCDMName %></td>
	<td><%= oiMall.FItemList(i).FtenCDSName %></td>
	<td><%= oiMall.FItemList(i).FItemcnt %></td>
	<% If oiMall.FItemList(i).FDispNo="" or isNull(oiMall.FItemList(i).FDispNo) then %>
	<td colspan="5"><input type="button" class="button" value="�Ե����̸� ��Ī" onClick="popLotteiMallCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%= oiMall.FItemList(i).FDispLrgNm&">"&oiMall.FItemList(i).FDispMidNm&">"&oiMall.FItemList(i).FDispSmlNm %>" onClick="popLotteiMallCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).getDisptpcdName %></td>
	<td title="<%= oiMall.FItemList(i).FDispLrgNm&">"&oiMall.FItemList(i).FDispMidNm&">"&oiMall.FItemList(i).FDispSmlNm %>" onClick="popLotteiMallCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).FgroupCode %></td>
	<td title="<%= oiMall.FItemList(i).FDispLrgNm&">"&oiMall.FItemList(i).FDispMidNm&">"&oiMall.FItemList(i).FDispSmlNm %>" onClick="popLotteiMallCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).FDispNo %></td>
	<td title="<%= oiMall.FItemList(i).FDispLrgNm&">"&oiMall.FItemList(i).FDispMidNm&">"&oiMall.FItemList(i).FDispSmlNm %>" onClick="popLotteiMallCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).FDispNm %></td>
	<td><%= oiMall.FItemList(i).FDispLrgNm&">"&oiMall.FItemList(i).FDispMidNm&">"&oiMall.FItemList(i).FDispSmlNm %></td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oiMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage - 1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
			<% If i>oiMall.FTotalpage Then Exit For %>
			<% If CStr(page) = CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>

		<% If oiMall.HasNextScroll Then %>
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
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
