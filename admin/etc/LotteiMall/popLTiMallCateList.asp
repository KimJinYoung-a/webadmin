<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<%
	dim oiMall, i, page, isMapping, srcDiv, srcKwd
    dim disptpcd
    
	page		= request("page")
	isMapping	= request("ismap")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
    disptpcd    = request("disptpcd")
    
	if page="" then page=1
	if srcDiv="" then srcDiv="LCD"

	'// ��� ����
	Set oiMall = new cLotteIMall
	oiMall.FPageSize = 20
	oiMall.FCurrPage = page
	oiMall.FRectIsMapping = isMapping
	oiMall.FRectSDiv = srcDiv
	oiMall.FRectKeyword = srcKwd
	oiMall.FRectCDL = request("cdl")
	oiMall.FRectCDM = request("cdm")
	oiMall.FRectCDS = request("cds")
	oiMall.FRectdisptpcd = disptpcd
	oiMall.getTenLTiMallCateList

dim cateAllNm
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

	// �Ե�iMall ī�װ� ��Ī �˾�
	function popLotteCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popLTiMallCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>�Ե�iMall ī�װ� ����</strong></font></td>
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
		
		���ñ��� : 
	    <select name="disptpcd" class="select">
	        <option value="">����</option>
			<option value="B" <%=chkIIF(disptpcd="B","selected","")%>>����</option>
			<option value="D" <%=chkIIF(disptpcd="D","selected","")%>>�Ϲ�</option>
		</select> /
		
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>�̸�Ī</option>
		</select> /
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>�Ե�iMall �ڵ�</option>
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
	<td colspan="4">�ٹ����� ī�װ�</td>
	<!-- td colspan="2">�Ե�iMall ��ǰ�з�</td -->
	<td colspan="4">�Ե�iMall ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<!--td>�ڵ�</td-->
	<td>����</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>iMall ����(�ѱ�)</td>
</tr>
<% if oiMall.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	else
		for i=0 to oiMall.FresultCount-1
		cateAllNm = oiMall.FItemList(i).FDispLrgNm & ">" & oiMall.FItemList(i).FDispMidNm & ">" & oiMall.FItemList(i).FDispSmlNm & ">" & oiMall.FItemList(i).FDispThnNm

%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(oiMall.FItemList(i).FCateIsUsing="Y","#FFFFFF","#CCCCCC") %>">
	<td><%= oiMall.FItemList(i).FtenCateLarge & oiMall.FItemList(i).FtenCateMid & oiMall.FItemList(i).FtenCateSmall %></td>
	<td><%= oiMall.FItemList(i).FtenCDLName %></td>
	<td><%= oiMall.FItemList(i).FtenCDMName %></td>
	<td><%= oiMall.FItemList(i).FtenCDSName %></td>
	<% if oiMall.FItemList(i).FDispNo="" or isNull(oiMall.FItemList(i).FDispNo) then %>
	<td colspan="3"><input type="button" class="button" value="�Ե�iMall ī�� ��Ī" onClick="popLotteCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','')"></td>
	<% else %>
	<td title="<%=cateAllNm%>" onClick="popLotteCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).getDispGubunNm %></td>
	<td title="<%=cateAllNm%>" onClick="popLotteCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).FDispNo %></td>
	<td title="<%=cateAllNm%>" onClick="popLotteCateMap('<%= oiMall.FItemList(i).FtenCateLarge %>','<%= oiMall.FItemList(i).FtenCateMid %>','<%= oiMall.FItemList(i).FtenCateSmall %>','<%=oiMall.FItemList(i).FDispNo%>')" style="cursor:pointer"><%= oiMall.FItemList(i).FDispNm %></td>
	<td><%=cateAllNm%>
	<% end if %>
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
		<% if oiMall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
			<% if i>oiMall.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oiMall.HasNextScroll then %>
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
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
