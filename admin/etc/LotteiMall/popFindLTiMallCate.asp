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
	dim oiMall, i, page, isMapping, srcDiv, srcKwd, useyn, research
    dim disptpcd
    
	page		= request("page")
	isMapping	= request("ismap")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
	disptpcd    = request("disptpcd")
	useyn       = request("useyn")
	research    = request("research")
	
	if page="" then page=1
	if srcDiv="" then srcDiv="CNM"

    if (research="") and useyn="" then useyn="Y"
    if (research="") and disptpcd="" then disptpcd="B"
    
	'// ��� ����
	Set oiMall = new cLotteiMall
	oiMall.FPageSize = 20
	oiMall.FCurrPage = page
	oiMall.FRectIsMapping = isMapping
	oiMall.FRectSDiv = srcDiv
	oiMall.FRectKeyword = srcKwd
	oiMall.FRectdisptpcd = disptpcd
	oiMall.FRectCateUsingYn = useyn
	oiMall.getLTiMallCategoryList

%>
<script language="javascript">
<!--
	// �Ե�iMall ���� ī�װ� ����
	function refreshLotteiMallCate() {
		if(confirm("���� ī�װ��� �Ե�iMall �������� �����޾� �����Ͻðڽ��ϱ�?\n\n�� 1.��Ż��¿����� �ټ� �ð��� ���� �ɸ� �� �ֽ��ϴ�.\n�� 2.������ �Ǿ��ִ� ī�װ��� �����ϰ� �ٽ� �Ե�iMall�� ������ �������� ���̹Ƿ� �����ϰ� �����ϼ���.")) {
			document.getElementById("btnRefresh").disabled=true;
			xLink.location.href="actLotteiMallReq.asp?cmdparam=getdispcate";
		}
	}

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

	// ī�װ� ����
	function fnSelCate(disptpcd,dspNo,dspNm) {
	    opener.document.frmAct.dspNo.value=dspNo;
		//opener.document.getElementById("brTT").rowSpan=2;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML="[" + dspNo + "] " + dspNm;
		self.close();
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
	<font color="red"><strong>�Ե�iMall ī�װ� �˻�</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
<form name="frm" method="GET" style="margin:0px;" onSubmit="serchItem();">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
	    ��뿩�� :
	    
	    <select name="useyn" class="select">
	        <option value="">����</option>
			<option value="Y" <%=chkIIF(useyn="Y","selected","")%>>���</option>
			<option value="N" <%=chkIIF(useyn="N","selected","")%>>������</option>
		</select> /
		
	    �з�/���ñ��� : 
	    <select name="disptpcd" class="select">
	        <option value="">����</option>
			<option value="B" <%=chkIIF(disptpcd="B","selected","")%>>����</option>
			<option value="D" <%=chkIIF(disptpcd="D","selected","")%>>�Ϲ�</option>
		</select> /
		
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>�Ե�iMall �ڵ�</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>ī�װ���</option>
		</select> /
		�˻��� :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text"> &nbsp;
		<input id="btnRefresh" type="button" class="button" value="�˻�" onclick="serchItem()">
	</td>
</tr>
<tr>
	<td align="left" style="padding-top:5px;">
	    <input id="btnRefresh" type="button" class="button" value="�Ե�iMall ����ī�װ� ����" onclick="refreshLotteiMallCate()">
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
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td>����</td>
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>���з�</td>
	<td>ī�װ���</td>
</tr>
<% if oiMall.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	else
		for i=0 to oiMall.FresultCount-1
%>
<tr align="center" height="25" onClick="fnSelCate('<%= oiMall.FItemList(i).Fdisptpcd %>','<%= oiMall.FItemList(i).FDispNo %>','<%=replace(oiMall.FItemList(i).FDispNm,"""","")%>')" style="cursor:pointer" title="ī�װ� ����" bgcolor="<%=chkIIF(oiMall.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oiMall.FItemList(i).getDispGubunNm %></td>
	<td><%= oiMall.FItemList(i).FDispNo %></td>
	<td><%= oiMall.FItemList(i).FDispLrgNm %></td>
	<td><%= oiMall.FItemList(i).FDispMidNm %></td>
	<td><%= oiMall.FItemList(i).FDispSmlNm %></td>
	<td><%= oiMall.FItemList(i).FDispThnNm %></td>
	<td><%= oiMall.FItemList(i).FDispNm %></td>
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
<iframe name="xLink" id="xLink" frameborder="1" width="610" height="100"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
