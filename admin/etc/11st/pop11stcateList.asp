<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/11st/11stcls.asp"-->
<%
Dim o11st, i, page, isMapping, srcDiv, srcKwd, orderby
Dim cateAllNm
Dim Depth1Nm, Depth2Nm, Depth3Nm, Depth4Nm

page		= request("page")
isMapping	= request("ismap")
srcDiv		= request("srcDiv")
srcKwd		= request("srcKwd")
orderby		= request("orderby")

If page = ""	Then page = 1
If srcDiv = ""	Then srcDiv = "CCD"
If orderby = ""	Then orderby = "1"

'// ��� ����
Set o11st = new C11st
	o11st.FPageSize 		= 50
	o11st.FCurrPage			= page
	o11st.FRectIsMapping	= isMapping
	o11st.FRectSDiv			= srcDiv
	o11st.FRectKeyword		= srcKwd
	o11st.FRectCDL			= request("cdl")
	o11st.FRectCDM			= request("cdm")
	o11st.FRectCDS			= request("cds")
	o11st.FRectOrderby		= orderby
	o11st.getTen11stCateList
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

	// 11st ī�װ� ��Ī �˾�
	function pop11stCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("pop11stCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
	<font color="red"><strong>11st ī�װ� ����</strong></font></td>
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
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>11st �ڵ�</option>
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
	<td colspan="4">11st ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>��ǰ��</td>
	<td>�ڵ�</td>
	<td>ī�װ���</td>
	<td>11st (�ѱ�)</td>
</tr>
<% If o11st.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to o11st.FresultCount - 1
			Depth1Nm = o11st.FItemList(i).FDepth1Nm
			Depth2Nm = o11st.FItemList(i).FDepth2Nm
			Depth3Nm = o11st.FItemList(i).FDepth3Nm
			Depth4Nm = o11st.FItemList(i).FDepth4Nm
			If Depth4Nm = "" Then
				cateAllNm 	= Depth1Nm &" > "& Depth2Nm & " > " & Depth3Nm
			Else
				cateAllNm 	= Depth1Nm &" > "& Depth2Nm & " > " & Depth3Nm & " > " & Depth4Nm
			End If
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(IsNull(o11st.FItemList(i).FDepthCode),"#CCCCCC","#FFFFFF") %>">
	<td><%= o11st.FItemList(i).FtenCateLarge & o11st.FItemList(i).FtenCateMid & o11st.FItemList(i).FtenCateSmall %></td>
	<td><%= o11st.FItemList(i).FtenCDLName %></td>
	<td><%= o11st.FItemList(i).FtenCDMName %></td>
	<td><%= o11st.FItemList(i).FtenCDSName %></td>
	<td><%= o11st.FItemList(i).FItemcnt %></td>
	<% If o11st.FItemList(i).FDepthCode="" OR isNull(o11st.FItemList(i).FDepthCode) Then %>
	<td colspan="3"><input type="button" class="button" value="11st ī�� ��Ī" onClick="pop11stCateMap('<%= o11st.FItemList(i).FtenCateLarge %>','<%= o11st.FItemList(i).FtenCateMid %>','<%= o11st.FItemList(i).FtenCateSmall %>','')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="pop11stCateMap('<%= o11st.FItemList(i).FtenCateLarge %>','<%= o11st.FItemList(i).FtenCateMid %>','<%= o11st.FItemList(i).FtenCateSmall %>','<%=o11st.FItemList(i).FDepthCode%>')" style="cursor:pointer"><%= o11st.FItemList(i).FDepthCode %></td>
	<td title="<%=cateAllNm%>" onClick="pop11stCateMap('<%= o11st.FItemList(i).FtenCateLarge %>','<%= o11st.FItemList(i).FtenCateMid %>','<%= o11st.FItemList(i).FtenCateSmall %>','<%=o11st.FItemList(i).FDepthCode%>')" style="cursor:pointer"><%= Chkiif(o11st.FItemList(i).FDepth4Nm="", o11st.FItemList(i).FDepth3Nm, o11st.FItemList(i).FDepth4Nm) %></td>
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
		<% If o11st.HasPreScroll Then %>
		<a href="javascript:goPage('<%= o11st.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + o11st.StartScrollPage to o11st.FScrollCount + o11st.StartScrollPage - 1 %>
			<% If i > o11st.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If o11st.HasNextScroll Then %>
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
<% Set o11st = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->