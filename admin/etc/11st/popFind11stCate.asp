<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/11st/11stcls.asp"-->
<%
Dim o11st, i, page, srcKwd, isNull4DpethNm
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// ��� ����
Set o11st = new C11st
	o11st.FPageSize = 5000
	o11st.FCurrPage = page
	o11st.FsearchName = srcKwd
	o11st.get11stCateList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// ��ǰ�з� ����
	function fnSelDispCate(dpCode, dp6nm) {
		opener.document.frmAct.depthcode.value=dpCode;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dp6nm;
		self.close();
	}

	// ���õ� ī�װ� ����
	function st11CateInfo() {
		var chkSel=0;
		try {
			if(frmSvArr.cksel.length>1) {
				for(var i=0;i<frmSvArr.cksel.length;i++) {
					if(frmSvArr.cksel[i].checked) chkSel++;
				}
			} else {
				if(frmSvArr.cksel.checked) chkSel++;
			}
			if(chkSel<=0) {
				alert("������ ī�װ��� �����ϴ�.");
				return;
			}
		}
		catch(e) {
			alert("ī�װ��� �����ϴ�.");
			return;
		}

	    if (confirm('11������ �����Ͻ� ' + chkSel + '�� ī�װ� ������ ȣ�� �Ͻðڽ��ϱ�?')){
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "GETCATE";
	        document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
	        document.frmSvArr.submit();
	    }
	}

//-->
</script>
<input class="button" type="button" id="btnOPTSel" value="ī������" onClick="st11CateInfo();">&nbsp;&nbsp;
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>11st ī�װ� �˻�</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=o11st.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>DepthCode</td>
	<td>Depth1Name</td>
	<td>Depth2Name</td>
	<td>Depth3Name</td>
	<td>Depth4Name</td>
</tr>
<% If o11st.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to o11st.FresultCount - 1
			If Trim(o11st.FItemList(i).Fdepth4Nm) = "" Then
				isNull4DpethNm = o11st.FItemList(i).Fdepth3Nm
			Else
				isNull4DpethNm = o11st.FItemList(i).Fdepth4Nm
			End If
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= o11st.FItemList(i).FdepthCode %>', '<%= replace(isNull4DpethNm, "'", "`") %>')" style="cursor:pointer" title="ī�װ� ����" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= o11st.FItemList(i).FdepthCode %>"></td>
	<td><%= o11st.FItemList(i).FdepthCode %></td>
	<td><%= o11st.FItemList(i).Fdepth1Nm %></td>
	<td><%= o11st.FItemList(i).Fdepth2Nm %></td>
	<td><%= o11st.FItemList(i).Fdepth3Nm %></td>
	<td><%= o11st.FItemList(i).Fdepth4Nm %></td>
</tr>
<%
		Next
	End If
%>
</table>
</form>
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
			<% If i>o11st.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

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
<!-- ǥ �ϴܹ� ��-->
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set o11st = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
