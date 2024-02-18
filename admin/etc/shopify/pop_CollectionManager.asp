<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shopify/shopifycls.asp"-->
<%
Dim oshopify, i, page, isMapping, srcDiv, srcKwd
Dim cateAllNm, matchCateNm
Dim Depth1Name, Depth2Name, Depth3Name
Dim ColType
page		= requestCheckvar(request("page"),10)
srcDiv		= requestCheckvar(request("srcDiv"),10)
srcKwd		= requestCheckvar(request("srcKwd"),60)
ColType     = requestCheckvar(request("ColType"),10)
If page = ""	Then page = 1

'// ��� ����
Set oshopify = new Cshopify
	oshopify.FPageSize 	= 20
	oshopify.FCurrPage	= page
	oshopify.FRectIsMapping	= isMapping
	oshopify.FRectColType   = ColType
	'oshopify.FRectSDiv		= srcDiv
	oshopify.FRectKeyword	= srcKwd
	
	oshopify.getShopifyCollectionList
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

	function fnReceiveCollection(){
	    if (confirm('collection ����� ���� �Ͻðڽ��ϱ�?')){
	        document.frmAct.target = "xLink";
	        document.frmAct.act.value="RCVCOLLECTIONS";
    		document.frmAct.action = "<%=apiURL%>/outmall/shopify/shopifyActProc.asp"
    		document.frmAct.submit();
	    }   
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr valign="top">
	<td><strong>* shopify Collection ����</strong></td>
	<td align="right"><input type="button" class="button" value="Collection��ϼ���" onclick="fnReceiveCollection()"></td>
</tr>

</table>
<!-- �׼� -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
	    Collection�з� :
	    <input type="radio" name="ColType" value="" <%=CHKIIF(ColType="","checked","")%> >��ü
	    <input type="radio" name="ColType" value="S" <%=CHKIIF(ColType="S","checked","")%> >Smart
	    <input type="radio" name="ColType" value="C" <%=CHKIIF(ColType="C","checked","")%> >Custome
	    &nbsp;|&nbsp;
	    <% if (FALSE) then %>
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>�̸�Ī</option>
		</select> /
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="CCD" <%=chkIIF(srcDiv="CCD","selected","")%>>shopify �ڵ�</option>
			<option value="CNM" <%=chkIIF(srcDiv="CNM","selected","")%>>10x10��ī�װ���</option>
		</select> /
	<% end if %>
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

<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>CollectionID</td>
	<td>Title</td>
	<td>Collection�з�</td>
	<td>���α���</td>
	<td>rules</td>
	<td>�����</td>
	<td>������</td>
	<td>��ǰ��</td>
	<td></td>
</tr>
<% If oshopify.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oshopify.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td><%= oshopify.FItemList(i).Fcollectionid %></td>
	<td><%= oshopify.FItemList(i).FTitle %></td>
	<td><%= oshopify.FItemList(i).getCollectionTypeName %></td>
	<td><%= oshopify.FItemList(i).getCollectionTypeSubName %></td>
	<td><%= oshopify.FItemList(i).getCollectionRuleStr %></td>
	<td><%= oshopify.FItemList(i).Fpublished_at %></td>
	<td><%= oshopify.FItemList(i).Fupdated_at %></td>
	<td><%= oshopify.FItemList(i).FCollectItemCount %></td>
	<td>&nbsp;</td>
</tr>
<%
		Next
	End If
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td valign="bottom" align="center">
		<% If oshopify.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oshopify.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oshopify.StartScrollPage to oshopify.FScrollCount + oshopify.StartScrollPage - 1 %>
			<% If i > oshopify.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oshopify.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
</tr>

</table>
<% Set oshopify = Nothing %>
<form name="frmAct" method="post" >
<input type="hidden" name="act">
</form>
<iframe name="xLink" id="xLink" frameborder="1" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->