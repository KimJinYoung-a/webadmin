<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü����Ʈ
' History : 2009.04.07 ������ ����
'			2012.09.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim ogroup, frmname, page, rectconame, rectDesigner, rectsocno, groupid , ceoname ,i, isusing, vTmpGr, vGrArr, vItemTotalCount
	frmname     = request("frmname")
	page        = requestCheckVar(request("page"),9)
	rectconame  = requestCheckVar(request("rectconame"),32)
	rectDesigner = requestCheckVar(request("rectDesigner"),32)
	rectsocno   = requestCheckVar(request("rectsocno"),16)
	groupid   = requestCheckVar(request("groupid"),16)
	ceoname     = request("ceoname")
	isusing = requestCheckVar(request("isusing"),1)

if page="" then page=1

'### ��ü ��ǰ�� ��Ÿ���� ���. �� �ʿ��ϴ�ϴ�;	'/2017.04.26 ���ر� �߰�(�������̻�� ����)	'/2017.04.26 �ѿ�� �ּ�ó��(�������̻���� �ٽ� ���޶�� �Ͻ�)
'vItemTotalCount = fnITemTotalCount()

set ogroup = new CPartnerGroup
	ogroup.FPageSize = 30
	ogroup.FCurrPage = page
	ogroup.FrectDesigner = rectDesigner
	ogroup.Frectconame = rectconame
	ogroup.FRectsocno = rectsocno
	ogroup.FRectGroupid = groupid
	ogroup.FRectceoname = ceoname
	ogroup.FRectIsusing = isusing
	
	if rectDesigner<>"" then
		ogroup.GetGroupInfoListByBrand
	else
		ogroup.GetGroupInfoList
	end if

	vTmpGr = ogroup.FGroupIdList
	If vTmpGr <> "" Then
		vTmpGr = Left(vTmpGr, Len(vTmpGr)-1)
		ogroup.FGroupIdList = vTmpGr
		vGrArr = ogroup.fnGroupInfoByItemCount
	End If
%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ȸ��� : <input type="text" name="rectconame" class="text" value="<%= rectconame %>" size=10 maxlength=32>
		&nbsp;&nbsp;
		�׷��ڵ� : <input type="text" name="groupid" class="text" value="<%= groupid %>" size=8 maxlength=6>
    	&nbsp;&nbsp;
    	����ڹ�ȣ : <input type="text" name="rectsocno" class="text" value="<%= rectsocno %>" size=15 maxlength=12>
    	&nbsp;&nbsp;
    	���Ժ귣�� : <input type="text" name="rectDesigner" class="text" value="<%= rectDesigner %>" Maxlength="32" size="16">
    	��ǥ�ڸ� : <input type="text" name="ceoname" class="text" value="<%= ceoname %>" Maxlength="8" size="8">
    	&nbsp;&nbsp;
    	��ü�˻� :
    	<select name="isusing">
    		<option value="" <%=CHKIIF(isusing="","selected","")%>>��ü</option>
    		<option value="Y" <%=CHKIIF(isusing="Y","selected","")%>>�����</option>
    		<option value="N" <%=CHKIIF(isusing="N","selected","")%>>����</option>
    	</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ogroup.FtotalCount %> ��</b>
		&nbsp;
		������ : <b><%= page %> / <%= ogroup.FTotalpage %></b>
		<!--&nbsp;
		��ǰ �� �� : <b><%'=FormatNumber(vItemTotalCount,0)%></b> (�Ǹſ��λ������ ������ΰ�)-->
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width="60" >��ü�ڵ�<br>(�׷��ڵ�)</td>
	<td width="130" >ȸ���</td>
	<td width="80" >����ڹ�ȣ</td>
	<td width="60" >��ǥ��</td>
	<td width="130" >��ȭ��ȣ<br>�ѽ���ȣ</td>
	<td width="80" >�����</td>
	<% if (FALSE) then %>
	<td>�ڵ�����ȣ<br>�̸����ּ�</td>
    <% end if %>
	<td>����귣��</td>
	<td>��ǰ��</td>
	<!-- <td >�귣���뿩��</td> -->
</tr>
<% if ogroup.FResultCount >0 then %>
<% for i=0 to ogroup.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ogroup.FItemList(i).FGroupID %></td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= ogroup.FItemList(i).FGroupID %>')"><%= ogroup.FItemList(i).Fcompany_name %></a></td>
	<td align="center"><%= socialnoReplace(ogroup.FItemList(i).Fcompany_no) %></td>
	<td align="center"><%= ogroup.FItemList(i).Fceoname %></td>
	<td>TEL : <%= ogroup.FItemList(i).Fcompany_tel %><br>FAX : <%= ogroup.FItemList(i).Fcompany_fax %></td>
	<td align="center"><%= ogroup.FItemList(i).Fmanager_name %></td>
	<% if (FALSE) then %>
	<td>H.P : <%= ogroup.FItemList(i).Fmanager_phone %><br>E-mail : <%= ogroup.FItemList(i).Fmanager_email %></td>
	<% end if %>
	<td <%=ChkIIF(ogroup.FItemList(i).getPartnerIdInfoStr="","bgcolor='#CCCCCC'","")%> ><%= ogroup.FItemList(i).getPartnerIdInfoStr %></td>
	<td align="center"><%=fnGroupListItemCntView(vGrArr,ogroup.FItemList(i).FGroupID)%> ��</td>
	<!-- <td></td> -->
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<% if ogroup.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ogroup.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ogroup.StartScrollPage to ogroup.FScrollCount + ogroup.StartScrollPage - 1 %>
		<% if i>ogroup.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ogroup.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=10 align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->



