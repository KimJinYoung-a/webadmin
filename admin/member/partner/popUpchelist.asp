<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%

dim ogroup, frmname, page, rectconame,rectDesigner, rectsocno
dim compname, vGubun

frmname         = request("frmname")
page            = requestCheckVar(request("page"),9)
rectconame      = requestCheckVar(request("rectconame"),32)
rectDesigner    = requestCheckVar(request("rectDesigner"),32)
rectsocno       = requestCheckVar(request("rectsocno"),16)
vGubun			= request("gb")

'' �߰�
compname        = request("compname")

if page="" then page=1

set ogroup = new CPartnerGroup
ogroup.FPageSize = 15
ogroup.FCurrPage = page
ogroup.FrectDesigner = rectDesigner
ogroup.Frectconame = rectconame
ogroup.FRectsocno = rectsocno

if (rectDesigner<>"") then
	ogroup.GetGroupInfoListByBrand
else
	ogroup.GetGroupInfoList
end if


dim i
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();
    
}

<% if (compname<>"") then %>
function SelectThis(frmbuf){
	var openformcomp = eval('opener.<%= frmname %>.<%= compname %>');
	openformcomp.value = frmbuf.groupid.value;
	window.close();
}

<% else %>
function SelectThis(gcode){
	<% If vGubun = "search" Then %>
		opener.document.frm.reqgcode.value = gcode;
		window.close();
	<% Else %>
		document.location.href = "upcheinfo_edit_parent.asp?groupid="+gcode+"";
	<% End If %>
}
<% end if %>
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="frmname" value="<%= frmname %>">
	<input type="hidden" name="compname" value="<%= compname %>">

	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	ȸ��� : <input type=text name=rectconame value="<%= rectconame %>" size=10 maxlength=32>&nbsp;&nbsp;
		        ����ڹ�ȣ : <input type=text name=rectsocno value="<%= rectsocno %>" size=12 maxlength=12>(- ����)&nbsp;&nbsp;
		        ���Ժ귣�� : <input type="text" name="rectDesigner" value="<%= rectDesigner %>" Maxlength="32" size="16">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        <input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border=0 cellspacing=1 cellpadding=2  class=a bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#DDDDFF" align="center">
	<td width=100>��ü�ڵ�</td>
	<td width=260>��ü��</td>
	<td width=140>����ڹ�ȣ</td>
	<td>����귣��</td>
	<td width=70>����</td>
</tr>
<% if ogroup.FResultCount>0 then %>
<% for i=0 to ogroup.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ogroup.FItemList(i).FGroupID %></td>
	<td><%= ogroup.FItemList(i).FCompany_Name %></td>
	<td><%= socialnoReplace(ogroup.FItemList(i).FCompany_No) %></td>
	<td <%=ChkIIF(ogroup.FItemList(i).getPartnerIdInfoStr="","bgcolor='#CCCCCC'","")%> ><%= ogroup.FItemList(i).getPartnerIdInfoStr %></td>
	<td width=70 align="center"><input type= button value="����" onClick="SelectThis('<%= ogroup.FItemList(i).FGroupID %>')"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height=30>
	<td colspan=10 align=center>
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
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->