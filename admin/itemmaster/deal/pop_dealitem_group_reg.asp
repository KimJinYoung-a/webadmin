<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/pop_dealitem_group.asp
' Description :  �� ��ǰ �׷���
' History : 2022.10.17 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
Dim idx : idx = Request("idx")
Dim groupCode : groupCode = Request("groupCode")
Dim sTarget : sTarget = request("sTarget")
dim cdealGroup, arrList
dim mode
if groupCode<>"" then
    mode="update"
else
    mode="add"
end if
set cdealGroup = new CDealSelect
cdealGroup.FRectDealCode = idx
cdealGroup.FRectGroupCode = groupCode
cdealGroup.fnGetDealItemGroupDetail
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
window.document.domain = "10x10.co.kr";
function jsGroupSubmit(){
	if(!document.frmG.title.value){
		alert("�׷���� �Է����ּ���");
		document.frmG.title.focus();
		return false;
	}else{
		document.frmG.submit();
	}
}
function jsDelGroup(groupcode){
	document.frmGM.groupCode.value=groupcode;
	document.frmGM.submit();
}
</script>
<form name="frmGM" method="post" action="dodealitemgroup.asp">
    <input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode" value="del">
	<input type="hidden" name="groupCode">
</form>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �� ��ǰ �׷� ���</div><hr>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
 <tr>
 	<td>
 		<form name="frmG" method="post" action="dodealitemgroup.asp"   onSubmit="return jsGroupSubmit(this);">
		<input type="hidden" name="idx" value="<%=idx%>">
		<input type="hidden" name="mode" value="<%=mode%>">
		<input type="hidden" name="groupCode" value="<%=groupCode%>">
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
							<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">�׷��</td>
							<td bgcolor="#FFFFFF"><textarea name="title" rows="2" cols="40"><%=cdealGroup.Ftitle%></textarea></td>
						</tr>
						<tr>
							<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>
							<td bgcolor="#FFFFFF"><input type="text" size="2" name="sort" class="text" value="<%=cdealGroup.Fsort%>"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td align="center"><p>
	    <input type="button" class="button" style="height:30px; width:100px;" value="����" onClick="jsGroupSubmit();">
	    </p> </td>
</tr>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<% set cdealGroup = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
