<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%

dim page, mastergubun

page = request("page")
mastergubun = request("mastergubun")
if page="" then page=1

if (mastergubun = "") then
	mastergubun = "30"		'// CS����
end if


dim ocsregtemplate
set ocsregtemplate = New CCSTemplate
ocsregtemplate.FCurrPage = page
ocsregtemplate.FPageSize=20
ocsregtemplate.FRectMasterGubun = mastergubun
ocsregtemplate.GetCSTemplateList

dim i
%>
<script language="javascript">

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){
	if (!CheckSelected()){
		alert("���þ������� �����ϴ�.");
		return;
	}

	var ret = confirm("���� �������� �����Ͻðڽ��ϱ�?");

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function AddIttems(){
	if (frmarr.cdl.value == ""){
		alert("ī�װ��� �������ּ���!");
		frmarr.cdl.focus();
	}
	else if (frmarr.linkurl.value == ""){
		alert("��ũ�ּҸ� �Է����ּ���!");
		frmarr.linkurl.focus();
	}
	else if (frmarr.bannerimg.value == ""){
		alert("��� �̹����� �־��ּ���!");
		frmarr.bannerimg.focus();
	}
	else if (confirm("�������� �߰��Ͻðڽ��ϱ�?")){
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function TnGoWrite(){
	document.all.addform.style.display="";
}
</script>

<p>

<input type="button" class="button" value="�űԵ��(<%= GetMasterGubunName(mastergubun) %>)" onClick="location.href='csreg_template_gubun_reg.asp?menupos=<%= menupos %>&mastergubun=<%= mastergubun %>&mode=addgubun'">

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="FFFFFF">
		<td colspan="6">
			�˻���� : <b><%= ocsregtemplate.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ocsregtemplate.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="70">����</td>
		<td>���и�</td>
		<td width="60">ǥ�ü���</td>
		<td width="50">���</td>
		<td width="200">������</td>
	</tr>
	<% for i=0 to ocsregtemplate.FResultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" action="" >
	<input type="hidden" name="idx" value="<%= ocsregtemplate.FItemList(i).Fidx %>">
	<tr align="center" bgcolor="#<% if (ocsregtemplate.FItemList(i).Fisusing = "N") then %>DDDDDD<% else %>FFFFFF<% end if %>">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= ocsregtemplate.FItemList(i).Fgubun %></td>
		<td><a href="csreg_template_gubun_reg.asp?menupos=<%= menupos %>&mode=editgubun&idx=<%= ocsregtemplate.FItemList(i).Fidx %>&mastergubun=<%= ocsregtemplate.FItemList(i).Fmastergubun %>"><%= ocsregtemplate.FItemList(i).Fgubunname %></a></td>
		<td><%= ocsregtemplate.FItemList(i).Fdisporder %></td>
		<td><%= ocsregtemplate.FItemList(i).Fisusing %></td>
		<td><%= ocsregtemplate.FItemList(i).Flastupdate %></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="6" align="center">
			<% if ocsregtemplate.HasPreScroll then %>
				<a href="?page=<%= ocsregtemplate.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ocsregtemplate.StartScrollPage to ocsregtemplate.FScrollCount + ocsregtemplate.StartScrollPage - 1 %>
				<% if i>ocsregtemplate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ocsregtemplate.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set ocsregtemplate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
