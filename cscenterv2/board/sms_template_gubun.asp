<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/board/cs_templatecls.asp"-->
<%

dim page, mastergubun

page = RequestCheckvar(request("page"),10)
if page="" then page=1

mastergubun = "10"		'// SMS

dim osmstemplate
set osmstemplate = New CCSTemplate
osmstemplate.FCurrPage = page
osmstemplate.FPageSize=20
osmstemplate.FRectMasterGubun = mastergubun
osmstemplate.GetCSTemplateList

dim i
%>
<script language='javascript'>

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
		alert('���þ������� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');

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
	else if (confirm('�������� �߰��Ͻðڽ��ϱ�?')){
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function TnGoWrite(){
	document.all.addform.style.display="";
}
</script>

<p>

<input type="button" class="button" value="�űԵ��" onClick="location.href='sms_template_gubun_reg.asp?menupos=<%= menupos %>&mode=addgubun'">
<!--
&nbsp;
<input type="button" class="button" value="���þ����ۻ�����" onclick="delitems(delform);">
-->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="FFFFFF">
		<td colspan="6">
			�˻���� : <b><%= osmstemplate.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= osmstemplate.FTotalpage %></b>
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
	<% for i=0 to osmstemplate.FResultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" action="" >
	<input type="hidden" name="idx" value="<%= osmstemplate.FItemList(i).Fidx %>">
	<tr align="center" bgcolor="#<% if (osmstemplate.FItemList(i).Fisusing = "N") then %>DDDDDD<% else %>FFFFFF<% end if %>">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= osmstemplate.FItemList(i).Fgubun %></td>
		<td><a href="sms_template_gubun_reg.asp?menupos=<%= menupos %>&mode=editgubun&idx=<%= osmstemplate.FItemList(i).Fidx %>"><%= osmstemplate.FItemList(i).Fgubunname %></a></td>
		<td><%= osmstemplate.FItemList(i).Fdisporder %></td>
		<td><%= osmstemplate.FItemList(i).Fisusing %></td>
		<td><%= osmstemplate.FItemList(i).Flastupdate %></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="6" align="center">
			<% if osmstemplate.HasPreScroll then %>
				<a href="?page=<%= osmstemplate.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + osmstemplate.StartScrollPage to osmstemplate.FScrollCount + osmstemplate.StartScrollPage - 1 %>
				<% if i>osmstemplate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osmstemplate.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>


<form name="delform" method="post" action="complimentgubun_del_process.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="masterid" value="01">
</form>
<%
set osmstemplate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
