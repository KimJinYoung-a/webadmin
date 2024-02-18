<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/board/qna_prefacecls.asp"-->
<%

dim page

page = request("page")
if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectmasterid = "01"
omd.GetPrefaceGubunList

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


<table width="600" align="left" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;">
<tr>
	<td>

<!-- �׼� ���� -->
<table width="600" align="left" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�űԵ��" onClick="location.href='cs_qna_preface_gubun_reg.asp?menupos=<%= menupos %>&mode=add'">
			&nbsp;
			<input type="button" class="button" value="���þ����ۻ�����" onclick="delitems(delform);">
		</td>
		<td align="right">
			
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

	</td>
</tr>
<tr>
	<td>

<table width="600" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= omd.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= omd.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center">�ڵ�</td>
		<td align="center">�з�����</td>
	</tr>
	<% for i=0 to omd.FResultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" action="" >
	<input type="hidden" name="itemid" value="<%= omd.FItemList(i).Fcode %>">
	<tr bgcolor="#FFFFFF">
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td align="center"><%= omd.FItemList(i).Fcode %></td>
		<td align="center"><a href="cs_qna_preface_gubun_reg.asp?menupos=<%= menupos %>&mode=edit&code=<%= omd.FItemList(i).Fcode %>"><%= omd.FItemList(i).Fcname %></a></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center">
		<% if omd.HasPreScroll then %>
			<a href="?page=<%= omd.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1 %>
			<% if i>omd.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if omd.HasNextScroll then %>
			<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

	</td>
</tr>
</table>

<form name="delform" method="post" action="prefacegubun_del_process.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="masterid" value="01">
</form>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->