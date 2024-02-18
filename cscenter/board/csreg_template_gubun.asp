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
	mastergubun = "30"		'// CS접수
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
		alert("선택아이템이 없습니다.");
		return;
	}

	var ret = confirm("선택 아이템을 삭제하시겠습니까?");

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
		alert("카테고리를 선택해주세요!");
		frmarr.cdl.focus();
	}
	else if (frmarr.linkurl.value == ""){
		alert("링크주소를 입력해주세요!");
		frmarr.linkurl.focus();
	}
	else if (frmarr.bannerimg.value == ""){
		alert("배너 이미지를 넣어주세요!");
		frmarr.bannerimg.focus();
	}
	else if (confirm("아이템을 추가하시겠습니까?")){
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function TnGoWrite(){
	document.all.addform.style.display="";
}
</script>

<p>

<input type="button" class="button" value="신규등록(<%= GetMasterGubunName(mastergubun) %>)" onClick="location.href='csreg_template_gubun_reg.asp?menupos=<%= menupos %>&mastergubun=<%= mastergubun %>&mode=addgubun'">

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="FFFFFF">
		<td colspan="6">
			검색결과 : <b><%= ocsregtemplate.FTotalcount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ocsregtemplate.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="70">구분</td>
		<td>구분명</td>
		<td width="60">표시순서</td>
		<td width="50">사용</td>
		<td width="200">수정일</td>
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
