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
if page="" then page=1

mastergubun = "20"		'// MAIL

dim omailtemplate
set omailtemplate = New CCSTemplate
omailtemplate.FCurrPage = page
omailtemplate.FPageSize=20
omailtemplate.FRectMasterGubun = mastergubun
omailtemplate.GetCSTemplateList

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
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('선택 아이템을 삭제하시겠습니까?');

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
	else if (confirm('아이템을 추가하시겠습니까?')){
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function TnGoWrite(){
	document.all.addform.style.display="";
}
</script>

<p>

<input type="button" class="button" value="신규등록" onClick="location.href='mail_template_gubun_reg.asp?menupos=<%= menupos %>&mode=addgubun'">
<!--
&nbsp;
<input type="button" class="button" value="선택아이템사용안함" onclick="delitems(delform);">
-->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="FFFFFF">
		<td colspan="6">
			검색결과 : <b><%= omailtemplate.FTotalcount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= omailtemplate.FTotalpage %></b>
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
	<% for i=0 to omailtemplate.FResultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" action="" >
	<input type="hidden" name="idx" value="<%= omailtemplate.FItemList(i).Fidx %>">
	<tr align="center" bgcolor="#<% if (omailtemplate.FItemList(i).Fisusing = "N") then %>DDDDDD<% else %>FFFFFF<% end if %>">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= omailtemplate.FItemList(i).Fgubun %></td>
		<td><a href="mail_template_gubun_reg.asp?menupos=<%= menupos %>&mode=editgubun&idx=<%= omailtemplate.FItemList(i).Fidx %>"><%= omailtemplate.FItemList(i).Fgubunname %></a></td>
		<td><%= omailtemplate.FItemList(i).Fdisporder %></td>
		<td><%= omailtemplate.FItemList(i).Fisusing %></td>
		<td><%= omailtemplate.FItemList(i).Flastupdate %></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="6" align="center">
			<% if omailtemplate.HasPreScroll then %>
				<a href="?page=<%= omailtemplate.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + omailtemplate.StartScrollPage to omailtemplate.FScrollCount + omailtemplate.StartScrollPage - 1 %>
				<% if i>omailtemplate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if omailtemplate.HasNextScroll then %>
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
set omailtemplate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->