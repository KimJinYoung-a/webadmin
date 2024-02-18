<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/qna_prefacecls.asp"-->
<%

dim gubun, page
gubun = request("gubun")
page = request("page")
if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectGubun = gubun
omd.FRectmasterid = "02"
omd.GetMDSRecommendList

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
<table width="650" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td><a href="user_qna_preface_reg.asp?mode=add"><font color="red">New</font></a></td>
</tr>
</table>
<table width="650" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a">
		분류 :
			<% SelectBoxQnaPrefaceGubun "02",gubun %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="650" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF" height="25">
	<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="100" align="center">분류</td>
	<td align="center">내용</td>
	<td align="center" width="50">사용유무</td>
	<td align="center" width="100">등록일</td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name="itemid" value="<%= omd.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= omd.FItemList(i).GubunName %></td>
	<td align="center"><a href="user_qna_preface_reg.asp?mode=edit&idx=<%= omd.FItemList(i).Fidx %>"><%= omd.FItemList(i).Fcontents %></a></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center"><%= FormatDate(omd.FItemList(i).Fregdate,"0000.00.00") %></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&gubun=<% =gubun %>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&gubun=<% =gubun %>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&gubun=<% =gubun %>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="650">
<tr>
	<td><input type="button" value="선택아이템사용안함" onclick="delitems(delform);" class="button"></td>
</tr>
</table>
<form name="delform" method="post" action="delpreface.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
</form>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->