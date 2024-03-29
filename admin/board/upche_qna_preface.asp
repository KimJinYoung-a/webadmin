<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/qna_prefacecls.asp"-->
<%

dim gubun, page
gubun = request("gubun")
page = request("page")
if page="" then page=1

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = page
omd.FPageSize=10
omd.FRectGubun = gubun
omd.FRectmasterid = "03"	'업체게시판
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

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			분류 : <% SelectBoxQnaPrefaceGubun "03",gubun %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="신규등록" onClick="location.href='upche_qna_preface_reg.asp?mode=add&menupos=<%=menupos%>'">
			&nbsp;
			<input type="button" class="button" value="선택아이템사용안함" onclick="delitems(delform);">
		</td>
		<td align="right">
		
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= omd.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= omd.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="30" align="center">코드</td>
		<td width="150" align="center">분류</td>
		<td align="center">내용</td>
		<td align="center" width="50">사용유무</td>
		<td align="center" width="100">등록일</td>
	</tr>
	<% for i=0 to omd.FResultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" action="" >
	<input type="hidden" name="itemid" value="<%= omd.FItemList(i).Fidx %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= omd.FItemList(i).Fgubun %></td>
		<td><%= omd.FItemList(i).Fcname %></td>
		<td align="left"><a href="upche_qna_preface_reg.asp?mode=edit&idx=<%= omd.FItemList(i).Fidx %>&menupos=<%=menupos%>"><%= omd.FItemList(i).Fcontents %></a></td>
		<td><%= fnColor(omd.FItemList(i).Fisusing , "yn") %></td>
		<td><%= FormatDate(omd.FItemList(i).Fregdate,"0000.00.00") %></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center">
		<% if omd.HasPreScroll then %>
			<a href="?page=<%= omd.StartScrollPage-1 %>&gubun=<% =gubun %>&menupos=<%= menupos %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1 %>
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

<form name="delform" method="post" action="upche_preface_del_proc.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
</form>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->