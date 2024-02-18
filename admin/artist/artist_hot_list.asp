<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 아티스트 브랜드 관리 페이지   
' History : 2012.03.27 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
'// 변수 선언
Dim page, isusing, designerid, i, mm
	mm = request("mm")
	page = request("page")
	isusing = request("isusing")
	designerid = request("designerid")
	
	if page="" then page=1
	if isusing="" then isusing=""

'// 목록 접수
Dim oGallery
	set oGallery = New cposcode_list
	oGallery.FCurrPage = page
	oGallery.FPageSize=20
	oGallery.Hotorder = "Y"
	oGallery.FArtistBrandList
%>
<script>
function goView(ii){
	location.href = "artist_brand_write.asp?mode=edit&idx="+ii;
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}
function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function changeSort(upfrm) {
	if (!CheckSelected()){
		alert('체크 해주세요.');
		return;
	}
	var ret = confirm('선택하신 순서를 지정하신 번호로 변경하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.sortNo.value  = upfrm.sortNo.value + frm.sortNo.value + "," ;
					upfrm.idx.value 	= upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="hot_ChangeSort";
		upfrm.submit();
	}
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('체크 해주세요.');
		return;
	}

	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택하신 브랜드를 메인에 노출시킵니다.');
	} else {
		var ret = confirm('선택하신 브랜드를 메인에서 뺍니다.');
	}


	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="hot_isUsingValue";
		upfrm.submit();

	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="40">위치 : 
		<select onchange="location.href=this.value;" class="select">
			<option value="artist_main.asp?menupos=<%=menupos%>&mm=1" <% If mm = 1 Then response.write "selected"%>>메인 멀티배너
			<option value="artist_hot_list.asp?menupos=<%=menupos%>&mm=2" <% If mm = 2 Then response.write "selected"%>>HOT ARTIST
			<option value="artist_notice_board_list.asp?menupos=<%=menupos%>&mm=3">공지사항
			<option value="artist_selectshop.asp?menupos=<%=menupos%>&mm=4">Select Shop
		</select>		
	</td>
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>HOT ARTIST</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="arrFrm" method="post" action="doMDChoice.asp">
<input type="hidden" name="idx">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
<input type="hidden" name="sortNo">
<input type="hidden" name="fidx">
<tr>
	<td colspan="8" align="left" bgcolor="WHITE">
		<select name="allusing"  class="select">
			<option value="Y">선택 사용 -> Y </option>
			<option value="N">선택 사용 -> N </option>
		</select><input type="button" value="적용" class="button" onclick="changeUsing(arrFrm);"> /
		<input type="button" value="순서적용" class="button" onclick="changeSort(arrFrm);">
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="50">번호</td>
	<td width="190">브랜드</td>
	<td>서브 배너 이미지</td>
	<td width="60">등록일</td>
	<td width="60">사용</td>
	<td width="60">프론트노출</td>
	<td width="60">순서</td>
</tr>
<% If oGallery.FTotalCount = 0 Then %>
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'">
	<td align="center" colspan="6">[데이터가 없습니다.]</td>
</tr>
<% End If %>

<% For i=0 to oGallery.FResultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="idx" value="<%= oGallery.FItemList(i).FIdx %>">
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center" width="50"><%=oGallery.FItemList(i).fidx%></td>
	<td align="center" width="190"><%=oGallery.FItemList(i).fdesignerid%></td>
	<td><img src="<%=uploadUrl%>/artist/brandbanner/<%=oGallery.FItemList(i).ffile2%>" height="50" onClick="goView('<%=oGallery.FItemList(i).fidx%>')" style="cursor:pointer" ></td>
	<td align="center" width="160"><%=oGallery.FItemList(i).fregdate%></td>
	<td align="center" width="60"><%=oGallery.FItemList(i).fisusing%></td>
	<td align="center" width="60"><%=oGallery.FItemList(i).fmainHOT%></td>
	<td align="center" width="60"><input type="text" name="sortNo" value="<%= oGallery.FItemList(i).FsortNo %>" size="3" style="text-align:right;" /></td>
</tr>
</form>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="8" align="center">
       	<% If oGallery.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + oGallery.StartScrollPage to oGallery.StartScrollPage + oGallery.FScrollCount - 1 %>
			<% If (i > oGallery.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(oGallery.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If oGallery.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</table>
<form name="frm" method="get">
<input type="hidden" name="page">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->