<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 아티스트 갤러리 이미지 & 링크 파일 생성 리스트 페이지   
' History : 2012.03.26 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
Dim mm, gubun
Dim research,isusing, fixtype, linktype
Dim page
	mm = request("mm")
	isusing = request("isusing")
	page    = request("page")
    isusing = "Y"
    gubun = 2
If page="" Then page=1

Dim oMainContents
set oMainContents = new cposcode_list
	oMainContents.FPageSize = 20
	oMainContents.FCurrPage = page
	oMainContents.FGubun = gubun
	oMainContents.fcontents_list
dim i
%>
<script language="javascript">
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
function AssignReal(upfrm,imagecount){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
			}
		}
	}
	var tot;
	tot = upfrm.fidx.value;
	upfrm.fidx.value = ""

	var AssignimageReal;
	AssignimageReal = window.open("<%=wwwUrl%>/chtml/make_artist_ViewBannerJS.asp?idx=" +tot + '&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
	AssignimageReal.focus();
}

//이미지신규등록 & 수정
function AddNewMainContents(idx){
	var AddNewMainContents = window.open('/admin/artist/imagemake_contents.asp?idx='+ idx + '&gubun=2','AddNewMainContents','width=1200,height=600,scrollbars=yes,resizable=yes');
	AddNewMainContents.focus();
}
document.domain ='10x10.co.kr';
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="40">위치 : 
		<select onchange="location.href=this.value;" class="select">
			<option value="artist_weekmonth.asp?menupos=<%=menupos%>&mm=1" <% If mm = 1 Then response.write "selected"%>>아티스트 뷰 상단배너
			<option value="artist_MonthItemList.asp?menupos=<%=menupos%>&mm=2">아티스트 샵 월별 상품
		</select>		
	</td>
</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<tr><td align="left"><a href="javascript:AssignReal(frm,1);"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a></td></tr>
<tr><td align="left"><input type="button" value="신규등록" class="button" onClick="javascript:AddNewMainContents('0');"></td></tr>
</form>
</table>
<!-- 액션 끝 -->
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If oMainContents.FResultCount > 0 Then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">검색결과 : <b><%= oMainContents.FTotalCount %></b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
 		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	    <td align="center">Idx</td>
	    <td align="center">Image</td>
	    <td align="center">등록일</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% For i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr bgcolor="#DDDDDD">
		<% else %>
			<tr bgcolor="#FFFFFF">
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= oMainContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<img width=40 height=40 src="<%=uploadUrl%>/artist/<%= oMainContents.FItemList(i).fimagepath %>" border="0">
	    	</a>
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fregdate %></td> 
	</tr>
	</form>	
	<% Next %>
    </tr>   
<% Else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% End If %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% If oMainContents.HasPreScroll Then %>
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>">[pre]</a></span>
			<% Else %>
			[pre]
			<% End If %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->