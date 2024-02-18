<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  브랜드 이미지 등록
' History : 2018-04-16 이종화 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim page, lbrand, i
Dim makerid, isUsing, mode, frmName

mode = requestCheckVar(request("mode"),6)
frmName= requestCheckVar(request("frmName"),32)
if frmName="" then frmName="frm"
page    = requestCheckVar(request("page"),6)
makerid = requestCheckVar(request("makerid"),32)
isUsing = requestCheckVar(request("isusing"),1)
if isUsing="" then isUsing="1"

Response.write makerid

If page = "" Then page = 1

'// 목록 접수
Set lbrand = New cBrandMain
	lbrand.FCurrPage = page
	lbrand.FRectMakerid = makerid
	lbrand.FRectIsUsing = chkIIF(isUsing="A","",isUsing)
	lbrand.FPageSize=20
	lbrand.sBrandImageGetList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
//이미지신규등록 & 수정
function AddNewMainContents(idx){
	var AddNewMainContents = window.open('/admin/brand/brandimage/image_insert.asp?idx='+ idx,'AddNewMainContents','width=1250,height=650,scrollbars=yes,resizable=yes');
	AddNewMainContents.focus();
}

//선택 아이템 상태 적용
function SaveSelectedContents() {
	var selCnt = $("#frmList input:checkbox[name='idx']:checked").length;
	if(selCnt==0) {
		alert("선택된 이미지가 없습니다.");
		return false;
	}

	if(confirm("선택하신 "+selCnt+"건의 이미지를 저장하시겠습니까?")) {
		document.frmList.submit();
	}
}

$(function(){
	$("#frmList input:checkbox[name='idx']").click(function(){
		var ival = $(this).attr("data-idx");
		var iUs = $("#frmList input:radio[name='isus"+ival+"']:checked").val()
		$(this).val(ival+"/"+iUs);
	});
});

function fnSelectIMG(brandimage){
	opener.<%= frmName %>.mainimg.value = brandimage;
	$("#mainimg",opener.document).attr('src', brandimage);
	$("#imgurl",opener.document).html(brandimage);
	self.close();
}
</script>
<img src="/images/icon_arrow_link.gif"> <b>브랜드 이미지 관리</b>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		/ 사용여부
		<select name="isusing" class="select">
			<option value="A" <%=chkIIF(isUsing="A","selected","")%>>전체</option>
			<option value="1" <%=chkIIF(isUsing="1","selected","")%>>사용</option>
			<option value="0" <%=chkIIF(isUsing="0","selected","")%>>사용안함</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit()">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<% if mode="img" then %>
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0;">
<tr>
	<td align="left">
		<font style="color:red">이미지를 클릭하시면 선택 됩니다.</font>
	</td>
</tr>
</table>
<% else %>
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0;">
<tr>
	<td align="left">
		<input type="button" value="선택저장" class="button_auth" onclick="SaveSelectedContents();">
	</td>
	<td align="right">
		<input type="button" value="신규등록" class="button" onclick="AddNewMainContents('0');">
	</td>
</tr>
</table>
<% End If %>
<!-- 액션 끝 -->

<form name="frmList" id="frmList" method="post" action="image_proc.asp" style="margin:0px;">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7">검색결과 : <b><%=lbrand.FTotalCount%></b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% if mode<>"img" then %>
	    <td></td>
		<% End If %>
		<td align="center">No.</td>
		<td align="center" width="200">Image</td>
	    <td align="center">브랜드ID</td>
	    <td align="center">등록일</td>
	    <td align="center">수정일</td>
		<td align="center">사용여부</td>
    </tr>
	<% If lbrand.FResultCount > 0 Then %> 
   	<% For i = 0 to lbrand.FResultCount - 1 %>
    <tr align="center" <%= chkiif(lbrand.FItemList(i).FIsusing,"bgcolor='#FFFFFF'","bgcolor='#DDDDDD'") %> >
		<% if mode<>"img" then %>
	    <td align="center"><input type="checkbox" name="idx" value="" data-idx="<%= lbrand.FItemlist(i).Fidx %>"></td>
		<% End If %>
		<% if mode="img" then %>
		<td align="center"><%= lbrand.FItemlist(i).Fidx %></td>
		<td align="center">
	    	<a href="javascript:fnSelectIMG('<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>');">
	    	<img src="<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>" style="width:100px; border:1px #FDFDFD; border-radius:3px;" />
	    	</a>
	    </td>
		<% else %>
		<td align="center"><a href="javascript:AddNewMainContents('<%= lbrand.FItemlist(i).Fidx %>');"><%= lbrand.FItemlist(i).Fidx %></a></td>
		<td align="center">
			<% if lbrand.FItemlist(i).Fbrandimage<>"" then %>
	    	<a href="javascript:AddNewMainContents('<%= lbrand.FItemlist(i).Fidx %>');">
	    	<img src="<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>" style="width:100px; border:1px #FDFDFD; border-radius:3px;" />
	    	</a>
			<% End If %>
	    </td>
		<% End If %>
	    <td align="center"><%= lbrand.FItemlist(i).Fmakerid %></td>
	    <td align="center"><%= lbrand.FItemlist(i).FRegdate %><br/><%= lbrand.FItemlist(i).Fadminid %></td>
	    <td align="center"><%= lbrand.FItemlist(i).Flastupdate %><br/><%= lbrand.FItemlist(i).Flastadminid %></td>
		<td align="center">
			<label><input type="radio" name="isus<%= lbrand.FItemlist(i).Fidx %>" value="1" <%=chkIIF(lbrand.FItemList(i).FIsusing,"checked","")%> />사용</label>
			<label><input type="radio" name="isus<%= lbrand.FItemlist(i).Fidx %>" value="0" <%=chkIIF(lbrand.FItemList(i).FIsusing,"","checked")%> />사용안함</label>
		</td>
	</tr>
	<% Next %>
	<% Else %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="7" align="center" class="page_link">[등록된 데이터가 없습니다.]</td>
	</tr>
	<% End If %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="7" align="center">
	       	<% If lbrand.HasPreScroll Then %>
				<span class="list_link"><a href="?page=<%= lbrand.StartScrollPage-1 %>">[pre]</a></span>
			<% Else %>
			[pre]
			<% End If %>
			<% for i = 0 + lbrand.StartScrollPage to lbrand.StartScrollPage + lbrand.FScrollCount - 1 %>
				<% if (i > lbrand.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(lbrand.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if lbrand.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->