<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftShop_cls.asp" -->
<%
'###############################################
' Discription : GIFT SHOP 테마 관리
' History : 2014.04.07 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim themeIdx, isUsing, isPick, isOpen
Dim oGiftShop, lp
Dim page

'// 파라메터 접수
themeIdx = getNumeric(requestCheckVar(request("themeIdx"),10))
isusing = requestCheckVar(request("isusing"),1)
isPick = requestCheckVar(request("isPick"),1)
isOpen = requestCheckVar(request("isOpen"),1)
page = getNumeric(requestCheckVar(request("page"),10))
if isOpen="" then isOpen="Y"				'기본값 공개만
if isPick="" then isPick="A"				'기본값 전체 (A:테마전체, Y:관리테마, N:고객테마)
if page="" then page="1"

'// 페이지정보 목록
Set oGiftShop = new CGiftShop
oGiftShop.FPageSize=15
oGiftShop.FCurrPage=page
oGiftShop.FRectIsOpen = isOpen
oGiftShop.FRectIsUsing = isusing
oGiftShop.FRectIsPick = isPick
oGiftShop.GetThemeList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
  	//검색 버튼
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoOpen").buttonset().children().next().attr("style","font-size:11px;");
});

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 테마를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 테마의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

function goThemeWrite(idx) {
    location.href = '/admin/sitemaster/gift/shop/giftshop_themeWrite.asp?themeidx='+idx+'&menupos=<%= request("menupos") %>';
}

</script>

<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    관리구분:
		<select name="isPick" class="select">
			<option value="A" <%=chkIIF(isPick="A","selected","")%> >테마전체</option>
			<option value="Y" <%=chkIIF(isPick="Y","selected","")%> >10x10 Pick</option>
			<option value="N" <%=chkIIF(isPick="N","selected","")%> >User Pick</option>
		</select>
		&nbsp;/&nbsp;
	    공개여부:
		<select name="isOpen" class="select">
			<option value="A" <%=chkIIF(isOpen="A","selected","")%> >전체</option>
			<option value="Y" <%=chkIIF(isOpen="Y","selected","")%> >공개</option>
			<option value="N" <%=chkIIF(isOpen="N","selected","")%> >비공개</option>
		</select>
		&nbsp;/&nbsp;
	    사용구분:
		<select name="isusing" class="select">
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >사용함</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >사용안함</option>
		</select>
	</td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="검색" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
    <td align="left">
    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
    	<% if C_ADMIN_AUTH then %><input type="button" value="상태저장" class="button" onClick="saveList()" title="우선순위 및 노출여부를 일괄저장합니다."><% end if %>
    </td>
    <td align="right">
    	<input type="button" value="컨텐츠 등록" class="button" onClick="goThemeWrite('');">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oGiftShop.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oGiftShop.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="90" />
    <col width="*" />
    <col width="60" />
    <col width="60" />
    <col width="110" />
    <col width="90" />
    <col width="80" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>&nbsp;</td>
    <td>번호</td>
    <td>관리구분<br>[Pick형태]</td>
    <td>제목</td>
    <td>상품수</td>
    <td>우선<br>순위</td>
    <td>공개여부</td>
    <td>등록일</td>
    <td>등록자</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oGiftShop.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oGiftShop.FItemList(lp).IsOpend,"#FFFFFF","#DDDDDD")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oGiftShop.FItemList(lp).FthemeIdx%>" /></td>
    <td><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FthemeIdx%></a></td>
    <td><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).getPickType%></a>
    	<% if oGiftShop.FItemList(lp).FisPick="Y" then %><br><%=chkIIF(oGiftShop.FItemList(lp).FpickImage="" or isNull(oGiftShop.FItemList(lp).FpickImage),"<span style=""color:#608060;"">[일반형]</span>","<span style=""color:#806060;"">[배너형]</span>")%><% end if %>
    </td>
    <td align="left"><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FSubject & "<br><font color=""#606060"">" & oGiftShop.FItemList(lp).FSubDesc%></font></a></td>
    <td ><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FitemCount%></a></td>
    <td><input type="text" name="sort<%=oGiftShop.FItemList(lp).FthemeIdx%>" size="3" class="text" value="<%=oGiftShop.FItemList(lp).FsortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoOpen">
		<input type="radio" name="open<%=oGiftShop.FItemList(lp).FthemeIdx%>" id="rdoOpen<%=lp%>_1" value="Y" <%=chkIIF(oGiftShop.FItemList(lp).FisOpen="Y","checked","")%> /><label for="rdoOpen<%=lp%>_1">공개</label><input type="radio" name="open<%=oGiftShop.FItemList(lp).FthemeIdx%>" id="rdoOpen<%=lp%>_2" value="N" <%=chkIIF(oGiftShop.FItemList(lp).FisOpen="N","checked","")%> /><label for="rdoOpen<%=lp%>_2">안함</label>
		</span>
    </td>
    <td><%=left(oGiftShop.FItemList(lp).Fregdate,10)%></td>
    <td><%=oGiftShop.FItemList(lp).Fadminname%></td>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center">
    <% if oGiftShop.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGiftShop.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oGiftShop.StartScrollPage to oGiftShop.FScrollCount + oGiftShop.StartScrollPage - 1 %>
		<% if lp>oGiftShop.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oGiftShop.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- 목록 끝 -->
<%
	Set oGiftShop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
