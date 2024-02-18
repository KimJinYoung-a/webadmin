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
' Discription : GIFT SHOP 포장상품 관리
' History : 2014.04.08 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim packIdx
Dim oGiftShop, lp, i
Dim page

'// 파라메터 접수
packIdx = getNumeric(requestCheckVar(request("packIdx"),10))
page = getNumeric(requestCheckVar(request("page"),10))
if page="" then page="1"
if packIdx="" then packIdx="1"		'기본값 (플라워:1)

'// 페이지정보 목록
Set oGiftShop = new CGiftShop
oGiftShop.FPageSize=15
oGiftShop.FCurrPage=page
oGiftShop.FRectPackIdx = packIdx
oGiftShop.GetPackageList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
  	//검색 버튼
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$(".rdoOpen").buttonset().children().next().attr("style","font-size:11px;");

	$("input[name='packIdx']").click(function(){
		document.frm.submit();
	});
});

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function fnChkAll(elm) {
	$("#itemList input[name='itemid']").attr("checked",$(elm).is(":checked"));
}

function fnChkDelete() {
	var arrIID="";
	if(!$("#itemList input[name='itemid']").is(":checked")) {
		alert("선택된 상품이 없습니다.");
		return;
	}
	$("#itemList input[name='itemid']:checked").each(function(){
		if(arrIID!="") arrIID += ",";
		arrIID += $(this).val();
	});
	
	window.open("/admin/sitemaster/gift/shop/doPackItemCdArray.asp?packIdx=<%=packIdx%>&mode=d&subItemidArray="+arrIID, "popup_item", "width=300,height=200,scrollbars=yes,resizable=yes");
}

// 상품검색 일괄 등록
function popPackSearchItem() {
    var acUrl = encodeURIComponent("/admin/sitemaster/gift/shop/doPackItemCdArray.asp?packIdx=<%=packIdx%>&mode=i");
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL="+acUrl, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
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
		<span class="rdoOpen">
		<%
			For lp=1 to 6
				Response.Write "<input type=""radio"" name=""packIdx"" id=""rdoOpen" & lp & "_1"" value=""" & lp & """ " & chkIIF(cStr(packIdx)=cStr(lp),"checked","") & " /><label for=""rdoOpen" & lp & "_1"">" & getGiftPackName(lp) & "</label>"
			Next
		%>
		</span>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
    <td align="left">
    	총 <%=oGiftShop.FTotalCount%> 건 /
    	<input type="button" value="삭제" class="button" onClick="fnChkDelete()" />
    </td>
    <td align="right">
    	<input type="button" value="상품 추가" class="button" onClick="popPackSearchItem()" />
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<col width="30" />
<col width="70" />
<col width="70" />
<col width="110" />
<col width="*" />
<col width="100" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td><input type="checkbox" name="chkALL" value="all" onclick="fnChkAll(this)"></td>
    <td>상품코드</td>
    <td>이미지</td>
    <td>브랜드</td>
    <td>상품명</td>
    <td>판매가</td>
    <td>품절여부</td>
    <td>등록일</td>
</tr>
<tbody id="itemList">
<%	For i=0 to oGiftShop.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="itemid" value="<%=oGiftShop.FItemList(i).Fitemid%>"></td>
    <td><%=oGiftShop.FItemList(i).Fitemid%></td>
    <td><img src="<%=oGiftShop.FItemList(i).FsmallImage%>"></td>
    <td><%=oGiftShop.FItemList(i).Fbrandname%></td>
    <td align="left"><%=oGiftShop.FItemList(i).Fitemname%></td>
    <td><%=FormatNumber(oGiftShop.FItemList(i).FsellCash,0)%>원</td>
    <td><%=oGiftShop.FItemList(i).isSoldOut%></td>
    <td><%=Left(oGiftShop.FItemList(i).Fregdate,10)%></td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="8" align="center">
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
</table>
</form>
<%
	Set oGiftShop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
