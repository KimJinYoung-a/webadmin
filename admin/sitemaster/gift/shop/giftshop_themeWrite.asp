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
' Discription : GIFT SHOP 테마 상품 관리
' History : 2014.04.07 허진원 : 신규 생성
'###############################################

	'// 변수 선언
	Dim oGiftShop, i
	Dim themeIdx, subject, subDesc, userid, regdate, frontItemid, isOpen, isPick, isUsing, tag, sortNo
	Dim viewCount, commentCount, pickImage

	'// 파라메터 접수
	themeIdx = getNumeric(requestCheckVar(request("themeIdx"),10))

	'// 테마 정보 접수
	if themeIdx<>"" then
		Set oGiftShop = new CGiftShop
		oGiftShop.FRectIdx = themeIdx
		oGiftShop.GetThemeInfo
		if oGiftShop.FResultCount>0 then
			subject		= oGiftShop.FOneItem.Fsubject
			subDesc		= oGiftShop.FOneItem.FsubDesc
			userid		= oGiftShop.FOneItem.Fuserid
			regdate		= oGiftShop.FOneItem.Fregdate
			frontItemid	= oGiftShop.FOneItem.FfrontItemid
			isOpen		= oGiftShop.FOneItem.FisOpen
			isPick		= oGiftShop.FOneItem.FisPick
			isUsing		= oGiftShop.FOneItem.FisUsing
			sortNo		= oGiftShop.FOneItem.FsortNo
			tag			= oGiftShop.FOneItem.Ftag
			viewCount	= oGiftShop.FOneItem.FviewCount
			commentCount	= oGiftShop.FOneItem.FcommentCount
			pickImage	= oGiftShop.FOneItem.FpickImage
		end if
		Set oGiftShop = Nothing
	end if

	if sortNo="" then sortNo=0
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function(){
	$(".chkBox").buttonset().children().next().attr("style","font-size:11px;");
	$(".btn").button();
});

function fnSelKeyword(elm){
	if($("#lyrKeyword input[type='checkbox']:checked").length>3) {
		alert("키워드는 3개까지 선택 가능합니다.");
		$(elm).attr("checked",false);
	}
	var keyId="";
	$("#lyrKeyword input[type='checkbox']:checked").each(function(){
		if(keyId!="") keyId += ",";
		keyId += $(this).val();
	});
	document.frm.arrKeyIdx.value = keyId;
}

function SaveTheme(frm){
	if(frm.subject.value=="") {
		alert("테마 제목을 입력해주세요.");
		frm.subject.focus();
		return;
	}

	if(frm.sortNo.value=="") {
		alert("테마 정렬 우선순위를 입력해주세요.");
		frm.sortNo.focus();
		return;
	}

	if(!$("input[type='checkbox']").is(":checked")) {
		alert("키워드를 선택해주세요.");
		return;
	}

	if(frm.isOpen.value=="Y"&&$("#itemList input[name='itemid']").length<4) {
		alert("등록된 상품이 4개 미만일 경우에는 공개설정을 할 수 없습니다.");
		return;
	}

	if(frm.isOpen.value=="Y"&&frm.frontItemid.value=="0") {
		alert("대표 상품을 지정해주세요.");
		return;
	}

	frm.submit();
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
	
	window.open("/admin/sitemaster/gift/shop/doRegItemCdArray.asp?themeIdx=<%=themeIdx%>&mode=d&subItemidArray="+arrIID, "popup_item", "width=300,height=200,scrollbars=yes,resizable=yes");
}

// 상품검색 일괄 등록
function popRegSearchItem() {
    var acUrl = encodeURIComponent("/admin/sitemaster/gift/shop/doRegItemCdArray.asp?themeIdx=<%=themeIdx%>&mode=i");
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL="+acUrl, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_theme_upload.asp?yr=<%=Year(regdate)%>&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   $("#"+sName).val('');
	   $("#"+sSpan).fadeOut();
	}
}

function fnChgIsPick(elm) {
	if($(elm).val()=="Y") {
		$("#rowTTImg").show();
	} else {
		$("#rowTTImg").hide();
	}
}

function fnChkFrontItem(iid) {
	document.frm.frontItemid.value=iid;
}
</script>
<!-- 메인페이지 정보 시작 -->
<form name="frm" method="POST" action="doGiftShopTheme.asp" style="margin:0;">
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<input type="hidden" name="frontItemid" value="<%= frontItemid %>" />
<input type="hidden" name="arrKeyIdx" value="<%= tag %>" />
<input type="hidden" name="mode" value="<%=chkIIF(themeIdx="","i","u")%>" />
<p><b>▶ 테마 정보</b></p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<colgroup>
	<col width="120" />
	<col width="*" />
	<col width="120" />
	<col width="*" />
</colgroup>
<% if themeIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">테마 번호</td>
    <td>
        <%=themeIdx %>
        <input type="hidden" name="themeIdx" value="<%=themeIdx %>" />
    </td>
    <td bgcolor="#DDDDFF">등록일시</td>
    <td>
        <%=regdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">조회수</td>
    <td>
        <%=viewCount %>
    </td>
    <td bgcolor="#DDDDFF">댓글수</td>
    <td>
        <%=commentCount %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">테마 제목 <span style="color:#F03030" title="필수">＊</span></td>
    <td colspan="3">
		<input type="text" name="subject" size="24" maxlength="18" value="<%=subject%>" class="text" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">부가 설명</td>
    <td colspan="3">
		<input type="text" name="subDesc" size="60" maxlength="40" value="<%=subDesc%>" class="text" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">공개여부</td>
    <td>
		<select name="isOpen" class="select">
		<option value="Y" <%=chkIIF(isOpen="Y","selected","")%>>공개</option>
		<option value="N" <%=chkIIF(isOpen="N" or isOpen="","selected","")%>>비공개</option>
		</select>
    </td>
    <td bgcolor="#DDDDFF">관리여부</td>
    <td>
		<select name="isPick" class="select" onchange="fnChgIsPick(this)">
		<option value="Y" <%=chkIIF(isPick="Y" or isPick="","selected","")%>>10x10's Pick</option>
		<option value="N" <%=chkIIF(isPick="N","selected","")%>>User's Pick</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">우선순위 <span style="color:#F03030" title="필수">＊</span></td>
    <td>
		<input type="text" name="sortNo" size="4" value="<%=sortNo%>" class="text" />
    </td>
    <td bgcolor="#DDDDFF">사용여부</td>
    <td>
		<select name="isUsing" class="select">
		<option value="Y" <%=chkIIF(isUsing="Y" or isUsing="","selected","")%>>사용</option>
		<option value="N" <%=chkIIF(isUsing="N","selected","")%>>삭제</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">테마 키워드 <span style="color:#F03030" title="필수">＊</span></td>
    <td colspan="3" id="lyrKeyword">
		<%=getGiftKeyword("fnSelKeyword(this)",tag)%>
    </td>
</tr>
<% if (isPick="Y" or themeIdx="") or date<"2014-04-15" then %>
<tr bgcolor="#FFFFFF" id="rowTTImg" style="<%=chkIIF(isPick="Y" or themeIdx="","","display:none;")%>">
    <td bgcolor="#DDDDFF">타이틀 이미지</td>
    <td colspan="3">
		<input type="hidden" name="pickImage" id="pickImage" value="<%=pickImage%>">
		<input type="button" value="이미지 등록" onClick="jsSetImg('<%=pickImage%>','pickImage','lyTitleImg')" class="button">
		<span style="color:#A06060; font-size:11px;">※ 1100 × 170px (200kb이하의 JPEG, GIF, PNG)</span>
		<div id="lyTitleImg" style="padding: 5 5 5 5">
		<% if Not(pickImage="" or isNull(pickImage)) then %>
			<img src="<%=pickImage%>" width="100%">
			<a href="javascript:jsDelImg('pickImage','lyTitleImg');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
		</div>
    </td>
</tr>
<% end if %>
<tr bgcolor="#F8F8F8">
    <td colspan="4" align="center">
    	<input type="button" value=" 취소 " onClick="history.back();" class="btn"> &nbsp;
    	<input type="button" value=" 저 장 " onClick="SaveTheme(this.form);" class="btn">
    </td>
</tr>
</table>
</form>
<%
	'// 등록된 테마라면
	if themeIdx<>"" then
		Set oGiftShop = new CGiftShop
		oGiftShop.FPageSize=200
		oGiftShop.FRectIdx = themeIdx
		oGiftShop.GetThemeItemList
%>
<p><b>▶ 상품 정보</b></p>
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="7">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oGiftShop.FTotalCount%> 건 /
		    	<input type="button" value="삭제" class="button" onClick="fnChkDelete()" />
		    </td>
		    <td align="right">
		    	<input type="button" value="상품 추가" class="button" onClick="popRegSearchItem()" />
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="90" />
<col width="70" />
<col width="*" />
<col width="110" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td><input type="checkbox" name="chkALL" value="all" onclick="fnChkAll(this)"></td>
    <td>상품코드(대표)</td>
    <td>이미지</td>
    <td>상품명</td>
    <td>판매가</td>
    <td>품절여부</td>
    <td>등록일</td>
</tr>
<tbody id="itemList">
<%	For i=0 to oGiftShop.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="itemid" value="<%=oGiftShop.FItemList(i).Fitemid%>"></td>
    <td>
    	<label>
    	<input type="radio" name="chkFront" value="<%=oGiftShop.FItemList(i).Fitemid%>" <%=chkIIF(oGiftShop.FItemList(i).Fitemid=frontItemid,"checked","")%> onclick="fnChkFrontItem(this.value)">
    	<%=oGiftShop.FItemList(i).Fitemid%>
    	</label>
    </td>
    <td><img src="<%=oGiftShop.FItemList(i).FsmallImage%>"></td>
    <td align="left">
    	<font color="#606060">[<%=oGiftShop.FItemList(i).Fbrandname%>]</font>
    	<%=oGiftShop.FItemList(i).Fitemname%>
    </td>
    <td><%=FormatNumber(oGiftShop.FItemList(i).FsellCash,0)%>원</td>
    <td><%=oGiftShop.FItemList(i).isSoldOut%></td>
    <td><%=Left(oGiftShop.FItemList(i).Fregdate,10)%></td>
</tr>
<%	Next %>
</tbody>
</table>
</form>
<%
		Set oGiftShop = Nothing
	end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->