<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<!-- #include virtual="/lib/classes/sitemasterclass/GiftCardImageCls.asp" -->					   
<%
	dim page
	dim i
	dim giftCardImgList
	dim currentPath	
    dim isusing

	currentPath = request.ServerVariables("PATH_INFO")	    	
    isusing             = request("isusing")
	page 				= request("page")

    if isusing = "" then
        isusing = "1"
    end if

	if page="" then page=1	

	set giftCardImgList = new GiftCardImageCls
	giftCardImgList.FPageSize			= 20
	giftCardImgList.FCurrPage			= page	
    giftCardImgList.FRectIsusing		= isusing
    giftCardImgList.GetContentsList
%>
<style type="text/css">

</style>
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
$(function(){
	$("li a").click(function(e){
		e.stopPropagation();
	});	
	$("li span").click(function(e){
		e.stopPropagation();
	});				
    $('#datepicker').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yymm'        
    });	
    $('#startDate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',        
    });		
    $('#endDate').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',        
    });			
})
function jsmodify(v){
	location.href = "addgiftcardimage.asp?idx="+v;
}
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}
function jsOpen(sPURL,sTG){ 
	if (sTG =="M" ){ 
		var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	}
}
</script>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>[CS]GIFT카드 &gt; <strong><a href="">기프트카드 이미지등록</a></strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
		</div>
	</div>	
	<div class="tab" style="margin:0 0 0 -1px;">
		<ul>
			<li class="col11 <%=chkIIF(currentPath = "/admin/giftcardimage/index.asp","selected","")%> "><a href="index.asp">이미지등록</a></li>			
		</ul>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm" method="post" style="margin:0px;" action="/admin/giftcardimage/index.asp">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<!-- search -->
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<p class="formTit">상태 :</p>
					<select class="formSlt" id="open" title="옵션 선택" name="isusing">						
						<option value="1" <%=chkIIF(isusing="1" or isusing="","selected","")%>>사용</option>
						<option value="0" <%=chkIIF(isusing="0","selected","")%>>사용 안함</option>
					</select>
				</li>
			</ul>
		</div>	
		<dfn class="line"></dfn>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btnRegist btn bold fs12" value="이미지 등록" onclick="document.location.href='addgiftcardimage.asp'"/>
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">총 등록수 : <strong><%=giftCardImgList.FtotalCount%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
                            <p style="width:5%">designid</p>
                            <p style="width:30%">이미지</p>                            
							<p style="width:10%">최초 등록정보</p>
							<p style="width:10%">최종 수정정보</p>								
						</li>
					</ul>
					<!-- 리스트 -->
					<ul class="tbDataList">
<% 
	for i=0 to giftCardImgList.FResultCount-1 
%>					
						<li style="cursor:pointer;" onclick="jsmodify(<%=giftCardImgList.FItemList(i).Fidx%>)" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">
                            <p style="width:5%"><%=giftCardImgList.FItemList(i).FdesignId%></p>
                            <p style="width:30%">
                                <img id="DsnImg" src="<%=giftCardImgList.FItemList(i).FGiftCardImage%>" width="455" height="275" alt="카드디자인">
                            </p>                            
							<p style="width:10%"><%=giftCardImgList.FItemList(i).FAdminName%><br /><%=giftCardImgList.FItemList(i).FRegistDate%></p>
							<p style="width:10%"><%=giftCardImgList.FItemList(i).FAdminModifyerName%><br /><%=giftCardImgList.FItemList(i).FLastUpDate%></p>							
						</li>						
<% Next %>						
					</ul>
					<div class="ct tPad20 cBk1">
						<% if giftCardImgList.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= giftCardImgList.StartScrollPage-1 %>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + giftCardImgList.StartScrollPage to giftCardImgList.StartScrollPage + giftCardImgList.FScrollCount - 1 %>
							<% if (i > giftCardImgList.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(giftCardImgList.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if giftCardImgList.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
						<% else %>
						[next]
						<% end if %>						
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
