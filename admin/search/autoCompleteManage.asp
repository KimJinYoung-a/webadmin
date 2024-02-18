<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)
'// 즐겨찾기
dim IsMenuFavoriteAdded
IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)


Dim i, cAuto, vIdx, vAutoType, vTitle, vURL_PC, vURL_M, vIcon, vRegUserName, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN, vSortNo
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cAuto = New CSearchMng
	cAuto.FRectIdx = vIdx
	cAuto.sbAutoCompleteDetail

	vAutoType = cAuto.FOneItem.Fautotype
	vTitle = cAuto.FOneItem.Ftitle
	vURL_PC = cAuto.FOneItem.Furl_pc
	vURL_M = cAuto.FOneItem.Furl_m
	vIcon = cAuto.FOneItem.Ficon
	vRegUserName = cAuto.FOneItem.Fregusername
	vRegdate = cAuto.FOneItem.Fregdate
	vLastUserName = cAuto.FOneItem.Flastusername
	vLastdate = cAuto.FOneItem.Flastdate
	vMemo = cAuto.FOneItem.Fmemo
	vSortNo = cAuto.FOneItem.Fsortno
	vUseYN = cAuto.FOneItem.Fuseyn
	Set cAuto = Nothing
Else
	vUseYN = "y"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
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

function jsAutoCompleteSave(){
	if($(":radio[name=autotype]:checked").length == "0"){
		alert("자동완성 속성을 선택하세요.");
		return;
	}
	if($("#title").val() == ""){
		alert("제목을 입력하세요.");
		return;
	}
	if(!frm1.autotype[3].checked){
		if($("#url_pc").val() == ""){
			alert("URL PC를 입력하세요.");
			return;
		}
		if($("#url_m").val() == ""){
			alert("URL M를 입력하세요.");
			return;
		}
	}else{
		$("#url_pc").val("");
		$("#url_m").val("");
	}
	if($(":radio[name=icon]:checked").length == "0"){
		alert("아이콘 설정을 선택하세요.");
		return;
	}

	frm1.submit();
}

//링크값선택
function showDrop(g){
	$("#selectLink"+g+"").show();
}

function linkcopy(g){
	var val = $("#url_"+g+"").val();
	$("#selectLink"+g+"").css("display","none");
}

//선택입력
function populateTextBox(v,g){
	var val = v;
	$("#url_"+g+"").val(val);
	$("#selectLink"+g+"").css("display","none");
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<div class="contSectFix scrl">
	<div class="cont">
		<form name="frm1" action="autoCompleteProc.asp" method="post">
		<input type="hidden" name="idx" value="<%=vIdx%>">
		<div class="searchWrap inputWrap">
			<h3>- 자동완성 정보</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>자동완성 속성 *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="sc" name="autotype" value="sc" <%=CHKIIF(vAutoType="sc","checked","")%> /> <label for="sc">바로가기</label></span>
						<span class="rMar10"><input type="radio" id="ca" name="autotype" value="ca" <%=CHKIIF(vAutoType="ca","checked","")%> /> <label for="ca">카테고리</label></span>
						<span class="rMar10"><input type="radio" id="br" name="autotype" value="br" <%=CHKIIF(vAutoType="br","checked","")%> /> <label for="br">브랜드</label></span>
						<span class="rMar10"><input type="radio" id="ky" name="autotype" value="ky" <%=CHKIIF(vAutoType="ky","checked","")%> /> <label for="ky">키워드</label></span>
					</td>
				</tr>
				<tr>
					<th><div>제목 *</div></th>
					<td><input type="text" class="formTxt" id="title" name="title" value="<%=vTitle%>" maxlength="20" placeholder="20자 이내의 자동완성 제목을 입력해주세요." style="width:50%" /></td>
				</tr>
				<tr>
					<th></th>
					<td><strong>
						<font color="blue">※ 자동완성 속성이 "바로가기" 인 경우<br />카테고리, 브랜드, 키워드 등 앱 네이티브로 열리는 페이지는 불가합니다. 웹뷰페이지만 입력하세요.</font>
					</strong></td>
				</tr>
				<tr>
					<th><div>URL PC *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_PC="","링크값 입력(선택)",vURL_PC)%>" onclick="showDrop('pc');" id="url_pc" name="url_pc" onkeyup="linkcopy('pc');" maxlength="200" />
							<ul style="display:none;" id="selectLinkpc">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_PC="","",vURL_PC)%>','pc');">선택안함</li>
								<li onclick="populateTextBox('/category/category_prd.asp?itemid=상품코드','pc');">/category/category_prd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/shopping/category_list.asp?disp=카테고리','pc');">/shopping/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','pc');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','pc');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','pc');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','pc');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','pc');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>URL M *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_M="","링크값 입력(선택)",vURL_M)%>" onclick="showDrop('m');" id="url_m" name="url_m" onkeyup="linkcopy('m');" maxlength="200" />
							<ul style="display:none;" id="selectLinkm">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_M="","",vURL_M)%>','m');">선택안함</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=상품코드','m');">/category/category_itemPrd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리','m');">/category/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','m');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','m');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','m');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','m');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','m');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>아이콘 설정 *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="none" name="icon" value="none" <%=CHKIIF(vIcon="none","checked","")%> /> <label for="none">사용안함</label></span>
						<span class="rMar10"><input type="radio" id="best" name="icon" value="best" <%=CHKIIF(vIcon="best","checked","")%> /> <label for="best">베스트</label></span>
						<span class="rMar10"><input type="radio" id="jump" name="icon" value="jump" <%=CHKIIF(vIcon="jump","checked","")%> /> <label for="jump">급상승 검색어</label></span>
					</td>
				</tr>
				<tr>
					<th><div>사용여부 *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">사용함</label></span>
						<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">사용안함</label></span>
					</td>
				</tr>
				<% If vIdx <> "" Then %>
				<tr>
					<th><div>작성자</div></th>
					<td>최초작업자 : <%=vRegUserName%>, 마지막작업자 : <%=vLastUserName%></td>
				</tr>
				<tr>
					<th><div>작성일</div></th>
					<td>최초작성일 : <%=vRegdate%>, 마지막작성일 : <%=vLastdate%></td>
				</tr>
				<% End If %>
				<tr>
					<th><div>비고</div></th>
					<td><textarea class="formTxtA" rows="6" style="width:99%;" id="memo" name="memo"><%=vMemo%></textarea></td>
				</tr>
				</tbody>
			</table>
			<div class="tMar20 ct">
				<input type="button" value="저장" onclick="jsAutoCompleteSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			</div>
		</div>
		</form>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->