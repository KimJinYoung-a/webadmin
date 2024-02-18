<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 플레이모바일
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type='text/javascript'>

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
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USB키 처리 시작 (2008.06.23;허진원) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<%
	Dim i, cPlay, vPage, vIsUsing, vType, vTitle, vPartMDID, vPartWDID, vPartPBID, vState
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vIsUsing = requestCheckVar(request("isusing"),1)
	vType = requestCheckVar(request("playtype"),2)
	vTitle = requestCheckVar(request("title"),200)
	vState = requestCheckVar(request("state"),2)
	vPartMDID = NullFillWith(requestCheckVar(request("partmdid"),50),"")
	vPartWDID = NullFillWith(requestCheckVar(request("partwdid"),50),"")
	vPartPBID = NullFillWith(requestCheckVar(request("partpbid"),50),"")
	
	SET cPlay = New CPlayMoContents
	cPlay.FCurrPage = vPage
	cPlay.FPageSize = 20
	cPlay.FRectIsusing = vIsUsing
	cPlay.FRectType = vType
	cPlay.FRectTitle = vTitle
	cPlay.FRectState = vState
	cPlay.FRectMDID = vPartMDID
	cPlay.FRectWDID = vPartWDID
	cPlay.FRectPBID = vPartPBID
	cPlay.fnPlayMoList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goNewReg(idx){
	var winPlay;
	winPlay = window.open('write.asp?idx='+idx,'winPlay','width=1400, height=800, scrollbars=yes');
	winPlay.focus();
}
function goPlayType(idx){
	var winPlayType;
	winPlayType = window.open('pop_type.asp','winPlayType','width=410, height=570');
	winPlayType.focus();
}
function goStyleCode(idx){
	var winStyleCode;
	winStyleCode = window.open('pop_style.asp','winStyleCode','width=410, height=570');
	winStyleCode.focus();
}
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}
function jsTagview(idx,type){	
	var poptagm;
	poptagm = window.open('pop_tagReg.asp?idx='+idx+'&playcate='+type+'','poptagm','width=500,height=400,scrollbars=yes,resizable=yes');
	poptagm.focus();
}
function jsItem(idx,type){
	var popPItem;
	popPItem = window.open('item.asp?playidx='+idx+'&playcate='+type+'','popPItem','width=1200,height=1000,scrollbars=yes,resizable=yes');
	popPItem.focus();
}
</script>

<div class="contSectFix scrl">
	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">사용 여부 :</label>
					<select name="isusing" class="formSlt">
						<option value=""> - 선택 - </option>
						<option value="Y" <%=CHKIIF(vIsUsing="Y","selected","")%>>Y(사용함)</option>
						<option value="N" <%=CHKIIF(vIsUsing="N","selected","")%>>N(사용안함)</option>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">상 태 :</label>
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">분 류 :</label>
					<select name="playtype" class="formSlt">
						<%=fnTypeSelectBox("select",vType,"Y")%>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">제 목 :</label>
					<input type="text" class="formTxt" name="title" value="<%=vTitle%>" style="width:200px" />
				</li>
				<li>
					<label class="formTit" for="term1">담당자 :</label>
					<% sbGetpartid "partmdid",vPartMDID,"","11,14,21,22,23" %>
				</li>
				<li>
					<label class="formTit" for="term1">WD :</label>
					<% sbGetpartid "partwdid",vPartWDID,"","12" %>
				</li>
				<li>
					<label class="formTit" for="term1">퍼블리셔 :</label>
					<select name="partpbid">
						<option value="">선택</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>최선미</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>조경애</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>진연미</option>
					</select>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>
	
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cPlay.FTotalCount%> 개 / idx, 시작일 기준 Sorting 되어있습니다.</p>
			</div>
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="javascript:goStyleCode('');"><span class="eIcon"><em class="fIcon">스타일코드관리</em></span></a></p>
				<p class="ftLt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
				<p class="btn2 cBk1 ftLt"><a href="javascript:goPlayType('');"><span class="eIcon"><em class="fIcon">분류관리</em></span></a></p>
				<p class="ftLt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
				<p class="btn2 cBk1 ftLt"><a href="javascript:goNewReg('');"><span class="eIcon"><em class="fIcon">신규등록</em></span></a></p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>idx</div></th>
					<th><div>No | 텍스트</div></th>
					<th><div>분류</div></th>
					<th><div>이미지</div></th>
					<th><div>제목</div></th>
					<th><div>사용여부</div></th>
					<th><div>우선순위번호</div></th>
					<th><div>조회수(인기순)</div></th>
					<th><div>시작일</div></th>
					<th><div>작업일자</div></th>
					<th><div>담당자</div></th>
					<th><div>작업자</div></th>
					<th><div>비고</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cPlay.FResultCount > 0 Then
						For i=0 To cPlay.FResultCount-1
				%>
						<tr>
							<td><%=cPlay.FItemList(i).Fidx%></td>
							<td><%=cPlay.FItemList(i).Fviewno%> | <%=cPlay.FItemList(i).Fviewnotxt%></td>
							<td><%=cPlay.FItemList(i).Ftypename%></td>
							<td><img src="<%=cPlay.FItemList(i).Flistimg%>" height="100" style="cursor:pointer;" onclick="jsImgView('<%=cPlay.FItemList(i).Flistimg%>');"></td>
							<td>
								<%= ReplaceBracket(cPlay.FItemList(i).Ftitle) %>
								<br /><br /><%=fnStateSelectBox("one",cPlay.FItemList(i).Fstate)%>
							</td>
							<td><%=cPlay.FItemList(i).Fisusing%></td>
							<td><%=cPlay.FItemList(i).Fsortno%></td>
							<td><%=cPlay.FItemList(i).Ffavcnt%></td>
							<td><%=cPlay.FItemList(i).Fstartdate%></td>
							<td>
								등록일 : <%=cPlay.FItemList(i).Fregdate%><br /><br />
								수정 : <%=cPlay.FItemList(i).Flastadminid%><br /><%=cPlay.FItemList(i).Flastupdate%></td>
							<td><%=cPlay.FItemList(i).FpartMDname%></td>
							<td>
								WD:<%=cPlay.FItemList(i).FpartWDname%><br />
								PB:<%=cPlay.FItemList(i).FpartPBname%>
							</td>
							<td>
								<input type="button" onClick="goNewReg('<%=cPlay.FItemList(i).Fidx%>');" value="수 정"><br /><br />
								<input type="button" onClick="jsTagview('<%=cPlay.FItemList(i).Fidx%>','<%=cPlay.FItemList(i).Ftype%>');" value="태 그"><br /><br />
								<input type="button" onClick="jsItem('<%=cPlay.FItemList(i).Fidx%>','<%=cPlay.FItemList(i).Ftype%>');" value="상 품">
							</td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<div class="ct tPad20 cBk1">
				<% if cPlay.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cPlay.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cPlay.StartScrollPage to cPlay.FScrollCount + cPlay.StartScrollPage - 1 %>
	    			<% if i>cPlay.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cPlay.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cPlay = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->