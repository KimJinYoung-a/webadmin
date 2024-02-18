<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/index.asp
' Description :  딜 이벤트 관리
' History : 2017.08.22 정태훈
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->


<%
	'변수선언
	Dim iCurrpage, iPageSize, iPerCnt, isResearch, sSdate, sEdate, intLoop, stext, dispCate
	Dim oDeal, arrList, iTotCnt, iTotalPage, strTxt, sdiv, datediv, viewdiv, isusing, arrCate, maxDepth

	dispCate	= requestCheckVar(Request("disp"),16) 		'전시 카테고리
	maxDepth = 2
	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

	isusing 		= requestCheckVar(Request("isusing"),1)
	viewdiv 		= requestCheckVar(Request("viewdiv"),1)
	datediv 		= requestCheckVar(Request("datediv"),1)
	sdiv 		= requestCheckVar(Request("sdiv"),10)
	strTxt 		= requestCheckVar(Request("stext"),32)
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## 검색 #############################
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	'데이터 가져오기
	set oDeal = new ClsDeal
		oDeal.FCPage = iCurrpage		'현재페이지
		oDeal.FPSize = iPageSize		'한페이지에 보이는 레코드갯수
		oDeal.FSearchDateDiv 	= datediv	'검색일 구분
		oDeal.FSsDate 	= sSdate	'검색 시작일
		oDeal.FSeDate 	= sEdate	'검색 종료일
		oDeal.FSearchDiv 	= sdiv	'검색구분
		oDeal.FSeTxt 	= strTxt	'검색어
		oDeal.FSViewDiv 	= viewdiv	'유형 구분
		oDeal.FSIsUsing 	= isusing	'사용 구분
		oDeal.FSdispCate 	= dispCate	'전시카테고리 검색
 		arrList = oDeal.fnGetDealList	'데이터목록 가져오기
 		iTotCnt = oDeal.FTotCnt	'전체 데이터  수
 	set oDeal = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<!--
	window.document.domain = "10x10.co.kr";
	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}
		if(frm.sdiv.value=="itemid" && frm.stext.value!=""){
			if(isNaN(frm.stext.value)){
				alert("상품번호 검색은 숫자만 입력해주세요!");
				return false;
			}
		}

		frm.submit();
	}
	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}
	function TnEditDeal(url){
		location.href=url;
	}

	//미리보기
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes,location=yes");
	    }
	}

	function fnDealInfoUpdate(){
		$.ajax({
			type: "POST",
			url: "ajaxDealInfoUpdate.asp",
			data: "mode=all",
			cache: false,
			async: false,
			success: function(message) {
				if(message=="OK") {
					alert("업데이트 완료.");
				} else {
					alert("제공 할 정보가 없습니다.");
				}
			}
		});
	}

	function fnDealItemInfoUpdate(itemid){
		$.ajax({
			type: "POST",
			url: "ajaxDealInfoUpdate.asp",
			data: "mode=one&itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				if(message=="OK") {
					alert("업데이트 완료.");
				} else {
					alert("제공 할 정보가 없습니다.");
				}
			}
		});
	}

    function TnDevDealSaveAPICall(itemid){
		$.ajax({
			type: "POST",
			url: "<%= ItemUploadUrl %>/linkweb/items/deal_itemregisterTempWithImage_process.asp",
			data: "itemid=" + itemid,
			dataType: "JSON",
			cache: false,
			success: function(data) {
				alert(data.message);
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
    }
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">검색 조건</td>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table>
		<tr>
			<td>
				기간:
				<select name="datediv">
					<option value="S"<% If datediv="S" Then Response.write " selected" %>>시작일</option>
					<option value="E"<% If datediv="E" Then Response.write " selected" %>>종료일</option>
					<option value="R"<% If datediv="R" Then Response.write " selected" %>>작성일</option>
				</select>
				<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<tr>
			<td>
				유형 : 
				<select name="viewdiv" class="select">
					<option value="" selected>전체</option>
					<option value="1"<% If viewdiv="1" Then Response.write " selected" %>>상시딜</option>
					<option value="2"<% If viewdiv="2" Then Response.write " selected" %>>기간딜</option>
				</select>
				사용여부 : 
				<select name="isusing" class="select">
					<option value="" selected>전체</option>
					<option value="Y"<% If isusing="Y" Then Response.write " selected" %>>사용</option>
					<option value="N"<% If isusing="N" Then Response.write " selected" %>>사용안함</option>
				</select>
				전시 카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			</td>
		</tr>
		<tr>
			<td>
				검색어 : 
				<select name="sdiv" class="select">
					<option value="itemid"<% If sdiv="itemid" Then Response.write " selected" %>>딜상품코드</option>
					<option value="itemname"<% If sdiv="itemname" Then Response.write " selected" %>>상품명</option>
					<option value="register"<% If sdiv="register" Then Response.write " selected" %>>작성자</option>
					<option value="makerid"<% If sdiv="makerid" Then Response.write " selected" %>>브랜드아이디</option>
				</select>
				<input type="text" name="stext" size="50" value="<%=strTxt%>" onkeydown="if(event.keyCode==13) jsSearch('E');">
			</td>
		</tr>
		</table>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" align="center"><input type="button" class="button_s" value="검색" onClick="javascript:jsSearch('E');"></td>
</tr>
</form>
</table><br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">
			<table width="100%">
			<tr>
				<td>검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
				<td align="right"><% if session("ssBctId")="seojb1983" or C_ADMIN_AUTH then %><input type="button" class="button" style="width:105;" value="정보업데이트" onclick="fnDealInfoUpdate();">&nbsp;&nbsp;&nbsp;<% end if %><input type="button" class="button" style="width:105;" value="딜 매출통계" onclick="jsGoUrl('/admin/dataanalysis/report/weeklysimplereport.asp?menupos=4019&reporttype=dealsales');">&nbsp;&nbsp;&nbsp;<input type="button" class="button" style="width:105;" value="등록" onclick="jsGoUrl('/admin/itemmaster/deal/new_deal_reg.asp?menupos=<%=menupos%>');"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>딜코드</td>
		<td>딜상품코드</td>
		<td>카테고리</td>
		<td>유형</td>
		<td>노출기간</td>
		<td>사용여부</td>
		<td>딜상품명</td>
		<td>텐바이텐가</td>
		<td>할인율</td>
		<td>작성자</td>
		<td>작성일</td>
		<td>미리보기</td>
		<td>관리</td>
	 </tr>
	 <% If isArray(arrList) Then %>
	 <% For intLoop = 0 To UBound(arrList,2) %>
	 <% if arrList(14,intLoop)="0" then %>
	 <tr bgcolor="#EEEEEE">
	 <% else %>
	 <tr bgcolor="#FFFFFF">
	 <% End If %>
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center">
		<%
			If arrList(12,intLoop) <> "" Then
			arrCate = Split(arrList(12,intLoop),"^^")
			If ubound(arrCate)>0 Then
			Response.write arrCate(0) & " > " & arrCate(1)
			Else
			Response.write arrCate(0)
			End If
			End If
		%>
		</td>
		<td align="center"><% If arrList(2,intLoop)="1" Then %>상시딜<% Else %>기간딜<% End If %></td>
		<td align="center"><% If arrList(2,intLoop)<>"2" Then %>상시 노출<% Else %><%=FormatDateTime(arrList(3,intLoop),2)%> ~ <%=FormatDateTime(arrList(4,intLoop),2)%> <% End If %></td>
		<td align="center"><% if arrList(14,intLoop)="0" then %>등록대기<% Else %><% If arrList(13,intLoop) = "Y" Then %>사용<% Else %>사용안함<% End If %><% End If %></td>
		<td><%=arrList(5,intLoop)%></td>
		<td align="right"><%=FormatNumber(arrList(6,intLoop),0)%>원<% If arrList(10,intLoop) = "Y" Then %>~<% Else %>&nbsp;&nbsp;<% End If %></td>
		<td align="center"><% If arrList(11,intLoop) = "Y" Then %>~<% End If %><%=arrList(7,intLoop)%>%</td>
		<td align="center"><%=arrList(8,intLoop)%></td>
		<td align="center"><%=left(arrList(9,intLoop),10)%></td>
		<td align="center"><a href="<%=vwwwUrl%>/deal/deal.asp?itemid=<%=arrList(1,intLoop)%>" target="_blank"><img src="/images/iexplorer.gif" border="0"></a>&nbsp;<a href="javascript:jsOpen('<%=vmobileUrl%>/deal/deal.asp?itemid=<%=arrList(1,intLoop)%>','M');"><img src="/images/iexplorer.gif" border="0"></a></td>
		<td align="center">
			<% if arrList(14,intLoop)="0" and (application("Svr_Info")="Dev") then %>
			<input type="button" class="button" style="width:105;" value="등록컨펌" onclick="TnDevDealSaveAPICall(<%= arrList(1,intLoop) %>);">
			<% End If %>
			<input type="button" class="button" style="width:105;" value="수정" onclick="TnEditDeal('/admin/itemmaster/deal/new_deal_edit.asp?idx=<%= arrList(0,intLoop) %>');"<% if arrList(14,intLoop)="0" then %> disabled<% End If %>>
			<input type="button" class="button" style="width:105;" value="정보업데이트" onclick="fnDealItemInfoUpdate(<%= arrList(1,intLoop) %>);">
		</td>
	 </tr>
	 <% Next %>
	 <% Else %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" align="center" height="25">
			등록된 내용이 없습니다.
		</td>
	 </tr>
	 <% End If %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" bgcolor="#FFFFFF" align="center">
			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	 </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->