<% Option Explicit %>
<%
'###########################################################
' Description : piece 터미널(리스트) 페이지
' Hieditor : 2017.08.28 원승현 생성
'			 2017-11-29 이종화 추가 / 수정
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/piece/piececls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->

<%
Dim gubun	'1 : 조각, 2 :파이, 3 : 베스트키워드, 4: 배너, 5:회원조각
Dim oPieceUser, loginUserId, oPieceList, i, oPieceOpening, currpage, pagesize, deal, open, research, keyword, schWord
Dim state , schdate

loginUserId = session("ssBctId")
currpage = requestcheckvar(request("page"), 20)
deal = requestcheckvar(request("deal"), 20)
open = requestcheckvar(request("open"), 20)
research = requestcheckvar(request("research"), 20)
keyword = requestcheckvar(request("keyword"), 20)
schWord = requestcheckvar(request("schWord"), 500)
state = requestcheckvar(request("state"), 1)

'// 시작일 검색
schdate = requestcheckvar(request("prevDate"), 10)

If keyword = "snum" And Not(isNumeric(schWord)) Then
	Response.write "<script>alert('번호(idx) 를 확인 해주세요');</script>"
	schWord = ""
End If

If Trim(currpage)="" Then
	currpage = "1"
End If
pagesize = 30

'// 현재 들어온 관리자가 piece에 등록된 관리자인지 확인한다.
set oPieceUser = new Cgetpiece
	oPieceUser.FRectadminid = loginUserId
	oPieceUser.adminPieceUser()

'// 오프닝 데이터를 가져온다.
set oPieceOpening = new Cgetpiece
	oPieceOpening.getPieceOpening()

'// 리스트를 가져온다.
set oPieceList = new Cgetpiece
	oPieceList.FRectcurrpage = currpage
	oPieceList.FRectpagesize = pagesize
	If Trim(research)="on" Then
		oPieceList.FRectDeal = deal
		oPieceList.FRectOpen = open
		oPieceList.FRectkeyword = keyword
		oPieceList.FRectSchword = schWord
		oPieceList.FRectState = state
		oPieceList.FRectStartdate = schdate
	End If
	oPieceList.GetpieceList()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>
document.domain = "10x10.co.kr";

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

// 관리자 입력/수정
function fnPieceUseract()
{
	// nickname값을 넣어줌.
	$("#frmnickname").val(escape($("#usernickname").val()));
	// occupation값을 넣어줌
	$("#frmoccupation").val($("#selectOccupation").val());

	if ($("#selectOccupation").val()=="A")
	{
		alert("직종을 선택해주세요.");
		return;
	}
	if ($("#usernickname").val()=="")
	{
		alert("닉네임을 입력해주세요.");
		return;
	}

	<% if trim(oPieceUser.FoneUser.Fnickname)<>"" then %>
		$("#frmmode").val("upd");
	<% else %>
		$("#frmmode").val("ins");
	<% end if %>

	$.ajax({
		type:"GET",
		url:"/admin/sitemaster/piece/act_pieceUser.asp",
		data:$("#frmpieceUser").serialize(),
		async:false,
		cache:true,
		success : function(Data, textStatus, jqXHR){
			if (jqXHR.readyState == 4) {
				if (jqXHR.status == 200) {
					if(Data!="") {
						res = Data.split("||");
						if (res[0]=="OK")
						{
							if (res[1]=="2")
							{
								alert("수정완료");
							}
							document.location.reload();
						}
						else
						{
							errorMsg = res[1].replace(">?n", "\n");
							alert(errorMsg);
							document.location.reload();
							return false;
						}
					} else {
						alert("잘못된 접근 입니다.");
						document.location.reload();
						return false;
					}
				}
			}
		},
		error:function(jqXHR, textStatus, errorThrown){
			alert("잘못된 접근 입니다.");
			<% if false then %>
				//var str;
				//for(var i in jqXHR)
				//{
				//	 if(jqXHR.hasOwnProperty(i))
				//	{
				//		str += jqXHR[i];
				//	}
				//}
				//alert(str);
			<% end if %>
			document.location.reload();
			return false;
		}
	});
}

function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&deal=<%=deal%>&open=<%=open%>&keyword=<%=keyword%>&schWord=<%=schWord%>&state=<%=state%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchPiece()
{
//	if ($("#deal").val()=="0"&&$("#open").val()=="A"&&$("#schWord").val()=="")
//	{
//		alert("검색을 하기 위해선 구분, 노출여부, 키워드검색 중 하나를\n선택해주시거나 입력해주셔야 합니다.");
//		return;
//	}
	document.frm1.submit();
}

function fnPieceDelact(idx, gubun)
{
	$("#frmDelidx").val(idx);

	var result

	if (gubun=="1")
	{
		result = confirm("조각을 삭제하시겠습니까?");
	}
	if (gubun=="2")
	{
		result = confirm("파이를 삭제하시겠습니까?");
	}
	if (gubun=="3")
	{
		result = confirm("베스트키워드를 삭제하시겠습니까?");
	}
	if (gubun=="4")
	{
		result = confirm("배너를 삭제하시겠습니까?");
	}
	if (gubun=="5")
	{
		result = confirm("회원조각을 삭제하시겠습니까?");
	}

	if (result)
	{
		$.ajax({
			type:"GET",
			url:"/admin/sitemaster/piece/act_pieceDelete.asp",
			data:$("#frmpieceDel").serialize(),
			async:false,
			cache:true,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							res = Data.split("||");
							if (res[0]=="OK")
							{
								document.location.reload();
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg);
								document.location.reload();
								return false;
							}
						} else {
							alert("잘못된 접근 입니다.");
							document.location.reload();
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("잘못된 접근 입니다.");
				<% if false then %>
					//var str;
					//for(var i in jqXHR)
					//{
					//	 if(jqXHR.hasOwnProperty(i))
					//	{
					//		str += jqXHR[i];
					//	}
					//}
					//alert(str);
				<% end if %>
				document.location.reload();
				return false;
			}
		});
	}
	else
	{
		return;
	}
}

</script>
<div class="">
	<div class="tab" style="margin:0 0 0 -1px;">
		<ul>
			<li class="col11 selected"><a href="#unitType01">공간관리</a></li>
			<li class="col11 "><a href="#unitType02">수집된 조각</a></li>
		</ul>
		<div class="managerInfo">
			<p><%=oPieceUser.FoneUser.Foccupation%> <strong><%=oPieceUser.FoneUser.Fnickname%></strong> <button type="button" class="memEdit">변경</button></p>
			<p style="min-width:80px;">나의 조각 <strong><%=pieceMyCnt(loginUserId)%></strong></p>
		</div>
	</div>

	<%' 상단 검색폼 시작 %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/piece/piece_terminal.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">구분 :</label>
					<select class="formSlt" id="deal" name="deal" title="옵션 선택">
						<option value="0" <% If deal = "" Or deal = "0" Then %> selected <% End If %>>전체</option>
						<option value="1" <% If deal = "1" Then %> selected <% End If %>>조각</option>
						<option value="2" <% If deal = "2" Then %> selected <% End If %>>파이</option>
						<option value="3" <% If deal = "3" Then %> selected <% End If %>>베스트 키워드</option>
						<option value="4" <% If deal = "4" Then %> selected <% End If %>>배너</option>
						<option value="5" <% If deal = "5" Then %> selected <% End If %>>회원 조각</option>
					</select>
				</li>
				<li>
					<p class="formTit">노출여부 :</p>
					<select class="formSlt" id="open" name="open" title="옵션 선택">
						<option value="A" <% If open = "" Or open = "A" Then %> selected <% End If %>>전체</option>
						<option value="Y" <% If open = "Y" Then %> selected <% End If %>>공개</option>
						<option value="N" <% If open = "N" Then %> selected <% End If %>>비공개</option>
					</select>
				</li>
				<li>
					<p class="formTit">진행상태</p>
					<select class="formSlt" id="state" name="state" title="옵션 선택">
						<option value="" <% If state = ""  Then %> selected <% End If %>>전체</option>
						<option value="1" <% If state = "1" Then %> selected <% End If %>>등록대기</option>
						<option value="2" <% If state = "2" Then %> selected <% End If %>>이미지 등록요청</option>
						<option value="3" <% If state = "3" Then %> selected <% End If %>>디자인 작업중</option>
						<option value="4" <% If state = "4" Then %> selected <% End If %>>오픈요청</option>
						<option value="7" <% If state = "7" Then %> selected <% End If %>>오픈</option>
						<option value="8" <% If state = "8" Then %> selected <% End If %>>보류</option>
						<option value="9" <% If state = "9" Then %> selected <% End If %>>종료</option>
					</select>
				</li>
				<li>
					<p class="formTit">시작일</p>
					<input type="text" id="prevDate" name="prevDate" value="<%=schdate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "prevDate", trigger    : "prevDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">키워드 검색 :</label>
					<select class="formSlt" id="keyword" name="keyword" title="키워드 검색">
						<option value="snum" <% If keyword = "snum" Then %> selected <% End If %>>번호</option>
						<option value="stitle" <% If keyword ="" Or keyword = "stitle" Then %>selected<% End If %>>구분제목</option>
						<option value="sname" <% If keyword = "sname" Then %> selected <% End If %>>작성자</option>
					</select>
					<input type="text" class="formTxt" id="schWord" name="schWord" style="width:400px" placeholder="키워드를 입력하여 검색하세요." value="<%=schWord%>" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onclick="goSearchPiece();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<!-- 20170824 보류
					<input type="button" class="btnOdrChg btn cBl1 fs12" value="순서변경" />
					-->
					<input type="button" class="btnRegist btn bold fs12" value="등록" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp',null,'height=800,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" />
					<% If Trim(research)="on" Then %>
						<input type="button" class="btnRegist btn bold fs12" value="검색초기화" onclick="document.location.href='/admin/sitemaster/piece/piece_terminal.asp';" />
					<% End If %>
				</div>
				<!-- 20170824 보류
				<div class="ftLt">
					<p class="infoTxt">
						<span><img src='/images/ico_odrchg.png' alt='순서변경' /> 를 길게 눌러 위, 아래로 이동 후 변경완료 버튼을 클릭해주세요.</span>
						!-- for dev msg:검색조건 적용 후 순서변경 버튼 클릭시 노출됩니다. <span>검색조건 적용 시 순서를 변경할 수 없습니다. <button type="button">검색 초기화</button></span> --
					</p>
				</div>
				-->
			</div>

			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">총 등록수 : <strong><%=FormatNumber(oPieceList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:80px">번호(idx)</p>
							<p style="width:100px">구분</p>
							<p class="">구분제목</p>
							<p style="width:50px">공유</p>
							<p style="width:150px">작성자<br/><span class="cRd1">최종수정자</span></p>
							<p style="width:65px">노출여부</p>
							<p style="width:120px">등록일</p>
							<p style="width:120px">최종수정일</p>
							<p style="width:120px">시작일</p>
							<p style="width:65px">진행상태</p>
							<p style="width:120px">수정/삭제</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<% If Not(Trim(research)="on") Then %>
							<%'// 오프닝 데이터를 먼저 불러온다. %>
							<% If oPieceOpening.FResultCount > 0 Then %>
								<%'' for dev msg : 오프닝으로 선택된 항목은 li에 class="ui-state-disabled" 적용해주세요 %>
								<li class="ui-state-disabled">
									<p style="width:80px">고정<br/><a href="" onclick="window.open('http://m.10x10.co.kr/piece/piece_preview.asp?idx=<%=oPieceOpening.FOneOpening.FIdx%>',null,'height=720,width=375,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes');return false;" class="cBl1 tLine">[미리보기]</a></p>
									<p style="width:100px">오프닝</p>
									<p class="lt">
										<% If oPieceOpening.FOneOpening.Fgubun="1" Then %>
											<%=chrbyte(oPieceOpening.FOneOpening.Flisttext,75,"Y")%>
										<% Else %>
											<%=oPieceOpening.FOneOpening.Flisttitle%>
										<% End If %>
									</p>
									<p style="width:50px"><%=FormatNumber(oPieceOpening.FOneOpening.Fsnsbtncnt, 0)%></p>
									<p style="width:150px"><%=oPieceOpening.FOneOpening.Foccupation&" "&oPieceOpening.FOneOpening.Fnickname%><br/><span class="cRd1"><%=oPieceOpening.FOneOpening.Flastoccupation&" "&oPieceOpening.FOneOpening.Flastnickname%></span></p>
									<p style="width:65px">
										<% If oPieceOpening.FOneOpening.Fisusing="Y" Then %>
											공개
										<% Else %>
											비공개
										<% End If %>
									</p>
									<p style="width:120px"><%=Mid(Trim(oPieceOpening.FOneOpening.Fregdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Fregdate), 11, 30)%></p>
									<p style="width:120px" class="cRd1"><%=Mid(Trim(oPieceOpening.FOneOpening.Flastupdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Flastupdate), 11, 30)%></p>
									<p style="width:120px"><%=Mid(Trim(oPieceOpening.FOneOpening.Fstartdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Fstartdate), 11, 30)%></p>
									<p style="width:65px;"><%=nowstatus(oPieceOpening.FOneOpening.Fstate)%></p>
									<p style="width:120px">
										<a href="" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp?idx=<%=oPieceOpening.FOneOpening.FIdx%>&page=<%=currpage%>&SearchDeal=<%=deal%>&SearchOpen=<%=open%>&SearchState=<%=state%>',null,'height=900,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" class="cBl1 tLine">[수정]</a>
										<a href="" onclick="fnPieceDelact('<%=oPieceOpening.FOneOpening.FIdx%>', '<%=Trim(oPieceOpening.FOneOpening.Fgubun)%>');return false;" class="cBl1 tLine">[삭제]</a>
									</p>
								</li>
							<% End If %>
						<% End If %>

						<%'// 오프닝 데이터를 제외한 리스트를 가져온다. %>
						<% If oPieceList.FResultcount > 0 Then %>
							<% For i=0 To oPieceList.Fresultcount-1 %>
							<li>
								<!--p style="width:80px"><%=(oPieceList.FtotalCount - pagesize * (currpage-1) - i)%></p-->
								<p style="width:80px"><%=oPieceList.FPieceList(i).FIdx%><br/><a href="" onclick="window.open('http://m.10x10.co.kr/piece/piece_preview.asp?idx=<%=oPieceList.FPieceList(i).FIdx%>',null,'height=720,width=375,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes');return false;" class="cBl1 tLine">[미리보기]</a></p>
								<p style="width:100px">
									<% Select Case Trim(oPieceList.FPieceList(i).Fgubun) %>
										<% Case "1" %>
											조각
										<% Case "2" %>
											파이
										<% Case "3" %>
											베스트키워드
										<% Case "4" %>
											배너
										<% Case "5" %>
											회원조각
									<% End Select %>
								</p>
								<p class="lt">
									<% If oPieceList.FPieceList(i).Fgubun="1" Then %>
										<%=chrbyte(oPieceList.FPieceList(i).Flisttext,75,"Y")%>
									<% Else %>
										<%=oPieceList.FPieceList(i).Flisttitle%>
									<% End If %>
								</p>
								<p style="width:50px"><%=oPieceList.FPieceList(i).Fsnsbtncnt%></p>
								<p style="width:150px"><%=oPieceList.FPieceList(i).Foccupation&" "&oPieceList.FPieceList(i).Fnickname%><br/><span class="cRd1"><%=oPieceList.FPieceList(i).Flastoccupation&" "&oPieceList.FPieceList(i).Flastnickname%></span></p>
								<p style="width:65px">
								<% If Trim(oPieceList.FPieceList(i).Fisusing)="Y" Then %>
									공개
								<% Else %>
									비공개
								<% End If %>
								</p>
								<p style="width:120px"><%=Mid(Trim(oPieceList.FPieceList(i).Fregdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Fregdate), 11, 30)%></p>
								<p style="width:120px" class="cRd1"><% If oPieceList.FPieceList(i).Flastnickname <> "" Then %><%=Mid(Trim(oPieceList.FPieceList(i).Flastupdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Flastupdate), 11, 30)%><% End If %></p>
								<p style="width:120px"><%=Mid(Trim(oPieceList.FPieceList(i).Fstartdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Fstartdate), 11, 30)%></p>
								<p style="width:65px"><%=nowstatus(oPieceList.FPieceList(i).Fstate)%></p>
								<p style="width:120px">
									<a href="" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp?idx=<%=oPieceList.FPieceList(i).FIdx%>&SearchDeal=<%=deal%>&SearchOpen=<%=open%>&SearchState=<%=state%>',null,'height=900,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" class="cBl1 tLine">[수정]</a>
									<a href="" onclick="fnPieceDelact('<%=oPieceList.FPieceList(i).FIdx%>', '<%=Trim(oPieceList.FPieceList(i).Fgubun)%>');return false;" class="cBl1 tLine">[삭제]</a>
								</p>
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%= fnDisplayPaging_New2017(currpage, oPieceList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="lyrBox">
	<div class="pieceMember">
		<strong>당신은 어떤 사람인가요?</strong>
		<p class="tPad10">당신의 직종와 Piece에서 사용할 별명을 입력해주세요.</p>
		<div class="whoAreYou">
			<p class="ftLt">
				<select class="formSlt" style="width:100px; height:30px;" name="selectOccupation" id="selectOccupation">
					<option value="A">직종선택</option>
					<option value="Member" <% If oPieceUser.FoneUser.Foccupation="Member" Then %>selected<% End If %>>Member</option>
					<option value="Planner" <% If oPieceUser.FoneUser.Foccupation="Planner" Then %>selected<% End If %>>Planner</option>
					<option value="Designer" <% If oPieceUser.FoneUser.Foccupation="Designer" Then %>selected<% End If %>>Designer</option>
					<option value="Publisher" <% If oPieceUser.FoneUser.Foccupation="Publisher" Then %>selected<% End If %>>Publisher</option>
					<option value="Developer" <% If oPieceUser.FoneUser.Foccupation="Developer" Then %>selected<% End If %>>Developer</option>
					<option value="MD" <% If oPieceUser.FoneUser.Foccupation="MD" Then %>selected<% End If %>>MD</option>
					<option value="Editor" <% If oPieceUser.FoneUser.Foccupation="Editor" Then %>selected<% End If %>>Editor</option>
				</select>
			</p>
			<p class="ftRt">
				<input type="text" placeholder="닉네임" class="formTxt" style="height:30px;" name="usernickname" id="usernickname" value="<%=oPieceUser.FoneUser.Fnickname%>" />
			</p>
		</div>
		<p>
			<input type="button" value="확인" class="cRd1" style="width:100px; height:30px;" onclick="fnPieceUseract();return false;" />
		</p>
	</div>
</div>
<form name="frmpieceUser" id="frmpieceUser">
	<input type="hidden" name="frmmode" id="frmmode">
	<input type="hidden" name="frmoccupation" id="frmoccupation">
	<input type="hidden" name="frmnickname" id="frmnickname">
	<input type="hidden" name="frmadminid" id="frmadminid" value="<%=loginUserId%>">
	<input type="hidden" name="frmidx" id="frmidx" value="<%=oPieceUser.FoneUser.Fidx%>">
</form>
<form name="frmpieceDel" id="frmpieceDel">
	<input type="hidden" name="frmDeladminid" id="frmDeladminid" value="<%=loginUserId%>">
	<input type="hidden" name="frmDelidx" id="frmDelidx">
</form>
<div class="dimmed"></div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$(".btnOdrChg").on('click',function() {
		if ($("#sortable").hasClass('sortable')) {
			$("#sortable").removeClass('sortable');
			$("#sortable li p:first-child").html("901"); //리스트 index값 들어가게끔
			$("#sortable li.ui-state-disabled p:first-child").html("고정");
			$("#sortable").sortable("destroy");
			$(".btnOdrChg").attr("value", "순서변경");
			//$(".btnOdrChg").prop("disabled", true); //검색조건 적용시 순서변경 버튼 비활성화
			$(".btnRegist").prop("disabled", false);
			$(".infoTxt").hide();
		} else {
			$("#sortable").addClass('sortable');
			$("#sortable li p:first-child").html("<img src='/images/ico_odrchg.png' alt='순서변경' />");
			$("#sortable li.ui-state-disabled p:first-child").html("고정");
			$("#sortable").sortable({
				placeholder:"handling",
				items:"li:not(.ui-state-disabled)"
			}).disableSelection();
			$(".btnOdrChg").attr("value", "변경완료");
			//$(".btnOdrChg").prop("disabled", false);
			$(".btnRegist").prop("disabled", true);
			$(".infoTxt").show();
		}
	});

	$(".memEdit").on('click',function() {
		$(".dimmed").show();
		$(".lyrBox").show();
	});

	<% if oPieceUser.FResultCount < 1 then %>
		$(".dimmed").show();
		$(".lyrBox").show();
	<% end if %>

});
</script>

</body>
</html>
<%
	Set oPieceUser = Nothing
	Set oPieceList = Nothing
	Set oPieceOpening = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
