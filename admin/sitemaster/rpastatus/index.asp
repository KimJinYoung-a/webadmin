<% Option Explicit %>
<%
'###########################################################
' Description : rpa 성공 실패 리스트
' Hieditor : 2021.07.20 원승현 생성
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/rpastatus/rpastatuscls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
Dim loginUserId, i, currpage, pagesize, research, startdate, enddate
Dim rpatype, rpatitle, rpacontents, rparegdate, rpaissuccess
Dim oRpaStatusList

loginUserId = session("ssBctId") '// 로그인한 사용자 아이디
currpage = requestcheckvar(request("page"), 20) '// 현재 페이지 번호
rpatype = requestcheckvar(request("rpatype"), 240) '// 타입명(cls에 타입 정의 참고)
research = requestcheckvar(request("research"), 20) '// 재검색여부
startdate = requestcheckvar(request("startdate"), 20) '// 등록일 시작 검색값
enddate = requestcheckvar(request("enddate"), 20) '// 등록일 종료 검색값
rpaissuccess = requestcheckvar(request("rpaissuccess"), 20) '// 성공 실패 여부

If Trim(currpage)="" Then
	currpage = "1"
End If
pagesize = 30


'// 리스트를 가져온다.
set oRpaStatusList = new CgetRpaStatus
	oRpaStatusList.FRectcurrpage = currpage
	oRpaStatusList.FRectpagesize = pagesize
	If Trim(research)="on" Then
		oRpaStatusList.FRectType        = rpatype
		oRpaStatusList.FRectIsSuccess   = rpaissuccess
		oRpaStatusList.FRectStartdate   = startdate
		oRpaStatusList.FRectEnddate     = enddate
	End If
    oRpaStatusList.GetHalfDeliveryPayList()
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

function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&rpatype=<%=rpatype%>&startdate=<%=startdate%>&enddate=<%=enddate%>&rpaissuccess=<%=rpaissuccess%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchRpaStatus()
{
	document.frm1.submit();
}

function jsChkAll(){
var frm;
frm = document.frm;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkidx) !="undefined"){
	   	   if(!frm.chkidx.length){
		   	 	frm.chkidx.checked = true;
		   }else{
				for(i=0;i<frm.chkidx.length;i++){
					frm.chkidx[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkidx) !="undefined"){
	  	if(!frm.chkidx.length){
	   	 	frm.chkidx.checked = false;
	   	}else{
			for(i=0;i<frm.chkidx.length;i++){
				frm.chkidx[i].checked = false;
			}
		}
	  }

	}
}

function goIsUsingModifyAll(tp) {
	var itemcount = 0;
	var frm;
	var ck=0;
	frm = document.frm;
	if(typeof(frm.chkidx) !="undefined"){
		if(!frm.chkidx.length){
			if(!frm.chkidx.checked){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
			//frm.itemidarr.value = frm.chkitem.value;
			//frm.itemdataarr.value = frm.viewitemdata.value;
		}else{
			//frm.itemidarr.value = "";
			for(i=0;i<frm.chkidx.length;i++){
				if(frm.chkidx[i].checked) {
					ck=ck+1;
					if (frm.itemisusingarr.value==""){
						frm.itemisusingarr.value =  frm.chkidx[i].value;
					}else{
						frm.itemisusingarr.value = frm.itemisusingarr.value + "," +frm.chkidx[i].value;
					}
				}
			}

			if (frm.itemisusingarr.value == ""){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
		}
	}else{
		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
		return;
	}

	$("#isusingtype").val(tp);
	if(confirm("선택하신 모든 상품의 사용여부가 변경됩니다.\n수정하시겠습니까?")) {
		document.frm.submit();
	} else {
		return false;
	}
}

function jsEtcSaleMarginJungsan(makerid){
	var upfrm1 = document.frmEtcJOne;
    upfrm1.makerid.value=makerid;

    if (confirm("작성 하시겠습니까?")){
        upfrm1.submit();
    }
}

</script>
<div class="">
	<%' 상단 검색폼 시작 %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/rpastatus/index.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">업무명 :</label>
                    <select name="rpatype">
                        <option value="">선택</option>			
                        <option value="네이버페이" <%IF trim(rpatype)="네이버페이" THEN%>selected<%END IF%>>네이버페이 정산내역 다운로드</option>
                        <option value="이세로" <%IF trim(rpatype)="이세로" THEN%>selected<%END IF%>>이세로 전자계산서 다운로드</option>
                        <option value="KICC승인" <%IF trim(rpatype)="KICC승인" THEN%>selected<%END IF%>>KICC 승인내역 다운로드</option>
                        <option value="KICC입금" <%IF trim(rpatype)="KICC입금" THEN%>selected<%END IF%>>KICC 입금내역 다운로드</option>
                        <option value="제휴몰정산" <%IF trim(rpatype)="제휴몰정산" THEN%>selected<%END IF%>>제휴몰 정산내역 다운로드(몰별)</option>
                        <option value="제휴사송장" <%IF trim(rpatype)="제휴사송장" THEN%>selected<%END IF%>>제휴사 송장 검토 및 변경</option>
                        <option value="출고지시" <%IF trim(rpatype)="출고지시" THEN%>selected<%END IF%>>출고지시</option>
                        <option value="카카오기프트옵션" <%IF trim(rpatype)="카카오기프트옵션" THEN%>selected<%END IF%>>카카오 기프트 옵션 재고 매칭</option>
                        <option value="법인카드" <%IF trim(rpatype)="법인카드" THEN%>selected<%END IF%>>법인카드 SCM 업로드</option>
                        <option value="샤방넷" <%IF trim(rpatype)="샤방넷" THEN%>selected<%END IF%>>샤방넷 문의사항 수집</option>
                        <option value="제휴몰주문" <%IF trim(rpatype)="제휴몰주문" THEN%>selected<%END IF%>>제휴몰 주문 수집</option>
                        <option value="매출재고대사" <%IF trim(rpatype)="매출재고대사" THEN%>selected<%END IF%>>매출재고 대사작업</option>
                    </select>		
				</li>
				<li>
					<p class="formTit">기간</p>
					<input type="text" id="startdate" name="startdate" value="<%=startdate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "startdate", trigger    : "startdate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
                     ~ 
					<p class="formTit"></p>
					<input type="text" id="enddate" name="enddate" value="<%=enddate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "enddate", trigger    : "enddate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>                     
				</li>
				<li>
					<p class="formTit">성공여부 :</p>
					<select class="formSlt" id="rpaissuccess" name="rpaissuccess" title="성공여부 선택">
						<option value="" <% If rpaissuccess = "" Then %> selected <% End If %>>전체</option>
						<option value="1" <% If rpaissuccess = "1" Then %> selected <% End If %>>성공</option>
						<option value="0" <% If rpaissuccess = "0" Then %> selected <% End If %>>실패</option>
					</select>
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onclick="goSearchRpaStatus();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">총 등록수 : <strong><%=FormatNumber(oRpaStatusList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:80px">번호(idx)</p>
							<p style="width:100px">업무명</p>
                            <p style="width:450px">제목</p>
							<!--p style="width:600px">내용</p-->
                            <p style="width:80px">성공여부</p>
							<p style="width:90px">등록일</p>
							<p style="width:150px"></p>
						</li>
					</ul>
					<ul id="sortable" class="tbDataList">
						<% If oRpaStatusList.FResultcount > 0 Then %>
							<% For i=0 To oRpaStatusList.Fresultcount-1 %>
							<% If oRpaStatusList.FrpaStatusList(i).FisSuccess = 0 Then %>
								<li style="background-color:#FFEDED">
							<% Else %>
								<li style="background-color:#F7FFE6">
							<% End If %>
								<p style="width:80px"><%=oRpaStatusList.FrpaStatusList(i).Fidx%></p>
								<p style="width:100px"><%=getRpaTypeName(oRpaStatusList.FrpaStatusList(i).Ftype)%></p>
								<p style="width:450px" align="left"><%=oRpaStatusList.FrpaStatusList(i).Ftitle%></p>
								<!--p style="text-align:left;width:600px;white-space:pre-line;"><%'replace(oRpaStatusList.FrpaStatusList(i).Fcontents,chr(13)&chr(10),"<br>")%></p-->
								<p style="width:80px"><%=getRpaIsSuccessName(oRpaStatusList.FrpaStatusList(i).FisSuccess)%></p>
								<p style="width:90px"><%=oRpaStatusList.FrpaStatusList(i).Fregdate%></p>
								<p style="width:150px"><button onclick="window.open('popviewrpastatus.asp?idx=<%=oRpaStatusList.FrpaStatusList(i).Fidx%>',null,'height=800,width=1000,status=yes,toolbar=no,menubar=no,location=no');return false;">내용확인</button></p>
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%=fnDisplayPaging_New2017(currpage, oRpaStatusList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
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
});
</script>

</body>
</html>
<%
	Set oRpaStatusList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
