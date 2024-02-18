<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 고객센터
' Hieditor : 2015.05.27 이상구 생성
'			 2017.07.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim orderserial, oaslist, totalascount
	orderserial = RequestCheckVar(request("orderserial"),11)

totalascount = 0

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

set oaslist = new CCSASList
if (orderserial <> "") then
    oaslist.FRectOrderSerial = orderserial
    oaslist.GetCSASTotalCount

    totalascount = oaslist.FResultCount
end if

dim ix

if (orderserial<>"") then
    if ojumun.FOneItem.IsForeignDeliver then
        ojumun.getEmsOrderInfo
    end if
end if

dim olecjumun
set olecjumun = new CRequestLecture
	olecjumun.FRectOrderSerial = orderserial

	if (orderserial <> "") then
		olecjumun.GetRequestLectureMasterOne
	end if

dim ojumundetail
set ojumundetail = new CRequestLecture

ojumundetail.FRectOrderSerial = orderserial

if (orderserial <> "") then
	if (ojumun.FOneItem.Fsitename <> "diyitem") then
		ojumundetail.CRequestLectureDetailList
	else
		ojumundetail.CRequestDIYItemDetailList
	end if
end if

dim olecture
set olecture = new CLecture

if (orderserial <> "") then
	olecture.FRectIdx = olecjumun.FOneItem.Fitemid
end if

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if

if (orderserial <> "") then
	olecture.GetOneLecture
end if

%>
<script type="text/javascript">

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_mileage(v){
	var popwin = window.open("/cscenter/mileage/cs_mileage.asp?userid=" + v,"cs_mileage","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_coupon(v){
	var popwin = window.open("/cscenter/coupon/cs_coupon.asp?userid=" + v,"cs_coupon","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function pop_cs_register(v){
	// var popwin = window.showModalDialog("/cscenter/action/pop_cs_register.asp?orderserial=" + v,"misendmaster","resizable:yes; scroll:yes; dialogWidth:825px; dialogHeight:800px ");
	var popwin = window.open("/cscenter/action/pop_cs_register.asp?orderserial=" + v,"misendmaster","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function order_receiver_info(v){
	var popwin = window.showModalDialog("/cscenter/ordermaster/order_receiver_info.asp?orderserial=" + v,"order_reciever_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:480px");
	popwin.focus();
}

function order_buyer_info(v){
	var popwin = window.showModalDialog("order_buyer_info.asp?orderserial=" + v,"order_buyer_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:270px");
	popwin.focus();
}

// ============================================================================
// CS등록관련

//카드취소
function PopOpenCancelCard(orderserial){
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_repay.asp?divcd=7&orderserial=" + orderserial,"PopOpenCancelCard","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

//외부몰취소요청
function PopOpenCancelOtherSite(orderserial){
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_repay.asp?divcd=5&orderserial=" + orderserial,"PopOpenCancelOtherSite","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

// ============================================================================
// 히스토리 관련
var selected_history_menu = "";

function ChangeWriteButton(menuname) {
        selected_history_menu = menuname;

        if (menuname == "memo") {
                document.frmhistory.writebutton.value = "MEMO작성";
        } else if (menuname == "cs") {
                document.frmhistory.writebutton.value = "CS리스트";
        } else if (menuname == "mileage") {
                document.frmhistory.writebutton.value = "마일리지관리";
        } else if (menuname == "coupon") {
                document.frmhistory.writebutton.value = "쿠폰관리";
        } else if (menuname == "qna") {
                document.frmhistory.writebutton.value = "1:1상담관리";
        }
}

function OpenHistoryWindow(userid, orderserial) {
        if (selected_history_menu == "memo") {
                GotoHistoryMemoWrite(userid, orderserial);
        } else if (selected_history_menu == "cs") {
                Cscenter_Action_List_Lecture(orderserial,'','')
        } else if (selected_history_menu == "mileage") {
                cs_mileage(userid)
        } else if (selected_history_menu == "coupon") {
                cs_coupon(userid)
        }
}

function GotoHistoryMemo(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("memo");

                document.history.location.href = "/cscenterv2/history/history_memo.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCS(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("cs");

                document.history.location.href = "/cscenterv2/history/history_lec_cs.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryMileage(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("mileage");

                document.history.location.href = "/cscenter/history/history_mileage.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryCoupon(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("coupon");

                document.history.location.href = "/cscenter/history/history_coupon.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}
function GotoHistoryQna(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
                ChangeWriteButton("qna");

                document.history.location.href = "/cscenter/history/history_qna.asp?userid=" + userid + "&orderserial=" + orderserial;
        }
}

function GotoHistoryMemoWrite(userid, orderserial) {
	var popwin = window.open("/cscenterv2/history/history_memo_write.asp?userid=" + userid + "&orderserial=" + orderserial + "&sitename=academy&backwindow=" + "opener.document.history","GotoHistoryMemoWrite","width=600 height=600 scrollbars=yes resizable=no");
	popwin.focus();
}

function FindByIpkumname(){
    var accountname;
    accountname = frmbuyerinfo.accountname.value;

    var gourl = "/cscenter/ordermaster/ordermaster_list.asp?searchfield=etcfield&etcfield=04&etcstring=" + accountname;

    top.listFrame.location.href = gourl;
}

// 올앳카드 매출전표 팝업
function receiptallat(tid){
	var receiptUrl = "http://www.allatpay.com/servlet/AllatBizPop/member/pop_card_receipt.jsp?" +
		"shop_id=10x10_2&order_no=" + tid;
	window.open(receiptUrl,"app","width=410,height=650,scrollbars=0");
}

// 신용카드 매출전표 팝업_이니시스
function receiptinicis(tid){
	var receiptUrl = "https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?" +
		"noTid=" + tid + "&noMethod=1";
	var popwin = window.open(receiptUrl,"INIreceipt","width=415,height=600");
	popwin.focus();
}

// 신용카드 매출전표 팝업_KCP
function receiptkcp(tid){
	var receiptUrl = "https://admin.kcp.co.kr/Modules/Sale/CARD/ADSA_CARD_BILL_Receipt.jsp?" +
		"c_trade_no=" + tid + "&mnu_no=AA000001";
	var popwin = window.open(receiptUrl,"KCPreceipt","width=415,height=600");
	popwin.focus();
}

// 전자보증서 팝업
function insurePrint(orderserial, mallid){
	var receiptUrl = "https://gateway.usafe.co.kr/esafe/ResultCheck.asp?oinfo=" + orderserial + "|" + mallid
	var popwin = window.open(receiptUrl,"insurePop","width=518,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//뱅크페이 현금영수증
function receiptbankpay(tid){
    alert('사용안함. - 관리자 문의 요망');
	//var receiptUrl = "http://www.bankpay.or.kr/pgmember/customcashreceipt.jsp?bill_key1=" + tid;
	//var popwin = window.open(receiptUrl,"BankPayreceipt","width=400,height=560");
	//popwin.focus();
}

//현금영수증 신청 or PopUp - 이니시스 실시간이체 or 무통장
function cashreceipt(iorderserial)
{
    cashreceiptInfo(iorderserial);
    /*
	var receiptUrl = "popcheckreceiptRedirect.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=380,height=750,scrollbars=yes,resizable=yes");
	popwin.focus();
	*/
}

function cashreceiptInfo(iorderserial){
	//var receiptUrl = "/cscenter/taxsheet/popCashReceipt.asp?orderserial=" + iorderserial;
	var receiptUrl = "/cscenterv2/taxsheet/popFnCashReceipt.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popMileageRequest(userid, orderserial, mileage, jukyo) {
	// 필수 : 아이디
	// 옵션 : 주문번호, 마일리지, 적요내용

	if (userid == "") {
		alert("아이디가 없습니다.");
		return;
	}

    var popwin = window.open('/cscenterv2/mileage/pop_mileage_request.asp?userid=' + userid + '&orderserial=' + orderserial + '&mileage=' + mileage + '&jukyo=' + jukyo,'popMileageRequest','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//가상계좌 변경등.
function popDacomCyberPayEdit(iorderserial){
    var popUrl = "/cscenter/cyberAcct/popCyberAcctChange.asp?orderserial=" + iorderserial;
	var popwin = window.open(popUrl,"DcCyberAcct","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//옵션변경
function goOptionChange(iorderserial,itemid)
{
	var popwin1 = window.open('/academy/lecture/popLecOptionEdit.asp?orderserial='+iorderserial+'&lec_idx='+itemid+'','popLecDateEdit','width=700,height=500,scrollbars=yes,resizable=yes');
	popwin1.focus();
}

<% if (orderserial<>"") then %>
//신청내용변경
function goWantChange(iorderserial)
{
	<% If olecjumun.FOneItem.Fitemoption = "0000" Then %>
	alert("강의 일정을 변경해주셔야 합니다.");
	<% Else %>
	var popwin2 = window.open('/cscenterv2/lecture/want_weclass_edit.asp?orderserial='+iorderserial+'','popwin2','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin2.focus();
	<% End If %>
}
<% end if %>
</script>

<% if (orderserial<>"") then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
	<tr height="25">
		<td align="left">
		    &nbsp;&nbsp;
			<input type="button" class="button" value="전체취소" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelOrderLecture('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="부분취소" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelItemLecture('<%= orderserial %>');">
			&nbsp;&nbsp;|&nbsp;&nbsp;
			<input type="button" class="button" value="강좌확정 후 일부환불" class="csbutton" style="width:150px;" onclick="javascript:PopOpenReceiveItemByUpcheLecture('<%= orderserial %>');">
			&nbsp;
			<!--
			<input type="button" class="button" value="회수요청" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByTenTen('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="반품접수" class="csbutton" style="width:70px;" onclick="javascript:PopOpenReceiveItemByUpche('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="맞교환" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemChange('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="누락재발송" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemOmit('<%= orderserial %>');">
			&nbsp;
			<input type="button" class="button" value="서비스발송" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemMore('<%= orderserial %>');">
	        &nbsp;&nbsp;|&nbsp;&nbsp;
			<input type="button" class="button" value="출고시유의사항" class="csbutton" style="width:90px;" onclick="javascript:PopOpenReadMe('<%= orderserial %>');">
			-->
			<!--
			&nbsp;|&nbsp;
			<input type="button" class="button" value="신용카드취소" class="csbutton" style="width:90px;" onclick="javascript:PopOpenCancelCard('<%= orderserial %>');">
			-->
			<input type="button" class="button" value="환불접수" class="csbutton" style="width:90px;" onclick="javascript:PopLectureCSActionCom('','<%= orderserial %>','regcsas','A003','');">
			&nbsp;
			<input type="button" class="button" value="마일리지적립" class="csbutton" style="width:90px;" onclick="javascript:popMileageRequest('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>',0,'');">
			<!--
			<input type="button" class="button" value="외부몰환불요청" class="csbutton" style="width:90px;" onclick="javascript:PopOpenCancelOtherSite('<%= orderserial %>');">
			-->
	    </td>
	    <td align="right">
			<!--<input type="button" class="button" value="기타사항등록" class="csbutton" style="width:90px;" onclick="javascript:PopOpenEtcNote('<%= orderserial %>');">-->
			<!--
			<input type="button" class="button" value="주문메일재발송" class="csbutton" style="width:90px;" onclick="javascript:PopCSMailSendOrder('<%= orderserial %>');">
			&nbsp;
			-->
			<input type="button" class="button" value="영수증재출력" style="width:90px;" onclick="javascript:popOrderReceipt('<%= orderserial %>');">

		</td>
	</tr>
	</table>

	<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
	<tr valign="top">
		<td width="50%">
			<!-- 신청강좌 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			  <tr>
				  <td height="25" colspan="5" bgcolor="<%= adminColor("topbar") %>">
					  <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>강좌 정보</b>
					  <%=CHKIIF(olecjumun.FOneItem.FWeClassYN="Y","&nbsp;&nbsp;&nbsp;-&nbsp;[<font color='blue'><b>단체강좌</b></font>]","")%>
			      </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">강좌명</td>
			    <td colspan="4"><%= olecture.FOneItem.Flec_title %></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">강사명</td>
			    <td>
			    	<%= olecture.FOneItem.Flecturer_name %>(
					<% if (session("ssAdminCLsn") >= 500) then %>
						<%= olecture.FOneItem.Flecturer_id %>
					<% else %>
						(<%= printUserId(olecture.FOneItem.Flecturer_id, 2, "*") %>)
					<% end if %>
			    	)
			    </td>
			    <td bgcolor="<%= adminColor("topbar") %>">강좌코드</td>
			    <td colspan="2"><%= olecjumun.FOneItem.Fitemid %></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">접수기간</td>
			    <td>
	<% if ((now < olecture.FOneItem.Freg_startday) or (now > olecture.FOneItem.Freg_endday)) then %>
			      <font color="red"><%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %></font>
	<% else %>
			      <%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
	<% end if %>
			    </td>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">접수여부</td>
			    <td width="120" colspan="2">
	<% if olecture.FOneItem.Freg_yn="Y" then %>
			접수중
	<% else %>
			      <font color="#CC3333">접수마감</font>
	<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">수강료</td>
			    <td>
	              <%= FormatNumber(olecture.FOneItem.Flec_cost,0) %>
			    </td>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">재료비</td>
			    <td width="120" colspan="2">
	<% if olecture.FOneItem.Fmatinclude_yn="C" then %>
			      선납(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
	<% else %>
			      현장(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
	<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" bgcolor="<%= adminColor("topbar") %>">강좌기간</td>
			    <td colspan="4"><%= olecjumun.FOneItem.Fitemoptionname %></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">강의시작일</td>
			    <td><%= Left(olecture.FOneItem.Flec_startday1, 10) %>

			    </td>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">강의 횟수/시간</td>
			    <td width="120">
					<%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">마감여부</td>
			    <td>
	<% if olecture.FOneItem.IsSoldOut then %>
			      <font color="#CC3333"><b>마감(사유 : <%= olecture.FOneItem.IsSoldOutCauseString %>)</b></font>
	<% else %>
			      접수중
	<% end if %>
			    </td>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">마일리지</td>
			    <td width="120" colspan="2"><%= olecture.FOneItem.Fmileage %> (point)</td>
			  </tr>
			</table>
			<!-- 신청강좌 정보 -->
		</td>
		<td><!-- 신청 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td height="25" colspan="5" bgcolor="<%= adminColor("topbar") %>"><img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>신청 정보</b><%=CHKIIF(olecjumun.FOneItem.FWeClassYN="Y","&nbsp;&nbsp;&nbsp;-&nbsp;[<font color='blue'><b>단체강좌</b></font>]","")%></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">신청인</td>
				<td><%= ojumun.FOneItem.FBuyName %></td>
				<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">아이디</td>
				<td colspan="2">
					<% if (session("ssAdminCLsn") >= 500) then %>
						(<%= ojumun.FOneItem.FUserID %>)
					<% else %>
						(<%= printUserId(ojumun.FOneItem.FUserID, 2, "*") %>)
					<% end if %>
					&nbsp;[<font color="<%= ojumun.FOneItem.GetUserLevelColor %>"><%= ojumun.FOneItem.GetUserLevelName %></font>]
				</td>
			</tr>
			<% If olecjumun.FOneItem.FWeClassYN = "Y" Then %>
				<tr bgcolor="#FFFFFF">
					<td height="25" bgcolor="<%= adminColor("topbar") %>">희망강좌시간</td>
					<td><%= olecjumun.FOneItem.FWantStudyYear %>년 <%= olecjumun.FOneItem.FWantStudyMonth %>월 <%= olecjumun.FOneItem.FWantStudyDay %>일 <%= olecjumun.FOneItem.FWantStudyAmPm %>
						<%= olecjumun.FOneItem.FWantStudyHour %>시 <%= olecjumun.FOneItem.FWantStudyMin %>분
						</td>
					<td width="100" bgcolor="<%= adminColor("topbar") %>">희망신청인원</td>
					<td width="120"><%=olecjumun.FOneItem.Fitemno%> 명(<%= fnWeClassStudyWho(olecjumun.FOneItem.FWantStudyWho) %>)</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">업체(동호회)명</td>
					<td colspan="4"><%= olecjumun.FOneItem.FWantStudyName %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">희망강의장소</td>
					<td colspan="4"><%= olecjumun.FOneItem.FWantStudyPlace %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">실제강의일정</td>
					<td colspan="4">[<%= olecjumun.FOneItem.Fitemoption %>]<%= olecjumun.FOneItem.Fitemoptionname %>
						&nbsp;<input type="button" class="button" value="강의일정변경" onClick="goOptionChange('<%= olecjumun.FOneItem.Forderserial %>','<%= olecjumun.FOneItem.Fitemid %>');">
						<% If olecjumun.FOneItem.Fitemoption = "0000" Then %><b><br>▲ 실제강의 일정을 꼭 변경해주세요!!</b><% End If %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>"></td>
					<td><input type="button" class="button" value="신청내용변경<%= CHKIIF(olecjumun.FOneItem.Fipkumdiv=2," 및 결제진행","") %>" onClick="goWantChange('<%= olecjumun.FOneItem.Forderserial %>');"></td>
					<td width="100" bgcolor="<%= adminColor("topbar") %>">신청상태</td>
					<td width="120"><font color="<%= olecjumun.FOneItem.CancelYnColor %>"><b><%= olecjumun.FOneItem.CancelYnName %></b></font>/<font color="<%= olecjumun.FOneItem.IpkumDivColor %>"><%= CHKIIF(olecjumun.FOneItem.FIpkumDiv="3","결제대기",olecjumun.FOneItem.IpkumDivName) %></font></td>
				</tr>
			<% Else %>
				<tr bgcolor="#FFFFFF">
					<td height="25" bgcolor="<%= adminColor("topbar") %>">강좌시간</td>
					<td colspan="4"><%= olecjumun.FOneItem.Flecturestartday %> &nbsp; - &nbsp; <%= olecjumun.FOneItem.Flectureendday %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">신청 / 정원</td>
					<td><%= olecjumun.FOneItem.Flimitsoldno %> / <%= olecjumun.FOneItem.Flimitmaxno %> (최소 : <%= olecjumun.FOneItem.Flimitminno %> / 대기 : <%= olecjumun.FOneItem.Flimitwaitno %>)</td>
					<td width="100" bgcolor="<%= adminColor("topbar") %>">신청상태</td>
					<td width="120"><font color="<%= olecjumun.FOneItem.CancelYnColor %>"><b><%= olecjumun.FOneItem.CancelYnName %></b></font>/<font color="<%= olecjumun.FOneItem.IpkumDivColor %>"><%= olecjumun.FOneItem.IpkumDivName %></font></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">접수여부</td>
					<td colspan="4"><% if olecture.FOneItem.IsSoldOut then %><font color="#CC3333"><b>마감(사유 : <%= olecture.FOneItem.IsSoldOutCauseString %>)</b></font><% else %>접수중<% end if %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">강의시작일</td>
					<td><%= Left(olecjumun.FOneItem.Flecturestartday, 10) %></td>
					<td width="100" bgcolor="<%= adminColor("topbar") %>">취소가능여부</td>
					<td width="120"><% if (Left(DateAdd("d",3,now), 10)  > Left(olecture.FOneItem.Flec_startday1,10)) then %><font color="red">취소불가</font><% else %>취소가능<% end if %></td>
				</tr>
			<% End If %>
			<tr bgcolor="#FFFFFF">
				<td height="25" width="100" bgcolor="<%= adminColor("topbar") %>">약도</td>
				<td colspan="5"><a href="<%= olecture.FOneItem.Flec_mapimg %>" target="_blank"><%= olecture.FOneItem.Flec_mapimg %></a></td>
			</tr>
			</table>
			<!-- 신청 정보 -->
		</td>
	</tr>
	</table>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr valign="top">
		<td>
			<!-- 구매상품정보 -->
			<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
			    <td colspan="10">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="500">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매상품정보</b>
					    	&nbsp;
					    	[<b><%= orderserial %></b>]
					    	&nbsp;
					    	<input type="button" class="button" value="관련CS <%= totalascount %>건" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List_Lecture('<%= orderserial %>','','');">
    				    </td>
    				    <td align="right">
    				    	<input type="button" class="button" value="미출고상품보기" class="csbutton" style="width:90px;" onclick="misendmaster('<%= orderserial %>');">
    				    </td>
    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr height="345" bgcolor="#FFFFFF">
			    <td valign="top">
					<table height="320" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
					<tr height="100%">
						<td colspan="13">
							<iframe name="orderdetail" src="lecturedetail_item_list.asp?orderserial=<%= orderserial %>" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
						</td>
					<tr>
					</table>
			    </td>
			</tr>
			</table>
			<!-- 구매상품정보 -->
		    <p>
			<!-- 하단부분 -->
	        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        	<tr valign="top" height="210">
        		<td colspan="3">
        		    <!-- 주문건 History -->
        		    <table width="100%" height="210" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    				<form name="frmhistory" onsubmit="return false;">
    				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
    				    <td colspan="10">
    				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    				    		<tr>
    				    			<td>
    				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMemo('','<%= orderserial %>')"><b>MEMO</b></a>
    		    				    	[<b><%= orderserial %></b>]
    		    				    	|
    		    				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCS('','<%= orderserial %>')"><b>CS 처리건</b></a>
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryMileage('<%= ojumun.FOneItem.FUserID %>','')"><b>마일리지</b></a>
    		                            [<b>
											<% if (session("ssAdminCLsn") >= 500) then %>
												(<%= ojumun.FOneItem.FUserID %>)
											<% else %>
												(<%= printUserId(ojumun.FOneItem.FUserID, 2, "*") %>)
											<% end if %>
    		                            </b>]
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryCoupon('<%= ojumun.FOneItem.FUserID %>','')"><b>쿠폰</b></a>
    		                            |
    		                            <img src="/images/icon_star.gif" align="absbottom">&nbsp;<a href="javascript:GotoHistoryQna('<%=ojumun.FOneItem.FUserID%>','<%If ojumun.FOneItem.FUserID = "" Then response.write orderserial End If %>')"><b>1:1상담</b></a>
    		    				    </td>
    		    				    <td width="100" align="right">
    		    				    	<input type="button" class="button" name="writebutton" value="History등록" class="csbutton" onclick="OpenHistoryWindow('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>');">
    		    				    </td>
    		    				</tr>
    		    			</table>
    		    		</td>
    				</tr>
    				</form>
       				<tr>
    				    <td>
    				        <iframe name="history" src="blank.asp" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
        `			    </td>
    				</tr>
        			</table>
        			<!-- 주문건 History-->
        		</td>
            	<td width="5"></td>
        		<td width="225">
        		    <!-- 주문정보 -->
        		    <table width="225" height="210" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    				    <td colspan="10">
    				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    				    		<tr>
    				    			<td width="100">
    				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문 정보</b>
    		    				    </td>
    		    				    <td align="right">
    		    				    	<input type="button" class="button" value="다음상태진행" class="csbutton" onclick="javascript:PopNextIpkumDiv('<%= orderserial %>');">
    		    				    </td>
    		    				</tr>
    		    			</table>
    		    		</td>
    				</tr>
    				<!--
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
    				    <td bgcolor="#FFFFFF"><%= orderserial %></td>
    				</tr>
    				-->
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">결제방법</td>
    				    <td bgcolor="#FFFFFF">
    				        <%= ojumun.FOneItem.JumunMethodName %>
    				        [<font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>]
    				        <% if ojumun.FOneItem.FCancelYn<>"N" then %>
    				        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
    				        <% end if %>
    				        <% if ojumun.FOneItem.FokcashbagSpend<>0 then %>
    				        <br>(캐시백사용 : <strong><%= formatNumber(ojumun.FOneItem.FokcashbagSpend,0) %></strong>)
    				        <% end if %>
    				    </td>
    				</tr>
    				<% if ojumun.FOneItem.FAccountDiv="7" then %>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">
    				    <% if ojumun.FOneItem.IsDacomCyberAccountPay then %>
    				    <a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')"><font color="red">가상</font></a>
    				    <% else %>
    				    <a href="javascript:popDacomCyberPayEdit('<%= orderserial %>')">일반</a>
    				    <% end if %>
    				    </td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FAccountNo %></td>
    				</tr>
    				<% end if %>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">주문일시</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">입금확인</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FIpkumDate %></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">주문통보</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbaljudate %></td>
    				</tr>
    				<!--
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">출고일시</td>
    				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbeadaldate %></td>
    				</tr>
    				-->
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">승인번호</td>
    				    <td bgcolor="#FFFFFF">
    				        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FAuthcode %>" readonly size="20">
    				    </td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
    				    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FPaygatetID %>" readonly></td>
    				</tr>
    				<tr>
    				    <td bgcolor="<%= adminColor("topbar") %>">증빙서류</td>
    				    <td bgcolor="#FFFFFF">

    				    	<!-- All@ 결제일 경우 -->
    	                  	<% if (trim(ojumun.FOneItem.Faccountdiv)="80") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    	                  	    <input type="button" class="button" value="신용" onclick="javascript:receiptallat('<%= ojumun.FOneItem.Fpaygatetid %>');">
    	                  	<% end if %>

    				        <!-- 신용카드 매출전표 -->
    	                  	<% if (ojumun.FOneItem.FAccountDiv="100") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    	                  		<% if ojumun.FOneItem.Fpaygatetid<>"" then %>
    		                  		<% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") then %>
    		                  		    <input type="button" class="button" value="신용" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.FPaygatetID %>');">
    		                  		<% else %>
    		                  		    <input type="button" class="button" value="신용" onclick="javascript:receiptkcp('<%= ojumun.FOneItem.FPaygatetID %>')">
    		                  		<% end if %>
    		                  	<% end if %>
    		                <% end if %>

    		                <!-- 현금영수증 발행확인 실시간이체-->
    						<% if (ojumun.FOneItem.FAccountDiv="20") and (ojumun.FOneItem.FIpkumDiv >= 4) then %>
    						    <% if (Left(ojumun.FOneItem.Fpaygatetid,9)="IniTechPG") then %>
    						        <% if ojumun.FOneItem.FAuthCode<>"" then %> <!-- 결제시 현금영수증 신청한 경우 -->

    						            <input type="button" class="button" value="현금" onclick="javascript:receiptinicis('<%= ojumun.FOneItem.Fpaygatetid %>')">

    						            <!-- input type="button" class="button" value="현금2" onclick="javascript:cashreceipt('<%= ojumun.FOneItem.ForderSerial %>')" -->
    						        <% elseif (ojumun.FOneItem.FcashreceiptReq="T") then %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">세금</a>
    						        <% else %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">상태보기</a>
    						        <% end if %>
    						    <% else %>
    						        <% if (ojumun.FOneItem.Fcashreceiptreq<>"") then %>
    						        <input type="button" class="button" value="현금" onclick="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">
    						        <% else %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">상태보기</a>
    						        <% end if %>
    						        <!-- <input type="button" class="button" value="현금" onclick="javascript:receiptbankpay('<%= ojumun.FOneItem.Fpaygatetid %>')"> -->
    						    <% end if %>
    						<% end if %>

    						<!-- 현금영수증 발행확인 무통장 경우 -->
    						<% if (ojumun.FOneItem.FAccountDiv="7") then %>
								<% if (ojumun.FOneItem.Fauthcode<>"") then %>
                                    <input type="button" class="button" value="현금" onclick="javascript:cashreceipt('<%= ojumun.FOneItem.ForderSerial %>')">
                                <% elseif (ojumun.FOneItem.FcashreceiptReq="R") then %>
                                <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">현금영수증 발행요청 상태</a>
                                <% elseif (ojumun.FOneItem.FcashreceiptReq="T") then %>
    						        <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">세금</a>
                                <% else %>
                                    <% if (ojumun.FOneItem.FIpkumdiv>3) then %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">상태보기</a>
                                    <% else %>
                                    <a href="javascript:cashreceiptInfo('<%= ojumun.FOneItem.ForderSerial %>')">상태보기</a>
                                    <% end if %>
                                <% end if %>

                            <% end if %>
                            <!-- 세금계산서 발행요청 여부 -->

    						<!-- 전자보증보험 -->
                    		<% if (ojumun.FOneItem.FInsureCd = "0") then %>
                    		    <input type="button" class="button" value="보증" onclick="javascript:insurePrint('<%= ojumun.FOneItem.ForderSerial %>','ZZcube1010')">
    						<% end if %>
    				    </td>
    				</tr>
        			</table>
        			<!-- 주문정보 -->
        		</td>
        	</tr>
	        </table>
		</td>
		<td width="5"></td>
		<td width="250" align="right">
			<!-- 구매자정보 -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmbuyerinfo" onsubmit="return false;">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			    <td colspan="2">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="100">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매자 정보</b>
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="구매자정보수정" class="csbutton" onclick="javascript:PopBuyerInfo('<%= orderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr height="24">
			    <td bgcolor="<%= adminColor("topbar") %>">구매자ID</td>
			    <td bgcolor="#FFFFFF">
					<% if (session("ssAdminCLsn") >= 500) then %>
						(<%= ojumun.FOneItem.FUserID %>)
					<% else %>
						(<%= printUserId(ojumun.FOneItem.FUserID, 2, "*") %>)
					<% end if %>
			    &nbsp;[<font color="<%= ojumun.FOneItem.GetUserLevelColor %>"><%= ojumun.FOneItem.GetUserLevelName %></font>]</td>
			</tr>
			<tr height="23">
			    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
			    <td bgcolor="#FFFFFF"><%= orderserial %></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">구매자명</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyname" value="<%= ojumun.FOneItem.FBuyName %>" size="8" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyphone" value="<%= ojumun.FOneItem.FBuyPhone %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="buyhp" value="<%= ojumun.FOneItem.FBuyHp %>" readonly>
			        <input type="button" name="buyhp" class="button" value="SMS" onclick="PopCSSMSSendNew({reqhp:'<%= ojumun.FOneItem.FBuyHp %>', orderserial:'<%= ojumun.FOneItem.Forderserial %>', userid:'<%= ojumun.FOneItem.Fuserid %>'});">
			    </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="buyemail" value="<%= ojumun.FOneItem.FBuyEmail %>" size="20" readonly>
			        <input type="button" name="email" class="button" value="mail" onclick="javascript:PopCSMailSend('<%= ojumun.FOneItem.FBuyEmail %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>');">
			    </td>
			</tr>

			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">입금자명</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="accountname" value="<%= ojumun.FOneItem.FAccountName %>" size="14" readonly>
			        <input type="button" class="button" value="검색" class="csbutton" onclick="FindByIpkumname()">
			        <acronym title="<%= ojumun.FOneItem.Faccountno %>"><%= left(ojumun.FOneItem.Faccountno,2) %></acronym>
			    </td>
			</tr>
			</form>
			</table>
			<!-- 구매자정보 -->

			<!-- 배송정보 -->
			<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmreqinfo" onsubmit="return false;">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			    <td colspan="2">
			    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			    		<tr>
			    			<td width="100">
			    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>배송 정보</b>
	    				    </td>
	    				    <td align="right">
	    				    	<input type="button" class="button" value="배송지정보수정" class="csbutton" onclick="javascript:PopReceiverInfo('<%= orderserial %>');">
	    				    </td>
	    				</tr>
	    			</table>
	    		</td>
			</tr>
			<tr>
			    <td width="100" bgcolor="<%= adminColor("topbar") %>">수령인명</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqName %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
			    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqPhone %>" readonly></td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqHp %>" readonly>
			        <input type="button" name="reqhp" class="button" value="SMS" onclick="PopCSSMSSendNew({reqhp:'<%= ojumun.FOneItem.FReqHp %>', orderserial:'<%= ojumun.FOneItem.Forderserial %>', userid:'<%= ojumun.FOneItem.Fuserid %>'});">
			    </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">배송주소</td>
			    <td bgcolor="#FFFFFF">
			        <input type="text" class="text_ro" name="txzip1" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7" readonly>
			        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqZipAddr %>" size="18" readonly><br>
			        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FReqAddress %></textarea>
                </td>
			</tr>
			<tr>
			    <td bgcolor="<%= adminColor("topbar") %>">기타사항</td>
			    <td bgcolor="#FFFFFF">
			        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FComment %></textarea>
			    </td>
			</tr>
			</form>
			</table>
			<!-- 배송정보 -->

			<!-- 해외배송일 경우 해외배송 관련 아닐경우, 플라워주문관련 -->

			<% if ojumun.FOneItem.IsForeignDeliver=true then %>
				<!-- 해외배송 관련 -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="2">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>해외배송</b>
		    				    </td>
		    				    <td align="right">
							    	<input type="button" class="button" value="국가별발송조건" class="csbutton" style="width:120px;" onclick="popForeignDeliverInfo('<%= ojumun.FOneItem.FDlvcountryCode %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr>
				    <td width="55" bgcolor="<%= adminColor("topbar") %>">상품중량</td>
				    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FitemWeigth %>(g)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">박스중량</td>
				    <td bgcolor="#FFFFFF">200(g)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">배송국가명</td>
				    <td bgcolor="#FFFFFF">
				    	<%= ojumun.FOneItem.FcountryNameEn %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">배송국가</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" value="<%= ojumun.FOneItem.FDlvcountryCode %>" size="2" readonly>
				    	<input type="text" class="text_ro" value="<%= ojumun.FOneItem.FemsAreaCode %>" size="2" readonly>
						<input type="button" class="button" value="요금표보기" class="csbutton" style="width:100px;" onclick="popForeignDeliverPay('<%= ojumun.FOneItem.FemsAreaCode %>');">
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">EMS요금</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" style="text-align:right;" value="<%= FormatNumber(ojumun.FOneItem.FemsDlvCost,0) %>" size="8" maxlength="10" readonly>원
				    </td>
				</tr>
				<%'If ojumun.FOneItem.FemsInsureYn = "Y" Then %>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">보험가입(<%=ojumun.FOneItem.FemsInsureYn%>)</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" style="text-align:right;" value="<%=FormatNumber(ojumun.FOneItem.FemsInsurePrice,0)%>" size="8" maxlength="10" readonly>원
				    </td>
				</tr>
				<%'End If %>
				</table>
			<% else %>
				<!-- 플라워 주문  -->
				<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
				    <td colspan="2">
				    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				    		<tr>
				    			<td width="100">
				    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>플라워관련</b>
		    				    </td>
		    				    <td align="right">
		    				    	<input type="button" class="button" value="플라워메세지변경" class="csbutton" onclick="javascript:PopFlowerDeliverInfo('<%= orderserial %>');">
		    				    </td>
		    				</tr>
		    			</table>
		    		</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">선택</td>
				    <td bgcolor="#FFFFFF">
				        <input type="radio" name="cardribbon" value="1" <% if ojumun.FOneItem.Fcardribbon="1" then response.write "checked" %> >카드
				        <input type="radio" name="cardribbon" value="2" <% if ojumun.FOneItem.Fcardribbon="2" then response.write "checked" %> >리본
				        <input type="radio" name="cardribbon" value="3" <% if ojumun.FOneItem.Fcardribbon="3" then response.write "checked" %> >없음
				    </td>
				</tr>
				<tr>
				    <td colspan="2" bgcolor="#FFFFFF">
				        <textarea class="textarea_ro" name="message" rows="3" cols="37" readonly><%= ojumun.FOneItem.Fmessage %></textarea><br>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">FROM</td>
				    <td bgcolor="#FFFFFF">
				        <input type="text" class="text_ro" name="fromname" value="<%= ojumun.FOneItem.Ffromname %>" size="20" maxlength="20" readonly>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">희망일</td>
				    <td bgcolor="#FFFFFF">
				        <input type="text" class="text_ro" name="reqdate" value="<%= ojumun.FOneItem.Freqdate %>" size="10" readonly>일
				        <input type="text" class="text_ro" name="reqtime" value="<%= ojumun.FOneItem.GetReqTimeText %>" size="10" readonly>
				    </td>
				</tr>
				</table>
			<% end if %>
		</td>
	</tr>
	</table>
	<br>
<% else %>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	    <tr height="50">
	        <td align="center"> [ 상세내역을 보시려면 주문번호를 선택 하세요 ]</td>
	    </tr>
	</table>
<% end if %>

<% if (orderserial <> "") then %>
	<script type="text/javascript">
	    GotoHistoryCS('','<%= orderserial %>');
	</script>
<% end if %>

<%
set ojumun = Nothing
set oaslist = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
