<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%

'==============================================================================
dim orderserial, oordermaster, oorderdetail, oorderdetailitemmakergroup, oaslist

orderserial = RequestCheckvar(request("orderserial"),16)

set oordermaster = new CRequestLecture
oordermaster.FRectOrderSerial = orderserial
oordermaster.GetRequestLectureMasterOne

set oorderdetail = new CRequestLecture
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.CRequestLectureDetailList

if (oordermaster.FResultCount < 1) then
        response.write "<script>alert('잘못된 주문번호입니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = oordermaster.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture


'==============================================================================
dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = oordermaster.FOneItem.Fitemid
if (olecschedule.FRectIdx = "") then
    olecschedule.FRectIdx = "0"
end if

olecschedule.GetOneLecSchedule


'==============================================================================
dim ocsaslist
set ocsaslist = New CCSASList

ocsaslist.FRectOrderSerial = orderserial

ocsaslist.GetCSASMasterList

dim totalrequestrepay, totalresultrepay

totalrequestrepay = 0
totalresultrepay = 0
for i = 0 to ocsaslist.FResultCount - 1
    if (ocsaslist.FItemList(i).Fdeleteyn = "N") then
        if (ocsaslist.FItemList(i).Fcurrstate = "7") then
            totalresultrepay = totalresultrepay + ocsaslist.FItemList(i).Frefundresult
        end if
        totalrequestrepay = totalrequestrepay + ocsaslist.FItemList(i).Frefundrequire
    end if
next


'==============================================================================
dim divcd, divcdname

divcd = request("divcd")
if (divcd = "3") then
        divcdname = "환불요청"
elseif (divcd = "5") then
        divcdname = "외부몰환불요청"
elseif (divcd = "6") then
        divcdname = "배송유의사항"
elseif (divcd = "7") then
        divcdname = "신용카드/상품권/실시간이체취소요청"
elseif (divcd = "8") then
        divcdname = "상품준비중취소"
elseif (divcd = "9") then
        divcdname = "기타내역"
elseif (divcd = "20") then
        divcdname = "전체취소"
elseif (divcd = "21") then
        divcdname = "부분취소"
else
        response.write "<script>alert('잘못된 접속입니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'==============================================================================
dim baesongmethodstr, refundbeasongpay

baesongmethodstr = ""
refundbeasongpay = 0



'==============================================================================
dim i, ix
dim haveupchebaesong, havetentenbaesong, isavailableitem

dim directBankCancelAvail : directBankCancelAvail = false

if (LCase(Left(oordermaster.FOneItem.Fpaygatetid,Len("initechpg_")))="initechpg_") then
    directBankCancelAvail = true
end if

%>


<script>
// ============================================================================
// 저장하기
function SubmitSave() {
        var e;
        var ischecked = false;

        if (frm.causecd.value == "") {
                alert("건별 사유구분을 선택하세요.");
                return;
        }

        if (frm.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }

        if (frm.returnmethod[frm.returnmethod.selectedIndex].value=="realtimetransfer"){
            <% if (Not directBankCancelAvail) then %>
            alert('이니페이가 아닌경우 실시간이체 취소 불가능! \n\n무통장 환불을 사용하세요');
            return;
            <% end if %>
        }

        if ((frm.returnmethod[frm.returnmethod.selectedIndex].value == "bank")) {
                if (frm.rebankaccount.value == "") {
                        alert("계좌번호를 입력하세요.");
                        return;
                }

                if (frm.rebankname.value == "") {
                        alert("예금주명을 입력하세요.");
                        return;
                }

                if (frm.rebankname.selectedIndex == 0) {
                        alert("은행을 선택하세요.");
                        return;
                }
        }

        if (confirm("등록하시겠습니까?") == true) {
                document.frm.submit();
        }
}

function CloseWindow() {
        opener.focus();
        window.close();
}


// ============================================================================
// 사유창 표시관련
function ShowCauseSelectWindow(idx) {
        var html = "<table bgcolor='#ED5F00' align='right' width='550' class='a' cellspacing='1'>";
        html = html + "<tr>";
        html = html + "<td height='25' width='100' bgcolor='#FFFFFF' colspan='2'><table width='540' class='a'><tr><td>사유선택</td><td align='right'><a href='javascript:WriteCause(\"" + idx + "\",\"\",\"\")'>[사유삭제]</a> <a href='javascript:hideCauseSelectWindow(\"" + idx + "\")'>[닫기]</a></td></tr></table></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>공통</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"1\")'>단순변심</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"2\")'>사이즈</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"3\")'>품절</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"4\")'>재주문</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"4\",\"99\")'>기타</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>상품관련</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"1\")'>상품불량</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"2\")'>상품불만족</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"3\")'>상품등록오류</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"4\")'>상품설명불량</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"5\",\"99\")'>기타</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>물류관련</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"1\")'>오발송</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"2\")'>구매상품누락</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"3\")'>사은품누락</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"4\")'>상품파손</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"5\")'>상품품절</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"6\")'>배송지연</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"6\",\"99\")'>기타</a></td>";
        html = html + "</tr>";
        html = html + "<tr>";
        html = html + "<td height='25' bgcolor='#FFFFFF'>택배사관련</td>";
        html = html + "<td bgcolor='#FFFFFF'><a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"1\")'>택배사파손</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"2\")'>택배사분실</a> / <a href='javascript:WriteCause(\"" + idx + "\",\"7\",\"99\")'>기타</a></td>";
        html = html + "</tr>";
        html = html + "<table>";

        var id = eval("causepop" + idx);
        id.innerHTML = html;
}

function hideCauseSelectWindow(idx) {
        var id = eval("causepop" + idx);
        id.innerHTML = "";
}

function WriteCause(idx, causecd, causedetail) {
        var icausestring = "";
        var index;

        icausestring = GetCauseString(causecd, causedetail);

        var ocausestring = eval("causestring" + idx);
        ocausestring.innerHTML = icausestring;

        var ocausecd = eval("frm.causecd" + idx);
        ocausecd.value = causecd;

        var ocausedetail = eval("frm.causedetail" + idx);
        index = icausestring.indexOf(" > ");
        if (index == -1) {
                ocausedetail.value = "";
        } else {
                ocausedetail.value = icausestring.substring(index + 3);
        }

        if (idx != "") {
                WriteMasterCause(causecd, causedetail);
        }
        hideCauseSelectWindow(idx);
}

function WriteMasterCause(causecd, causedetail) {
        var icausestring = "";

        icausestring = GetCauseString(causecd, causedetail);

        var ocausestring = eval("causestring");
        ocausestring.innerHTML = icausestring;

        var ocausecd = eval("frm.causecd");
        ocausecd.value = causecd;

        var ocausedetail = eval("frm.causedetail");
        index = icausestring.indexOf(" > ");
        if (index == -1) {
                ocausedetail.value = "";
        } else {
                ocausedetail.value = icausestring.substring(index + 3);
        }
}

function GetCauseString(causecd, causedetail) {
        var causestring = "등록하기";

        if (causecd == 4) {
                causestring = "공통";

                if (causedetail == 1) {
                        causestring = causestring + " > 단순변심";
                } else if (causedetail == 2) {
                        causestring = causestring + " > 사이즈";
                } else if (causedetail == 3) {
                        causestring = causestring + " > 품절";
                } else if (causedetail == 4) {
                        causestring = causestring + " > 재주문";
                } else {
                        causestring = causestring + " > 기타";
                }
        } else if (causecd == 5) {
                causestring = "상품관련";

                if (causedetail == 1) {
                        causestring = causestring + " > 상품불량";
                } else if (causedetail == 2) {
                        causestring = causestring + " > 상품불만족";
                } else if (causedetail == 3) {
                        causestring = causestring + " > 상품등록오류";
                } else if (causedetail == 4) {
                        causestring = causestring + " > 상품설명불량";
                } else {
                        causestring = causestring + " > 기타";
                }
        } else if (causecd == 6) {
                causestring = "물류관련";

                if (causedetail == 1) {
                        causestring = causestring + " > 오발송";
                } else if (causedetail == 2) {
                        causestring = causestring + " > 구매상품누락";
                } else if (causedetail == 3) {
                        causestring = causestring + " > 사은품누락";
                } else if (causedetail == 4) {
                        causestring = causestring + " > 상품파손";
                } else if (causedetail == 5) {
                        causestring = causestring + " > 상품품절";
                } else if (causedetail == 6) {
                        causestring = causestring + " > 배송지연";
                } else {
                        causestring = causestring + " > 기타";
                }
        } else if (causecd == 7) {
                causestring = "택배사관련";

                if (causedetail == 1) {
                        causestring = causestring + " > 택배사파손";
                } else if (causedetail == 2) {
                        causestring = causestring + " > 택배사분실";
                } else {
                        causestring = causestring + " > 기타";
                }
        }

        return causestring;
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <form name="frm" method="post" action="do_lec_write.asp" onsubmit="return false;">
    <input type="hidden" name="mode" value="cancelorder">
    <input type="hidden" name="sitename" value="<%= oordermaster.FOneItem.Fsitename %>">
    <input type="hidden" name="orderserial" value="<%= oordermaster.FOneItem.FOrderSerial %>">
    <input type="hidden" name="divcd" value="<%= divcd %>">
    <input type="hidden" name="causecd" value="">
    <input type="hidden" name="causedetail" value="">
    <input type="hidden" name="detailitemlist" value="">
    <tr height="10" valign="bottom">
	    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	    <td background="/images/tbl_blue_round_02.gif"></td>
	    <td background="/images/tbl_blue_round_02.gif"></td>
	    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td background="/images/tbl_blue_round_06.gif">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS처리 등록</b>
	    	[<b><%= oordermaster.FOneItem.FOrderSerial %></b>]
	    </td>
	    <td align="right" background="/images/tbl_blue_round_06.gif">
	    <input type="button" name="btnsave" value="등록하기" onclick="SubmitSave();">
	    <input type="button" value="닫기" onclick="CloseWindow();">
	    </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr>
	    <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2" background="/images/tbl_blue_round_06.gif">

            <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
                <tr height="30" bgcolor="#FFFFFF">
            		<td width="70" rowspan="2" bgcolor="#DDDDFF">구분</td>
            	    <td rowspan="2"><font style='line-height:100%; font-size:25px; color:blue; font-family:돋움; font-weight:bold'><%= divcdname %></font></td>
            	    <td width="100" bgcolor="#DDDDFF">접수일시</td>
            	    <td width="250"><b><%= now %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            	    <td bgcolor="#DDDDFF">등록자ID</td>
            	    <td><b><%= session("ssBctId") %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td bgcolor="#DDDDFF">제목</b></td>
            	    <td><input type="text" name="title" size="50" value="<%= divcdname %>"></td>
            	    <td bgcolor="#DDDDFF">주문번호</td>
            	    <td><b><%= oordermaster.FOneItem.FOrderSerial %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td bgcolor="#DDDDFF">사유구분</b></td>
            	    <td><a href="javascript:ShowCauseSelectWindow('')"><div id='causestring'>등록하기</div></a><div id="causepop" style="position:absolute;"></div></td>
            	    <td bgcolor="#DDDDFF">구매자명</td>
            	    <td><b><%= oordermaster.FOneItem.FBuyName %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            		<td rowspan="2" bgcolor="#DDDDFF">접수내용</td>
            	    <td rowspan="2"><textarea rows="2" cols="50" name="contents_jupsu"></textarea></td>
            	    <td bgcolor="#DDDDFF">구매자ID</td>
            	    <td><b><%= oordermaster.FOneItem.FUserID %></b></td>
            	</tr>
            	<tr height="30" bgcolor="#FFFFFF">
            	    <td bgcolor="#DDDDFF">상태 / 거래상태</td>
            	    <td><b><font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font> / <font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font></b></td>
            	</tr>
            </table>

        </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
    <tr height="20">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td background="/images/tbl_blue_round_06.gif">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>신청정보</b>
	    </td>
	    <td align="right" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
    <tr>
	    <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
			<!-- 신청강좌 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강좌명 / 코드</td>
			    <td colspan="3"><%= olecture.FOneItem.Flec_title %> / <%= oordermaster.FOneItem.Fitemid %></td>
			    <td rowspan="4" width="100"><img src="<%= olecture.FOneItem.Flistimg %>"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강사명</td>
			    <td><%= olecture.FOneItem.Flecturer_name %>(<%= olecture.FOneItem.Flecturer_id %>)</td>
			    <td width="100" bgcolor="#DDDDFF"></td>
			    <td width="250"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강의시작일</td>
			    <td><%= Left(olecture.FOneItem.Flec_startday1, 10) %>

			    </td>
			    <td width="100" bgcolor="#DDDDFF">취소가능여부</td>
			    <td width="250">
<% if (Left(DateAdd("d",3,now), 10)  > Left(olecture.FOneItem.Flec_startday1,10)) then %>
			      <font color="red">취소불가</font>
<% else %>
			      취소가능
<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">접수기간</td>
			    <td>
<% if ((now < olecture.FOneItem.Freg_startday) or (now > olecture.FOneItem.Freg_endday)) then %>
			      <font color="red"><%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %></font>
<% else %>
			      <%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
<% end if %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">접수여부</td>
			    <td width="250">
<% if olecture.FOneItem.Freg_yn="Y" then %>
			접수중
<% else %>
			      <font color="#CC3333">접수마감</font>
<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">수강료</td>
			    <td>
                  <%= FormatNumber(olecture.FOneItem.Flec_cost,0) %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">재료비</td>
			    <td width="250" colspan="2">
<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
			      포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
<% else %>
			      별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
<% end if %>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강의 횟수/시간</td>
			    <td>
                  <%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
			    </td>
			    <td width="100" bgcolor="#DDDDFF">인원</td>
			    <td width="250" colspan="2">
			      <%= olecture.FOneItem.Flimit_sold %> / <%= olecture.FOneItem.Flimit_count %> (최소 : <%= olecture.FOneItem.Fmin_count %>)
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">마감여부</td>
			    <td>
<% if olecture.FOneItem.IsSoldOut then %>
			      <font color="#CC3333"><b>마감(사유 : <%= olecture.FOneItem.IsSoldOutCauseString %>)</b></font>
<% else %>
			      접수중
<% end if %>
			    </td>
			    <td width="100" bgcolor="#DDDDFF">마일리지</td>
			    <td width="250" colspan="2"><%= olecture.FOneItem.Fmileage %> (point)</td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">약도</td>
			    <td colspan="5">
                  <%= olecture.FOneItem.Flec_mapimg %>
			    </td>
			  </tr>
			</table>
			<!-- 신청강좌 정보 -->
			<br>



			<!-- 신청인원 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">신청인원</td>
			    <td>
<% for i = 0 to oorderdetail.FResultCount - 1 %>
    <% if (oorderdetail.FItemList(i).Fcancelyn <> "N") then %>
                  <font color="<%= oorderdetail.FItemList(i).CancelStateColor %>"><%= oorderdetail.FItemList(i).Fentryname %>(<%= oorderdetail.FItemList(i).Fentryhp %>/<%= oorderdetail.FItemList(i).CancelStateStr %>)</font>
    <% else %>
                  <%= oorderdetail.FItemList(i).Fentryname %>(<%= oorderdetail.FItemList(i).Fentryhp %>/<%= oorderdetail.FItemList(i).CancelStateStr %>)
    <% end if %>
<% next %>
			    </td>
			  </tr>
			</table>
			<!-- 신청인원 정보 -->


        </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif"></td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10">
	    <td background="/images/tbl_blue_round_04.gif"></td>
	    <td colspan="2" background="/images/tbl_blue_round_06.gif">

	        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	            <tr height="20">
            	    <td background="/images/tbl_blue_round_06.gif">
            	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>환불금액 계산(<%= oordermaster.FOneItem.JumunMethodName %>)</b>
            	    </td>
            	    <td width="10"></td>
                    <td background="/images/tbl_blue_round_06.gif">
            	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>환불 계좌번호</b>&nbsp;[<%= oordermaster.FOneItem.FUserID %>]
            	    </td>
            	</tr>



	            <tr>
	                <td valign="top" width="50%">
	                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="BABABA">
                        	<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30" width="100">결제금액</td>
                        		<td align="right" width="170">
<% if (oordermaster.FOneItem.Fipkumdiv >= 4) then %>
                        		<%= oordermaster.FOneItem.Fsubtotalprice %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% else %>
                                        0&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% end if %>
                        		</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">사용쿠폰</td>
                        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.Ftencardspend,0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">사용마일리지</td>
                        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">기존환불요청</td>
                        		<td align="right">
                                  <%= totalrequestrepay %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">기존환불완료</td>
                        		<td align="right">
                                  <%= totalresultrepay %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">환불예정액</td>
                        		<td align="right">
<%
'환불예정금액
if (oordermaster.FOneItem.Fipkumdiv >= 4) then
    i = oordermaster.FOneItem.Fsubtotalprice
else
    i = 0
end if
%>
                        		  <%= FormatNumber(i,0) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        		</td>
                        		<td>
<% if ((oordermaster.FOneItem.Fcancelyn <> "Y") and (oordermaster.FOneItem.Ftencardspend > 0)) then %>
                                  사용쿠폰취소
<% end if %>
<% if ((oordermaster.FOneItem.Fcancelyn <> "Y") and (oordermaster.FOneItem.Fmiletotalprice > 0)) then %>
                                  사용마일리지취소
<% end if %>
                        		</td>
                        	</tr>
                        	<input type="hidden" name="refundrequire" value="<%= i %>">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">환불방식</td>
                        		<td>
                        		  <select name="returnmethod">
<% if ((oordermaster.FOneItem.Fipkumdiv < 4) or (i = 0)) then %>
                                            <option value="">환불없음</option>
<% end if %>
<% if (divcd = "20") then %>
        <% if ((oordermaster.FOneItem.Faccountdiv = "100") and (totalrequestrepay = 0)) then %>
                                            <option value="creditcard">신용카드 취소</option>
        <% elseif ((oordermaster.FOneItem.Faccountdiv = "20") and (totalrequestrepay = 0)) then %>
                                            <option value="realtimetransfer">실시간이체 취소</option>
        <% elseif ((oordermaster.FOneItem.Faccountdiv = "30") and (totalrequestrepay = 0)) then %>
                                            <option value="point">포인트 취소</option>
        <% elseif ((oordermaster.FOneItem.Faccountdiv = "50") and (totalrequestrepay = 0)) then %>
                                            <option value="mall">입점몰결제 취소</option>
        <% elseif ((oordermaster.FOneItem.Faccountdiv = "80") and (totalrequestrepay = 0)) then %>
                                            <option value="allatcard">All@카드결재 취소</option>
        <% elseif ((oordermaster.FOneItem.Faccountdiv = "90") and (totalrequestrepay = 0)) then %>
                                            <option value="ticket">상품권결제 취소</option>
        <% end if %>
<% end if %>
                                            <option value="bank">무통장 입금</option>
                        		  </select>
                        		</td>
                        		<td></td>
                        	</tr>
                        </table>
                    </td>
                    <td width="10"></td>
                    <td valign="top">
	                    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="BABABA">
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">계좌번호</td>
                        		<td><input type="text" name="rebankaccount" value=""></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">예금주명</td>
                        		<td><input type="text" name="rebankownername" value=""></td>
                        	</tr>
                                <tr bgcolor="FFFFFF">
                        		<td height="30">거래은행</td>
                        		<td><% DrawBankCombo "rebankname", "" %></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">&nbsp;</td>
                        		<td></td>
                        	</tr>
                        	<tr bgcolor="FFFFFF">
                        		<td height="30">&nbsp;</td>
                        		<td></td>
                        	</tr>
                        </table>
                    </td>
                </tr>
            </table>

	    </td>
	    <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>

    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>

<% if (directBankCancelAvail) and (oordermaster.FOneItem.FAccountDiv="20") then %>
<script language='javascript'>
function getOnload(){
    alert('이니페이 실시간 이체의 경우 실시간이체 취소 가능. : 개편 후 변경사항.');
}
window.onload=getOnload;
</script>
<% end if %>
<%

set oordermaster = Nothing
set oorderdetail = Nothing

%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->


