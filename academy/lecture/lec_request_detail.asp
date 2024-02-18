<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
.에러 생성 - 사용못하게 함.
dim orderserial

orderserial = request("orderserial")

'==============================================================================
dim ojumun
set ojumun = new CRequestLecture

ojumun.FRectOrderSerial = orderserial
ojumun.GetRequestLectureMasterOne


'==============================================================================
dim ojumundetail
set ojumundetail = new CRequestLecture

ojumundetail.FRectOrderSerial = orderserial
ojumundetail.CRequestLectureDetailList


'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = ojumun.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture


'==============================================================================
dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = ojumun.FOneItem.Fitemid
if (olecschedule.FRectIdx = "") then
    olecschedule.FRectIdx = "0"
end if

olecschedule.GetOneLecSchedule


'==============================================================================
dim ocsaslist
set ocsaslist = New CCSASList

if (orderserial = "") then
    ocsaslist.FRectOrderSerial = "-"
else
    ocsaslist.FRectOrderSerial = orderserial
end if
ocsaslist.GetCSASMasterList


'==============================================================================
dim ocsmemo
set ocsmemo = New CCSMemo

if (ojumun.FOneItem.FUserID <> "") then
        ocsmemo.FRectUserID = ojumun.FOneItem.FUserID
elseif (orderserial <> "") then
        ocsmemo.FRectOrderserial = orderserial
else
        ocsmemo.FRectUserID = "-"
end if

ocsmemo.GetCSMemoList


'==============================================================================
dim ix, i

%>





<script>

function misendmaster(v){
	var popwin = window.open("/cscenter/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cs_action_list(v){
	var popwin = window.open("/cscenter/action/cs_action.asp?orderserial=" + v,"cs_action_list","width=1000 height=700 scrollbars=yes resizable=yes");
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
    if (1 > 0) {
        alert("작업중입니다.");
        return;
    }
	var popwin = window.showModalDialog("/cscenter/ordermaster/order_buyer_info.asp?orderserial=" + v,"order_buyer_info","resizable:no; scroll:no; dialogWidth:250px; dialogHeight:270px");
	popwin.focus();
}


// ============================================================================
// 신청정보수정
function PopOpenModifyDetail(orderserial){
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/academy/lecture/lec_modify_detail.asp?orderserial=" + orderserial,"PopOpenModifyDetail","width=500 height=250 scrollbars=no resizable=no");
	popwin.focus();
}



// ============================================================================
// 신청정보수정
function MakeNextState(orderserial){
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

<% if ((ojumun.FOneItem.Fipkumdiv = "2") and (ojumun.FOneItem.Faccountdiv = "7") and (ojumun.FOneItem.Fcancelyn = "N")) then %>
    if (confirm("결재완료 전환하시겠습니까?") == true) {
    	var popwin = window.open("/academy/lecture/lec_donextstate.asp?orderserial=" + orderserial,"MakeNextState","width=50 height=50 scrollbars=no resizable=no");
    	popwin.focus();
    }
<% else %>
    alert("정상주문중 무통장에 대한 결재완료 전환만 가능합니다.");
    return;
<% end if %>
}


// ============================================================================
// 주문메일재발송
function ReSendMail(orderserial){
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

    if (confirm("주문메일을 재발송 하시겠습니까?") == true) {
    	var popwin = window.open("/academy/lecture/lec_doresendmail.asp?orderserial=" + orderserial,"ReSendMail","width=50 height=50 scrollbars=no resizable=no");
    	popwin.focus();
    }
}


// ============================================================================
// CS등록관련

function PopOpenCancelOrder(orderserial){
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/academy/cs/pop_lec_cancel.asp?divcd=20&orderserial=" + orderserial,"PopOpenCancelOrder","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenCancelItem(orderserial){
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/academy/cs/pop_lec_cancel_detail.asp?divcd=21&orderserial=" + orderserial,"PopOpenCancelItem","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenReValidateOrder(orderserial){
	if (orderserial == "") {
        alert("먼저 주문을 선택하세요.");
        return;
    }

<% if ((ojumun.FOneItem.Fipkumdiv = "2") and (ojumun.FOneItem.Faccountdiv = "7") and (ojumun.FOneItem.Fcancelyn = "Y")) then %>
	var popwin = window.open("/academy/cs/pop_lec_revalidate.asp?divcd=22&orderserial=" + orderserial,"PopOpenReValidateOrder","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
<% else %>
    alert("취소된 주문중 무통장 주문접수상태의 신청에 대해서만 정상전환이 가능합니다.");
    return;
<% end if %>
}




function PopOpenCancelCard(orderserial){
    if (1 > 0) {
        alert("작업중입니다.");
        return;
    }
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/academy/cs/pop_lec_repay.asp?divcd=7&orderserial=" + orderserial,"PopOpenCancelCard","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenCancelBank(orderserial){
    alert("작업중입니다.");
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/academy/cs/pop_lec_repay.asp?divcd=3&orderserial=" + orderserial,"PopOpenCancelBank","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenCancelOtherSite(orderserial){
    if (1 > 0) {
        alert("작업중입니다.");
        return;
    }
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_repay.asp?divcd=5&orderserial=" + orderserial,"PopOpenCancelOtherSite","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenReadMe(orderserial){
    if (1 > 0) {
        alert("작업중입니다.");
        return;
    }
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_etc.asp?divcd=6&orderserial=" + orderserial,"PopOpenReadMe","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopOpenEtcNote(orderserial){
    if (1 > 0) {
        alert("작업중입니다.");
        return;
    }
	if (orderserial == "") {
	        alert("먼저 주문을 선택하세요.");
	        return;
        }
	var popwin = window.open("/cscenter/action/pop_cs_write_etc.asp?divcd=9&orderserial=" + orderserial,"PopOpenReadMe","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<script>
function GotoHistoryMemoModify(id) {
	var popwin = window.open("/academy/cs/pop_cs_memo_write.asp?id=" + id + "&backwindow=" + "opener","GotoHistoryMemoModify","width=400 height=250 scrollbars=no resizable=no");
	popwin.focus();
}

function GotoHistoryMemoWrite(userid, orderserial) {
        if ((userid != "") || (orderserial != ""))  {
        	var popwin = window.open("/academy/cs/pop_cs_memo_write.asp?userid=" + userid + "&orderserial=" + orderserial + "&backwindow=" + "opener","GotoHistoryMemoWrite","width=400 height=250 scrollbars=no resizable=no");
        	popwin.focus();
        }
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
	    <td>
              <a href="javascript:PopOpenCancelOrder('<%= ojumun.FOneItem.FOrderSerial %>')"><img src="/images/cs_icon_01.gif" align="absbottom" border="0"></a> <!-- 전체취소 -->
              <a href="javascript:PopOpenCancelItem('<%= ojumun.FOneItem.FOrderSerial %>')"><img src="/images/cs_icon_02.gif" align="absbottom" border="0"></a> <!-- 부분취소 -->
              <a href="javascript:PopOpenReValidateOrder('<%= ojumun.FOneItem.FOrderSerial %>')">[취소주문정상화]</a> <!-- 취소주문정상화 -->
              &nbsp;|&nbsp;
              <a href="javascript:PopOpenCancelCard('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_07.gif" align="absbottom" border="0"></a> <!-- 신용카드/상품권/실시간이체취소요청 -->
              <a href="javascript:PopOpenCancelBank('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_08.gif" align="absbottom" border="0"></a> <!-- 계좌환불요청 -->
              <a href="javascript:PopOpenCancelOtherSite('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_09.gif" align="absbottom" border="0"></a> <!-- 외부몰환불요청 -->
              &nbsp;|&nbsp;
              <a href="javascript:PopOpenReadMe('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_10.gif" align="absbottom" border="0"></a> <!-- 배송유의사항 -->
              <a href="javascript:PopOpenEtcNote('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_11.gif" align="absbottom" border="0"></a> <!-- 기타사항등록 -->
              &nbsp;|&nbsp;
              <a href="javascript:ReSendMail('<%= ojumun.FOneItem.FOrderSerial %>');"><img src="/images/cs_icon_12.gif" align="absbottom" border="0"></a> <!-- 주문메일재발송 -->
              <img src="/images/cs_icon_13.gif" align="absbottom" border="0"> <!-- 영수증재출력 -->
	    </td>
	</tr>
</table>
<br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr>
    <td valign="top">
<% if (ojumun.FOneItem.Fsitename <> "diyitem") then %>
			<!-- 신청강좌 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강좌명</td>
			    <td><%= olecture.FOneItem.Flec_title %>
			    <% if (ojumundetail.FResultCount>0) then %>
			        &nbsp;&nbsp;&nbsp;(<%= ojumundetail.FITemList(0).Fitemoptionname %>)
			    <% end if %>
			    </td>
			    <td bgcolor="#DDDDFF">강좌코드</td>
			    <td><%= ojumun.FOneItem.Fitemid %></td>
			    <td rowspan="4" width="100"><img src="<%= olecture.FOneItem.Flistimg %>" width="100"></td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강사명</td>
			    <td><%= olecture.FOneItem.Flecturer_name %>(<%= olecture.FOneItem.Flecturer_id %>)</td>
			    <td width="100" bgcolor="#DDDDFF">신청상태</td>
			    <td width="120">
			      <font color="<%= ojumun.FOneItem.CancelYnColor %>"><b><%= ojumun.FOneItem.CancelYnName %></b></font>
			      /
			      <font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>
			    </td>
			  </tr>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">강의시작일</td>
			    <td><%= Left(olecture.FOneItem.Flec_startday1, 10) %>

			    </td>
			    <td width="100" bgcolor="#DDDDFF">취소가능여부</td>
			    <td width="120">
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
			    <td width="120">
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
			    <td width="120" colspan="2">
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
			    <td width="120" colspan="2">
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
			    <td width="120" colspan="2"><%= olecture.FOneItem.Fmileage %> (point)</td>
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
			<input type="button" value="신청인원수정" onClick="PopOpenModifyDetail('<%= ojumun.FOneItem.FOrderSerial %>');"><br>
			<img src="/images/blank.gif" width="0" height="5"><br>
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
			  <tr bgcolor="#FFFFFF">
			    <td height="25" width="100" bgcolor="#DDDDFF">신청인원</td>
			    <td>
<% for i = 0 to ojumundetail.FResultCount - 1 %>
    <% if (ojumundetail.FItemList(i).Fcancelyn <> "N") then %>
                  <font color="<%= ojumundetail.FItemList(i).CancelStateColor %>"><%= ojumundetail.FItemList(i).Fentryname %>(<%= ojumundetail.FItemList(i).Fentryhp %>/<%= ojumundetail.FItemList(i).CancelStateStr %>)</font>
    <% else %>
                  <%= ojumundetail.FItemList(i).Fentryname %>(<%= ojumundetail.FItemList(i).Fentryhp %>/<%= ojumundetail.FItemList(i).CancelStateStr %>)
    <% end if %>
<% next %>
			    </td>
			  </tr>
			</table>
			<!-- 신청인원 정보 -->
<% else %>
			<table width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
			    <tr>
			        <td>
			            <table width="100%" border="0" cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
			                <form name="frm" method="get" action="">
			                <input type="hidden" name="orderserial" value="<%= orderserial %>">
			                <input type="hidden" name="research" value="on">
                            <tr>
                                <td height="1" colspan="15" bgcolor="#BABABA"></td>
                            </tr>
				            <tr align="center" style="padding:2">
			                	<td width="30">구분</td>
			                	<td width="50">진행상태</td>
			                	<td width="40">CODE</td>
			                  	<td width="50">이미지</td>
			                    <td width="120">브랜드ID</td>
			                	<td>상품명<font color="blue">[옵션명]</font></td>
			                	<td width="30">수량</td>
			                	<td width="50">현재<br>소비자가</td>
			                	<td width="50">판매가</td>
			                	<td width="70">확인일</td>
			                	<td width="70">출고일</td>
			                	<td width="125">배송정보</td>
			                </tr>
			                <% for ix=0 to ojumundetail.FResultCount-1 %>
			                <% if ojumundetail.FItemList(ix).Fitemid =0 then %>

			                <% if ojumundetail.FItemList(ix).FCancelyn ="Y" then %>
			                <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
			                <% else %>
			                <tr align="center" height="25">
			                <% end if %>
			                    <td width="30"></td>
			                    <td width="50"></td>
			                	<td width="40">0</td>
			                   	<td width="50">
			                   	<!--
			                   	    <input type="checkbox" name="onimage" <% if onimage="on" then response.write "checked" %> onclick="javascript:document.frm.submit();" >
			                   	-->
			                   	</td>
			                	<td width="120" align="left"><%= ojumundetail.FItemList(ix).FMakerid %></td>
			                	<td align="left">배송비<font color="blue">[<%= ojumundetail.BeasongCD2Name(ojumundetail.FItemList(ix).Fitemoption) %>]</font></td>
			                	<td width="30"></td>
			                	<td width="50"></td>
			                	<td width="50" align="right"><font color="blue"><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></font></td>
			                	<td width="70"></td>
			                	<td width="70"></td>
			                	<td width="108"></td>
			                </tr>
			                <% end if %>
			                <% next %>
			                <tr>
			            		<td height="1" colspan="12" bgcolor="#BABABA"></td>
			            	</tr>
			                <% for ix=0 to ojumundetail.FResultCount-1 %>
			                <% if ojumundetail.FItemList(ix).Fitemid <>0 then %>

			                <% if ojumundetail.FItemList(ix).FCancelyn ="Y" then %>
			                <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
			                <% else %>
			                <tr align="center" height="25">
			                <% end if %>

			                    <td><font color="<%= ojumundetail.FItemList(ix).CancelStateColor %>"><%= ojumundetail.FItemList(ix).CancelStateStr %></font></td>
			                    <td><font color="<%= ojumundetail.FItemList(ix).GetStateColor %>"><%= ojumundetail.FItemList(ix).GetStateName %></font></td>
			                	<td>
			                	    <% if ojumundetail.FItemList(ix).Fisupchebeasong="Y" then %>
			                	    <a href="javascript:popOrderDetailEdit('<%= ojumundetail.FItemList(ix).Fdetailidx %>');"><font color="red"><%= ojumundetail.FItemList(ix).Fitemid %><br>(업체)</font></a>
			                        <% else %>
			                        <a href="javascript:popOrderDetailEdit('<%= ojumundetail.FItemList(ix).Fdetailidx %>');"><%= ojumundetail.FItemList(ix).Fitemid %></a>
			                        <% end if %>
			                    </td>
			                    <td align="center">
			                        &nbsp;
			                    </td>
			                    <td align="left">
			                        <a href="javascript:popSimpleBrandInfo('<%= ojumundetail.FItemList(ix).Fmakerid %>');">
			                        <acronym title="<%= ojumundetail.FItemList(ix).Fmakerid %>"><%= Left(ojumundetail.FItemList(ix).Fmakerid,12) %></acronym>
			                        </a>
			                    </td>
			                	<td align="left">
			                	    <acronym title="<%= ojumundetail.FItemList(ix).FItemName %>"><%= Left(ojumundetail.FItemList(ix).FItemName,35) %></acronym>
			                	    <% if ojumundetail.FItemList(ix).FItemoption<>"0000" then %>
				                	    <br>
				                	    <a href="javascript:popOrderDetailEditOption('<%=ojumundetail.FItemList(ix).Fdetailidx%>');"><font color="blue"><%= ojumundetail.FItemList(ix).FItemoptionName %></font></a>
			                	    <% end if %>
			                	    <% if ojumundetail.FItemList(ix).IsRequireDetailExistsItem then %>
			                	    	<br>
			                	    	<a href="javascript:EditRequireDetail('<%= orderserial %>','<%= ojumundetail.FItemList(ix).Fdetailidx%>')"><font color="red">[주문제작상품]</font>
			                	    	<br>
			                	    	<%= db2html(ojumundetail.FItemList(ix).getRequireDetailHtml) %>
			                	    	</a>
			                	    <% end if %>
			                	</td>

			                	<% if ojumundetail.FItemList(ix).FItemNo > 1 then %>
			                	<td><b><font color="blue"><%= ojumundetail.FItemList(ix).FItemNo %></font></b></td>
			                	<% elseif ojumundetail.FItemList(ix).FItemNo < 1 then %>
			                	<td><b><font color="red"><%= ojumundetail.FItemList(ix).FItemNo %></font></b></td>
			                	<% else %>
			                	<td><font color="blue"><%= ojumundetail.FItemList(ix).FItemNo %></font></td>
			                	<% end if %>

			                    <td align="right">--</td> <!-- 소비자가 -->

			                   	<% if ojumundetail.FItemList(ix).FItemNo < 1 then %>
			                   	<td align="center"><font color="red"><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></font></td>
			                   	<% else %>
			                   	<td align="right">
			                   	    <% if ojumundetail.FItemList(ix).Fissailitem="Y" then %>
			                   	    <span title="세일상품" style="cursor:hand"><font color="red"><b><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></b></font></span>
			                   	    <% elseif ojumundetail.FItemList(ix).Fissailitem="P" then %>
			                   	    <span title="플러스세일상품" style="cursor:hand"><font color="purple"><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></font></span>
			                   	    <% elseif ojumundetail.FItemList(ix).IsBonusCouponDiscountItem then %>
			                   	    <span title="비율보너스쿠폰할인상품" style="cursor:hand">
			                   	    <font color="blue">
			                   	        <%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %><br>
			                   	        <font color="#000000">(<%= FormatNumber(ojumundetail.FItemList(ix).FreducedPrice,0) %>)</font>
			                   	    </font>
			                   	    </span>
			                   	    <% elseif ojumundetail.FItemList(ix).IsItemCouponDiscountItem then %>
			                   	    <span title="상품보너스쿠폰할인상품" style="cursor:hand"><font color="green"><b><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></b></font></span>
		                   	    <% else %>
			                   	    <span title="정상가격" style="cursor:hand"><font color="#000000"><%= FormatNumber(ojumundetail.FItemList(ix).Fitemcost,0) %></font></span>
			                   	    <% end if %>
			                   	</td>
			                   	<% end if %>


			                	<td><acronym title="<%= ojumundetail.FItemList(ix).Fupcheconfirmdate %>"><%= Left(ojumundetail.FItemList(ix).Fupcheconfirmdate,10) %></acronym></td>
			                	<td><acronym title="<%= ojumundetail.FItemList(ix).Fbeasongdate %>"><%= Left(ojumundetail.FItemList(ix).Fbeasongdate,10) %></acronym></td>
			                	<td>
			                	    <%= ojumundetail.FItemList(ix).Fsongjangdivname %><br>
			                		<% if (ojumundetail.FItemList(ix).FsongjangDiv="24") then %>
		                		<a href="javascript:popDeliveryTrace('<%= ojumundetail.FItemList(ix).Ffindurl %>','<%= ojumundetail.FItemList(ix).Fsongjangno %>');"><%= ojumundetail.FItemList(ix).Fsongjangno %></a>
			                	    <% else %>
			                	    <a target="_blank" href="<%= ojumundetail.FItemList(ix).Ffindurl + ojumundetail.FItemList(ix).Fsongjangno %>"><%= ojumundetail.FItemList(ix).Fsongjangno %></a>
			                	    <% end if %>
			                	</td>
			                </tr>
			                <tr>
			            		<td height="1" colspan="15" bgcolor="#BABABA"></td>
			            	</tr>
			                <% end if %>
			                <% next %>


			                </form>
			            </table>
			        </td>
			    </tr>
			</table>
<% end if %>
            <br>
        	<!-- 상세 정보 -->
        	<table width="100%" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
        		<tr height="10" valign="bottom">
        		    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        		    <td background="/images/tbl_blue_round_02.gif"></td>
        		    <td background="/images/tbl_blue_round_02.gif"></td>
        		    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
        		</tr>
        		<tr height="25">
        		    <td background="/images/tbl_blue_round_04.gif"></td>
        		    <td>
        		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS 정보</b>
        		    </td>
        		    <td>
        		    </td>
        		    <td background="/images/tbl_blue_round_05.gif"></td>
        		</tr>
        	</table>
            <table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
                <tr bgcolor="#DDDDFF">
                    <td width="50" align="center">Idx</td>
                    <td width="75" align="center">구분</td>
                    <td align="center">제목</td>
                    <td width="60" align="center">상태</td>
                    <td width="70" align="center">환불<br>요청액</td>
                    <td width="80" align="center">등록일</td>
                    <td width="80" align="center">처리일</td>
                    <td width="30" align="center">삭제여부</td>
                </tr>
            <% for i = 0 to (ocsaslist.FResultCount - 1) %>
                <tr bgcolor="#FFFFFF" align="center" <% if (ocsaslist.FItemList(i).Fdeleteyn = "Y") then %>style="color:gray"<% end if %>>
                    <td height="25" nowrap><%= ocsaslist.FItemList(i).Fid %></td>
                    <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).FdivcdName %>"><%= Left(ocsaslist.FItemList(i).FdivcdName,6) %></acronym></td>
                    <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).Ftitle %>"><%= ocsaslist.FItemList(i).Ftitle %></acronym></td>
                    <td nowrap><%= ocsaslist.FItemList(i).Fcurrstatename %></td>
                    <td nowrap align="right"><%= FormatNumber(ocsaslist.FItemList(i).Frefundrequire,0) %></td>
                    <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                    <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
                    <td nowrap><%= ocsaslist.FItemList(i).Fdeleteyn %></td>
                </tr>
            <% next %>
            <% if (ocsaslist.FResultCount < 1) then %>
                <tr>
                  <td height="25" colspan="8" align="center" bgcolor="#FFFFFF">등록된 AS 가 없습니다.</td>
                </tr>
            <% end if %>
            </table>
        	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
        		<tr>
        			<td background="/images/tbl_blue_round_04.gif"></td>
        		    <td>&nbsp;</td>
        		    <td></td>
        		    <td background="/images/tbl_blue_round_05.gif"></td>
        		</tr>
        		<tr height="10" valign="top">
        			<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        		    <td background="/images/tbl_blue_round_08.gif"></td>
        		    <td background="/images/tbl_blue_round_08.gif"></td>
        		    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
        		</tr>
        	</table>
        	<!-- 상세 정보 -->

        	<br>

        	<!-- 상세 정보 -->
        	<table width="100%" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
        		<tr height="10" valign="bottom">
        		    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        		    <td background="/images/tbl_blue_round_02.gif"></td>
        		    <td background="/images/tbl_blue_round_02.gif"></td>
        		    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
        		</tr>
        		<tr height="25">
        		    <td background="/images/tbl_blue_round_04.gif"></td>
        		    <td>
        		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS 메모<% if (ojumun.FOneItem.FUserID <> "") then %>[고객아이디 : <%= ojumun.FOneItem.FUserID %>]<% else %>[비회원주문 : <%= orderserial %>]<% end if %></b>
        		    </td>
        		    <td align="right">
        		      <input type="button" value="메모작성" onClick="GotoHistoryMemoWrite('<%= ojumun.FOneItem.FUserID %>','<%= orderserial %>')">
        		    </td>
        		    <td background="/images/tbl_blue_round_05.gif"></td>
        		</tr>
        	</table>
<table width="100%" border=0 cellspacing=1 cellpadding=2 class=a bgcolor="BABABA">

    <tr align="center" bgcolor="F3F3FF">
        <td height="20" width="50">구분</td>
    	<td width="50">idx</td>
     	<td width="50">사이트</td>
    	<td width="80">주문번호</td>
    	<td>내용</td>
        <td width="80">등록자</td>
    	<td width="80">등록일</td>
    	<td width="30">완료</td>
    </tr>
<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="FFFFFF">
        <td height="20"><%= ocsmemo.FItemList(i).GetDivCDName %></td>
    	<td><%= ocsmemo.FItemList(i).Fid %></td>
     	<td><%= ocsmemo.FItemList(i).GetSiteName %></td>
    	<td><%= ocsmemo.FItemList(i).Forderserial %></td>
    	<td align="left"><a href="javascript:GotoHistoryMemoModify(<%= ocsmemo.FItemList(i).Fid %>)"><%= DDotFormat(ocsmemo.FItemList(i).Fcontents_jupsu,35) %></a></td>
        <td><%= ocsmemo.FItemList(i).Fwriteuser %></td>
    	<td><acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><% if (ocsmemo.FItemList(i).Ffinishyn = "Y") then %>완료<% end if %></td>
    </tr>
<% next %>
<% if (ocsmemo.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="20" colspan="9">등록된 메모가 없습니다.</td>
    </tr>
<% end if %>
</table>
        	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
        		<tr>
        			<td background="/images/tbl_blue_round_04.gif"></td>
        		    <td>&nbsp;</td>
        		    <td></td>
        		    <td background="/images/tbl_blue_round_05.gif"></td>
        		</tr>
        		<tr height="10" valign="top">
        			<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        		    <td background="/images/tbl_blue_round_08.gif"></td>
        		    <td background="/images/tbl_blue_round_08.gif"></td>
        		    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
        		</tr>
        	</table>
        	<!-- 상세 정보 -->


    </td>
    <td width="10"></td>
    <td width="260" valign="top">


			<!-- 구매자정보 -->
			<table width="250" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="10" valign="bottom">
				    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
				</tr>
				<tr height="25">
				    <td background="/images/tbl_blue_round_04.gif"></td>
				    <td background="/images/tbl_blue_round_06.gif">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매자 정보</b>
				    </td>
				    <td align="right" background="/images/tbl_blue_round_06.gif">
				        <input type="button" value="정보수정" onclick="javascript:order_buyer_info('<%= ojumun.FOneItem.FOrderSerial %>');">
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
			</table>
			<table width="250" height="185" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>구매자ID</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FUserID %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>구매자명</td>
				    <td>
				        <input type="text" value="<%= ojumun.FOneItem.FBuyName %>" size="8" readonly>
				        현재등급
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>전화번호</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FBuyPhone %>" readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>핸드폰</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FBuyHp %>" readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>이메일</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FBuyEmail %>" readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>입금자명</td>
				    <td>
				        <input type="text" value="<%= ojumun.FOneItem.FAccountName %>" size="14" readonly>
				        <input type="button" value="검색" onclick="alert('작업중입니다.');">
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr>
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td></td>
				    <td></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="10" valign="top">
					<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
				</tr>
			</table>
			<!-- 구매자정보 -->
<% if (ojumun.FOneItem.Fsitename = "diyitem") then %>
			<br>
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
				        <input type="button" name="reqhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend('<%= ojumun.FOneItem.FReqHp %>','<%= ojumun.FOneItem.Forderserial %>','<%= ojumun.FOneItem.Fuserid %>','');">
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
<% end if %>
            <br>
		    <!-- 주문정보 -->
			<table width="250" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="10" valign="bottom">
				    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
				</tr>
				<tr height="25">
				    <td background="/images/tbl_blue_round_04.gif"></td>
				    <td background="/images/tbl_blue_round_06.gif">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문 정보</b>
				    </td>
				    <td align="right" background="/images/tbl_blue_round_06.gif"><input type="button" value="다음상태진행" onClick="MakeNextState('<%= ojumun.FOneItem.FOrderSerial %>')"></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
			</table>
			<table width="250" height="185" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>주문번호</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FOrderSerial %>" size="11" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>결제방법</td>
				    <td>
				        <input type="text" value="<%= ojumun.FOneItem.JumunMethodName %>" size="10" style='background-color:#DDDDFF' readonly>
				        <input type="text" value="<%= ojumun.FOneItem.IpkumDivName %>" size="8" style='background-color:#DDDDFF;color:<%= ojumun.FOneItem.IpkumDivColor %>' readonly>
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>신청일</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FRegDate %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>입금일</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FIpkumDate %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>카드승인</td>
				    <td>
				        <input type="text" value="<%= ojumun.FOneItem.FAuthcode %>" size="8" style='background-color:#DDDDFF' readonly>
				        정산유무
				        <input type="text" value="<%= ojumun.FOneItem.Fjungsanflag %>" size="2" style='background-color:#DDDDFF' readonly>
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>카드승인</td>
				    <td><input type="text" value="<%= ojumun.FOneItem.FPaygatetID %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>결제결과</td>
				    <td><%= ojumun.FOneItem.Fresultmsg %></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr>
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td></td>
				    <td></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="10">
					<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
				</tr>
			</table>
			<!-- 주문정보 -->
    </td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td></td>
  </tr>
</table>





<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->