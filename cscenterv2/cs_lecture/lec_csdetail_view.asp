<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 고객센터 강좌CS처리 리스트
' Hieditor : 2015.05.27 이상구 생성
'			 2017.07.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/cs_lecture/lec_cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<%
dim i
dim id, divcd, currstate
id = RequestCheckvar(request("id"),10)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if

dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo

	if (orefund.FOneItem.Fencmethod = "TBT") then
		''orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
	elseif (orefund.FOneItem.Fencmethod = "PH1") then
	    orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
	end if

	if DateDiff("m", ocsaslist.FOneItem.Fregdate, Now) > 3 then
		if (orefund.FOneItem.Frebankaccount <> "") then
			orefund.FOneItem.Frebankaccount = ""
			orefund.FOneItem.Frebankownername = ""
			orefund.FOneItem.Frebankname = "<font color='red'>3개월경과(계좌정보 표시안함)</font>"
		else
			orefund.FOneItem.Frebankaccount = ""
			orefund.FOneItem.Frebankownername = ""
			orefund.FOneItem.Frebankname = ""
		end if
	end if
end if


dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount > 0) then
    oordermaster.FRectOrderSerial = ocsaslist.FOneItem.Forderserial
    oordermaster.QuickSearchOrderMaster

    divcd = ocsaslist.FOneItem.FDivCD
    currstate = ocsaslist.FOneItem.Fcurrstate
end if

if (oordermaster.FResultCount<1) and (Len(oordermaster.FRectOrderSerial)=11) and (IsNumeric(oordermaster.FRectOrderSerial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


dim OCsDetail
set OCsDetail = new CCSASList
OCsDetail.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsDetail.GetCsDetailList
end if


dim OCsDelivery
set OCsDelivery = new CCSASList
OCsDelivery.FRectCsAsID = id
if ocsaslist.FResultCount>0 then
    OCsDelivery.GetOneCsDeliveryItem
end if


''반품주소지 : requireupche : Y && makerid <>''
dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
if (ocsaslist.FResultCount>0) then
    if (ocsaslist.FOneItem.Frequireupche="Y") then
        OReturnAddr.FRectMakerid = ocsaslist.FOneItem.FMakerid
        'OReturnAddr.GetReturnAddress
    end if
end if

''확인요청정보 :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
    OCsConfirm.GetOneCsConfirmItem
end if

''업체 추가 정산
dim IsUpCheAddJungsanDisplay
if (id<>"") then
    IsUpCheAddJungsanDisplay = (ocsaslist.FOneItem.Fdivcd="A004") or (ocsaslist.FOneItem.Fdivcd="A700") ''반품접수, 업체 기타정산
end if
%>
<script language='javascript'>
function PopCSMailTest(iid){
    var popwin = window.open('cs_action_mail_view.asp?id=' + iid,'cs_action_mail_view','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function CardCancelProcess(iid){
    var popwin = window.open('pop_CardCancel.asp?id=' + iid,'PopCardCancelProcess','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regConfirmMsg(iid,fin){
    var popwin = window.open('pop_ConfirmMsg.asp?id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopCSAddUpchejungsanEdit(iid){
    var popwin = window.open('pop_AddUpchejungsanEdit.asp?id=' + iid ,'AddUpchejungsanEdit','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<% if ocsaslist.FResultCount>0 then %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr height="30">
		<td>
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center"bgcolor="#E6E6E6">
					<td <% if ocsaslist.FOneItem.Fcurrstate="B001" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[1]접수</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B002" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[2]미처리(사유)</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B003" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[3]택배사전송</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B004" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[4]운송장입력</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B005" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[5]확인요청</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B006" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[6]업체처리완료</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B007" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[7]완료</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B012" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[12]회수미처리(현대)</td>
					<td <% if ocsaslist.FOneItem.Fcurrstate="B013" then %> bgcolor="<%= adminColor("pink") %>" <% end if %> >[13]맞교환회수미처리(현대)</td>
				</tr>
			</table>
		</td>
	</tr>
</table>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frmdetail" onsubmit="return false;">
	<input type="hidden" name="id" value="<%= id %>">
	<tr valign="top" height="600">
		<td>
			<!-- 접수 정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="4">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>접수 정보</b>
				    		    	&nbsp;[<%= ocsaslist.FOneItem.GetAsDivCDName %>]
				    		    	&nbsp;<%= ocsaslist.FOneItem.Forderserial %>
				    		    </td>
				    		    <td align="right" >
				    		    <input class="button" type="button" value="CSmail" onclick="javascript:PopCSMailTest('<%= id %>');" >

				    		        <input class="button" type="button" value="정보수정" onclick="javascript:PopCSActionEdit_Lecture('<%= id %>','editreginfo');" >
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<tr>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">접수자</td>
				    <td width="80" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fwriteuser %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">접수일시</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fregdate %></td>
				</tr>
				<tr height="20">
				    <td bgcolor="<%= adminColor("topbar") %>">제목</td>
				    <td colspan="3" bgcolor="#F4F4F4"><input type="text" class="text_ro" name="title" value="<%= ocsaslist.FOneItem.FTitle %>" size="68" maxlength="60" ReadOnly></td>
				</tr>
				<tr bgcolor="#F4F4F4">
				    <td bgcolor="<%= adminColor("topbar") %>">사유구분</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<%= ocsaslist.FOneItem.GetCauseString() %> > <%= ocsaslist.FOneItem.GetCauseDetailString %>
				    </td>
				</tr>
				<tr bgcolor="#F4F4F4">
				    <td bgcolor="<%= adminColor("topbar") %>">접수내용</td>
				    <td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea_ro" name="contents_jupsu" cols="68" rows="8" ReadOnly><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea></td>
				</tr>
			</table>
			<!-- 접수 정보 -->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>
			<!-- 접수시 상품정보 시작-->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
			        <td>
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품 정보</b> (접수 된 내역)
				    		    </td>
				    		    <td align="right" >
				    		    <!-- ?
				    		    	<input class="button" type="button" value="이전CS 상품코드로 등록" onclick="" >
				    		        <input class="button" type="button" value="상세보기" onclick="alert('?');" >
				    		     -->
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<tr valign="top" bgcolor="<%= adminColor("topbar") %>">
				   	<td>
				   		<table width="100%" height="200" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			            	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
			            		<td style="width:30px; border-right:1px solid <%= adminColor("tablebg") %>;">구분</td>
				    		    <td style="width:60px; border-right:1px solid <%= adminColor("tablebg") %>;">접수시<br>진행상태</td>
				    		    <td style="width:40px; border-right:1px solid <%= adminColor("tablebg") %>;">CODE</td>
				    		    <td style="border-right:1px solid <%= adminColor("tablebg") %>;">상품명[옵션]</td>
				    		    <td style="width:50px; border-right:1px solid <%= adminColor("tablebg") %>;">판매가</td>
				    		    <td style="width:30px;">수량</td>
				    		</tr>
				    		<tr>
                                <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
                            </tr>
                            <% for i=0 to OCsDetail.FResultCount-1 %>
                            <tr height="25" align="center" bgcolor="#FFFFFF" >
				    			<td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).GetRegDetailStateName %></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemid %></td>
				    		    <td align="left" style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemname %>[<%= OCsDetail.FItemList(i).Fitemoptionname %>]</td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fitemcost %></td>
				    		    <td style="border-bottom:1px solid <%= adminColor("tablebg") %>;"><%= OCsDetail.FItemList(i).Fregitemno %></td>
				    		</tr>
                            <% next %>
                            <tr bgcolor="#FFFFFF">
                                <td colspan="6"></td>
                            </tr>
		    		    </table>
		    		    <!--
		    		    <table height="176" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                            <tr height="100%">
                                <td colspan="12">
                        	        <iframe name="" src="" border=1 frameSpacing=1 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
                                </td>
                            <tr>
                        </table>
                        -->
				   	</td>
				</tr>

			</table>
			<!-- 접수시 주문정보 끝-->
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>
			<!-- 접수시 주소정보 시작-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>반품주소 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		        <% if (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A010") or (divcd="A011") or (OCsDelivery.FResultCount>0) then %>
    				    		        <% if (currstate="B001") then %>
    				    		        <input class="button" type="button" value="주소변경" onclick="popEditCsDelivery('<%= id %>');" >
    				    		        <% else %>
    				    		        <input class="button" type="button" value="주소변경불가" onclick="alert('접수상태에서만 변경가능 합니다.');" >
    				    		        <% end if %>
				    		        <% end if %>
				    		    </td>
				    		 </tr>
		    		    </table>
				    </td>
				</tr>
				<% if (OCsDelivery.FResultCount>0) then %>
				<!-- 고객 교환/회수 주소 -->
				<tr>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("pink") %>">고객주소</td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">고객명</td>
				    <td width="80" bgcolor="#FFFFFF"><%= OCsDelivery.FOneItem.Freqname %></td>
				    <td width="50" bgcolor="<%= adminColor("pink") %>">연락처</td>
				    <td bgcolor="#FFFFFF"><%= OCsDelivery.FOneItem.Freqphone %> / <%= OCsDelivery.FOneItem.Freqhp %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("pink") %>">주소</td>
				    <td colspan="3" bgcolor="#FFFFFF">[<%= OCsDelivery.FOneItem.Freqzipcode %>] <%= OCsDelivery.FOneItem.Freqzipaddr %> &nbsp;<%= OCsDelivery.FOneItem.Freqetcaddr %></td>
				</tr>
				<% else %>
				<tr>
					<td width="50" bgcolor="<%= adminColor("pink") %>">고객주소</td>
					<td colspan="4" bgcolor="#FFFFFF">주문시 배송지</td>
				</tr>
				<% end if %>
				<!-- 반품 회수 주소 -->
				<tr>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("topbar") %>">반품회수<br>주소</td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">업체명</td>
				    <td width="80" bgcolor="#FFFFFF"><%= OReturnAddr.Freturnname %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">연락처</td>
				    <td bgcolor="#FFFFFF"><%= OReturnAddr.Freturnphone %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">주소</td>
				    <td colspan="3" bgcolor="#FFFFFF">[<%= OReturnAddr.Freturnzipcode %>] <%= OReturnAddr.Freturnzipaddr %> &nbsp;<%= OReturnAddr.Freturnetcaddr %></td>
				</tr>

			</table>
			<!-- 접수시 주소정보 끝-->
		</td>

		<td width="5"></td>

		<td width="30%">
			<!-- 환불관련정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="3">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>취소관련 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		        <input class="button" type="button" value="정보수정" onclick="PopCSActionEdit_Lecture('<%= id %>','editrefundinfo');">
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<% if (orefund.FresultCount>0) then %>
				<tr>
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">상품총액</td>
				    <td width="60" bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %></td>
				    <td bgcolor="#FFFFFF">원 주문 상품 총액</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">배송료</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Forgbeasongpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>">
				    <td>주문총액</td>
				    <td align="right"></td>
				    <td></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">마일리지사용</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum*-1,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">쿠폰사용</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum*-1,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">기타할인</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Forgallatdiscountsum*-1,0) %></td>
				    <td bgcolor="#FFFFFF">국민(0.5%) 올앳 (0.6%)</td>
				</tr>
				<tr bgcolor="<%= adminColor("green") %>">
				    <td>원 결제총액</td>
				    <td align="right"><%= FormatNumber(orefund.FOneItem.Forgsubtotalprice,0) %></td>
				    <td></td>
				</tr>
				<tr>
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">취소상품금액</td>
				    <td width="60" bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefunditemcostsum,0) %></td>
				    <td bgcolor="#FFFFFF">비고</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">배송료</td>
				    <td bgcolor="#FFFFFF" align="right"><%= FormatNumber(orefund.FOneItem.Frefundbeasongpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">회수 배송비</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefunddeliverypay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">마일리지</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundmileagesum,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">쿠폰</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundcouponsum,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">기타할인</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum,0) %></td>
				    <td bgcolor="#FFFFFF">국민(0.5%) 올앳 (0.6%)</td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">보정액</td>
				    <td bgcolor="#FFFFFF"align="right"><%= FormatNumber(orefund.FOneItem.Frefundadjustpay,0) %></td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>

				<tr bgcolor="<%= adminColor("green") %>">
				    <td>환불예정액</td>
				    <td align="right"><%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %></td>
				    <td></td>
				</tr>
				<% else %>
				<tr>
				    <td colspan="3" align="center" bgcolor="#FFFFFF">[환불 정보가 없습니다.]</td>
				</tr>
				<% end if %>
			</table>
			<!-- 환불관련정보 -->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>

			<!-- 계좌정보 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>환불관련 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		        <input class="button" type="button" value="정보수정" onclick="PopCSActionEdit_Lecture('<%= id %>','editrefundinfo');">
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<% if (orefund.FresultCount>0) then %>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">취소방식 선택</td>
				    <td bgcolor="#FFFFFF">
				        <%= orefund.FOneItem.FreturnmethodName %>
				        (<%= orefund.FOneItem.Freturnmethod %>)
				    </td>
				</tr>
				<% if (orefund.FOneItem.Freturnmethod="R007") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">은행</td>
				    <td bgcolor="#FFFFFF"><%= orefund.FOneItem.Frebankname %>
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">계좌번호</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text" name="refundaccount" value="<%= orefund.FOneItem.Frebankaccount %>" maxlength="20" size="25"> (대쉬 - 빼고 입력)
				    </td>
				</tr>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">예금주</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text" name="refundaccountname" value="<%= orefund.FOneItem.Frebankownername %>" maxlength="16" size="16"> (통장 예금주 명)
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R900") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">아이디</td>
				    <td bgcolor="#FFFFFF">
				    <%if oordermaster.FResultCount>0 then %>
						<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 then %>
							(<%= oordermaster.FOneItem.FUserID %>)
						<% else %>
							(<%= printUserId(oordermaster.FOneItem.FUserID, 2, "*") %>)
						<% end if %>
				    <% end if %>
				    </td>
				</tr>
				<% elseif (orefund.FOneItem.Freturnmethod="R100") or (orefund.FOneItem.Freturnmethod="R020") or (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") or (orefund.FOneItem.Freturnmethod="R080") or (orefund.FOneItem.Freturnmethod="R400") then %>
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" name="paygateTid" value="<%= orefund.FOneItem.FpaygateTid %>" size="32" readonly>
				        <% if ocsaslist.FOneItem.FCurrState="B001" then %>
				        <input type="button" class="button" value="완료처리" onclick="CardCancelProcess('<%= ocsaslist.FOneItem.Fid %>');">
				        <% end if %>
				    </td>
				</tr>
				<% end if %>
				<tr height="25" bgcolor="<%= adminColor("green") %>">
				    <td bgcolor="<%= adminColor("topbar") %>">환불예정액</td>
				    <td bgcolor="#FFFFFF"><%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>원</td>
				</tr>
				<!--
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">PG사</td>
				    <td bgcolor="#FFFFFF"></td>
				</tr>
				-->
				<!-- 현재 승인번호 저장 안함
				<tr height="25">
				    <td bgcolor="<%= adminColor("topbar") %>">승인번호</td>
				    <td bgcolor="#FFFFFF">
				    	<input type="text" class="text_ro" value="" size="45" readonly>
				    </td>
				</tr>
                -->

				<% else %>
				<tr height="25" >
				    <td colspan="2" align="center" bgcolor="#FFFFFF">[환불 계좌 정보가 없습니다.]</td>
				</tr>
				<% end if %>
			</table>
			<!-- 계좌정보 끝-->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>


			<% if (IsUpCheAddJungsanDisplay) or (ocsaslist.FOneItem.Fadd_upchejungsandeliverypay>0) then %>
			<!-- 업체 추가 정산 -->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>업체 추가 정산 내역</b>
				    		    </td>
				    		    <td align="right" >
				    		        <input class="button" type="button" value="정보수정" onclick="PopCSAddUpchejungsanEdit('<%= id %>','editrefundinfo');">
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">추가정산액</td>
				    <td width="280" bgcolor="#FFFFFF"><%= FormatNumber(ocsaslist.FOneItem.Fadd_upchejungsandeliverypay,0) %></td>
				</tr>
				<tr height="25">
				    <td width="80" bgcolor="<%= adminColor("topbar") %>">사유</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Fadd_upchejungsancause %></td>
				</tr>
			</table>
			<!-- 업체 추가 정산 끝-->
			<% end if %>

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td></td>
				</tr>
			</table>

			<!-- 카드 등 취소관련정보
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="2">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>신용카드/실시간이체 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		    </td>
				    		</tr>
		    		    </table>
				    </td>
				</tr>

			</table>
			 -->
		</td>

		<td width="5"></td>

		<td width="30%">
			<!-- 처리 정보 시작-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="5">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>처리 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		        <input class="button" type="button" value="완료처리" onclick="PopCSActionFinish_Lecture('<%= id %>','finishreginfo');" >
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리자</td>
				    <td width="80" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishuser %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>미처리<% end if %></td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리일시</td>
				    <td bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Ffinishdate %><% if isnull(ocsaslist.FOneItem.Ffinishuser) then %>미처리<% end if %></td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">처리내용</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				        <textarea class="textarea_ro" name="contents_finish" cols="48" rows="8"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
				    </td>
				</tr>
				<!--
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
				    <td colspan="3" bgcolor="#FFFFFF"><%= ocsaslist.FOneItem.Forderserial %>_<%= id %></td>
				</tr>
				-->
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">관련송장</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
					        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
					        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
					        <a href="<%= DeliverDivTrace(Trim(ocsaslist.FOneItem.Fsongjangdiv)) %><%= ocsaslist.FOneItem.Fsongjangno %>" target="_blank">추적</a>
				            <input type="button" class="button" value="수정" onClick="changeSongjang('<%= id %>');">
				        <% end if %>
				    </td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td rowspan="2">처리관련<br>고객오픈<br>내용입력</td>
				    <td colspan="3">
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td colspan="3">
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>

			</table>
			<!-- 처리 정보 끝-->

			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr height="5">
					<td>
					</td>
				</tr>
			</table>

			<!-- 확인요청 정보 시작-->
			<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr height="25" bgcolor="<%= adminColor("topbar") %>">
			        <td colspan="4">
			            <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" class="a" >
			            	<tr>
				    		    <td>
				    		    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>확인요청 정보</b>
				    		    </td>
				    		    <td align="right" >
				    		    <% if OCsConfirm.FResultCount>0 then %>
				    		        <input class="button" type="button" value="확인요청 수정" onclick="regConfirmMsg('<%= id %>','');" >
				    		        <input class="button" type="button" value="확인요청 완료" onclick="regConfirmMsg('<%= id %>','fin');" >
				    		    <% else %>
				    		        <input class="button" type="button" value="확인요청 정보등록" onclick="regConfirmMsg('<%= id %>','');" >
				    		    <% end if %>
				    		    </td>
				    		 </tr>
		    			</table>
		    		</td>
		    	</tr>
				<tr height="23">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">등록자</td>
				    <td width="80" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmreguserid %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">등록일시</td>
				    <td bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmregdate %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">등록내용</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				    <textarea class="textarea_ro" name="confirmregmsg" cols="48" rows="5" readonly ><%= OCsConfirm.FOneItem.Fconfirmregmsg %></textarea>
				    <% else %>
				    <textarea class="textarea_ro" name="confirmregmsg" cols="48" rows="5" readonly ></textarea>
				    <% end if %>
				    </td>
				</tr>
				<tr height="23">
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리자</td>
				    <td width="80" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmfinishuserid %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				    <td width="50" bgcolor="<%= adminColor("topbar") %>">처리일시</td>
				    <td bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				        <%= OCsConfirm.FOneItem.Fconfirmfinishdate %>
				    <% else %>
				        &nbsp;
				    <% end if %>
				    </td>
				</tr>
				<tr>
				    <td bgcolor="<%= adminColor("topbar") %>">처리내용</td>
				    <td colspan="3" bgcolor="#FFFFFF">
				    <% if OCsConfirm.FResultCount>0 then %>
				    <textarea class="textarea_ro" name="confirmfinishmsg" cols="48" rows="5" readonly ><%= OCsConfirm.FOneItem.Fconfirmfinishmsg %></textarea>
				    <% else %>
				    <textarea class="textarea_ro" name="confirmfinishmsg" cols="48" rows="5" readonly ></textarea>
				    <% end if %>
				    </td>
				</tr>
				<!--
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td rowspan="2">확인요청<br>고객오픈<br>내용입력</td>
				    <td colspan="3"><input type="text" class="text" name="" value="" size="48" maxlength="60"></td>
				</tr>
				<tr bgcolor="<%= adminColor("pink") %>">
				    <td colspan="3"><textarea class="textarea" name="" cols="48" rows="5">&nbsp;</textarea></td>
				</tr>
				-->
			</table>
			<!-- 확인요청 정보 끝-->
		</td>
	</tr>
	</form>
<table>




<% else %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="50">
	    <td align="center">[ 선택된 처리AS 가 없습니다. 먼저 처리 내역을 선택하세요 ]</td>
	</tr>
</table>
<% end if %>



<%
set ocsaslist   = Nothing
set orefund     = Nothing
set oordermaster = Nothing
set OCsDetail = Nothing
set OCsDelivery = Nothing
set OReturnAddr = Nothing
set OCsConfirm = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
