<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<%
dim idx
idx = requestCheckVar(request("idx"), 32)

'// ===========================================================================
dim ioneas,i
set ioneas = new CCSASList

ioneas.FRectMakerID = session("ssBctID")
ioneas.FRectCsAsID = idx
ioneas.GetOneCSASMaster

'==============================================================================
''환불정보
dim orefund

set orefund = New CCSASList

orefund.FRectCsAsID = requestCheckVar(request("idx"), 32)

orefund.GetOneRefundInfo



'// ===========================================================================
if (ioneas.FResultCount<1) then
    response.write "<script>alert('유효한 접수번호가 아닙니다.');</script>"
    response.write dbget.close()	:	response.End
end if

dim ioneasDetail
set ioneasDetail= new CCSASList
ioneasDetail.FRectCsAsID = idx
ioneasDetail.GetCsDetailList


'// ===========================================================================
dim sqlStr

if IsNull(ioneas.FOneItem.Fconfirmdate) then
	sqlStr = " update [db_cs].[dbo].tbl_new_as_list set confirmdate = getdate() where id = " + CStr(idx) + " "
	dbget.Execute sqlStr
end if


'// ===========================================================================
dim IsChangeReturn
dim ioneRefasDetail

dim ioneRefas, IsRefASExist, refasid
dim chulgoyn
dim receiveyn
dim receiveonly

dim divcd, refdivcd

chulgoyn = "N"
receiveyn = ""
IsRefASExist = False
receiveonly = requestCheckVar(request("receiveonly"), 32)
set ioneRefas = new CCSASList
IsChangeReturn = False
refasid = 0

divcd = ioneas.FOneItem.FDivCD
if ((ioneas.FOneItem.FDivCD = "A000") or (ioneas.FOneItem.FDivCD = "A100")) then
	'// 맞교환출고, 상품변경 맞교환출고

	if (ioneas.FOneItem.Fcurrstate >= "B006") then
		chulgoyn = "Y"
	end if

	ioneRefas.FRectMakerID = session("ssBctID")
	ioneRefas.FRectCsRefAsID = idx
	ioneRefas.GetOneCSASMaster

	refdivcd = ioneRefas.FOneItem.FDivCD

	if (ioneRefas.FResultCount>0) then
	    IsRefASExist = True
	    refasid = ioneRefas.FOneItem.FID

	    if (ioneRefas.FOneItem.Fcurrstate >= "B006") then
	    	receiveyn = "Y"
	    else
	    	receiveyn = "N"
	    end if
	end if

	IsChangeReturn = (ioneas.FOneItem.FDivCD = "A100")				'// 상품변경 맞교환출고

	set ioneRefasDetail = new CCSASList

	if (IsChangeReturn) then
		ioneRefasDetail.FRectMakerID = session("ssBctID")
		ioneRefasDetail.FRectCsRefAsID = idx
		ioneRefasDetail.GetCsDetailList
	end if
end if


'// ===========================================================================
dim currrowspan
dim tmpStr

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function ViewOrderDetail(frm){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "upcheorderpop", props);
    frm.target = 'upcheorderpop';
    frm.action="/designer/common/viewordermaster.asp"
	frm.submit();

}

function SaveReceiveFin(frm) {
	var ret = confirm('저장 하시겠습니까?');


	if (ret){
		frm.submit();
	}
}

function trim(value) {
 return value.replace(/^\s+|\s+$/g,"");
}

function SaveFin(frm){
	//alert('잠시 준비중입니다.');
	//return;
	var val;

	frm.finishmemo.value = trim(frm.finishmemo.value);
	if (frm.finishmemo.value.length<1){
		alert('처리 내용을 입력해 주세요.');
		frm.finishmemo.focus();
		return;
	}

	<% if (ioneas.FOneItem.FDivCD = "A100") then %>

		if (frm.customerrealbeasongpay) {
			if (frm.customerrealbeasongpay.value == "") {
				frm.customerrealbeasongpay.value = "0";
			}

			if (frm.customerrealbeasongpay.value*0 != 0) {
				alert("고객 추가 배송비는 숫자만 가능합니다.");
				frm.customerrealbeasongpay.focus();
				return;
			}

			if (frm.customerrealbeasongpay.value != "0") {
				frm.customerreceiveyn.value = "Y";
			} else {
				frm.customerreceiveyn.value = "N";
			}
		}

	<% end if %>

	if ($("#needChkYN_N").val()) {
		if ($("#needChkYN_N").prop("checked") === false && $("#needChkYN_Y").prop("checked") === false) {
			alert("즉시완료 여부를 선택하세요.");
			$("#needChkYN_N").focus();
			return;
		}

		if ($("#needChkYN_N").prop("checked") === true) {
			<% if (ioneas.FOneItem.FDivCD = "A000") then %>
			if ($("#needRefChkYN_N").prop("checked") === false && $("#needRefChkYN_Y").prop("checked") === false) {
				alert("교환회수내역의 즉시완료 여부를 선택하세요.");
				$("#needRefChkYN_N").focus();
				return;
			}
			<% End If %>
			if (confirm("고객센터의 확인잘차 없이 즉시 완료처리됩니다.\n\n계속 진행하시겠습니까?") === false) {
				return;
			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');


	if (ret){
		frm.submit();
	}
}

function SetReceiveYes(frm) {
	// not used
	var ret = confirm('회수완료 처리 하시겠습니까?');

	if (ret) {
		frm.receiveyn[0].checked = true;

		frm.submit();
	}
}



function GetRadioValue(obj) {
	for (var i=0; i < obj.length; i++) {
		if (obj[i].checked) {
			return obj[i].value;
		}
	}

	return "";
}

</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="upchecs_process.asp">
	<input type="hidden" name="orderserial" value="<%= ioneas.FOneItem.FOrderSerial %>">
	<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
	<input type="hidden" name="id" value="<%= ioneas.FOneItem.FID %>">
	<input type="hidden" name="refasid" value="<%= refasid %>">
	<input type="hidden" name="receiveonly" value="<%= receiveonly %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>배송CS 처리답변</b>
			&nbsp;&nbsp;
			작성자 :
	        	<% if(Lcase(ioneas.FOneItem.Fwriteuser)=Lcase(ioneas.FOneItem.FUserID)) then %>
	        	<b>고객 직접 작성</b>
	        	<% else %>
	        	텐바이텐 고객센터
	        	<b><% end if %></b>
        	&nbsp;&nbsp;
        	작성일 : <b><%= CStr(ioneas.FOneItem.Fregdate) %></b>
        	&nbsp;&nbsp;
        	<% if not IsNULL(ioneas.FOneItem.Ffinishdate) then %>
        	완료일 : <b><%= CStr(ioneas.FOneItem.Ffinishdate) %></b>
        	<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>" height="25">주문번호</td>
		<td>
			<%= ioneas.FOneItem.Forderserial %>
			<input type="button" class="button" value="주문상세보기" onclick="ViewOrderDetail(frm);">
		</td>
		<td width="45%" rowspan="7" valign="top">
			<% if (ioneas.FOneItem.Fdivcd="A000") or (ioneas.FOneItem.Fdivcd="A012") or (ioneas.FOneItem.Fdivcd="A100") or (ioneas.FOneItem.Fdivcd="A112") then %> <!-- 맞교환 설명 -->
				<b>* 맞교환 도움말</b>
			<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- 누락재발송 설명 -->
				<b>* 누락재발송 도움말</b>
			<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- 반품 설명 -->
				<b>* 반품관련 도움말</b>
				<br>반품접수가 될경우, 고객님께 발송하신 택배사 전화번호를 안내해드리며,
				<br>상품을 받으신 택배사를 통해 <font color="blue">착불반송</font>을 해주시도록 안내를 해드리고 있습니다.
				<br><font color="blue">변심 반품의 경우, 브랜드 상품 일부반품의 경우 반품배송비, 전체 반품의 경우 착불반송포함 왕복배송비를 차감한 금액을 고객님께 환불해드리며,
				<br>차감된 금액은 업체정산내역에 자동으로 등록됩니다.</font>
				<br><font color="red">(편도 2,500원 / 왕복 5,000원 차감)</font>
				<br>
				<br>반송상품이 도착하면, 접수내용과 확인하신 후,
				<br>아래쪽 처리내용에 내용을 남겨주시면, 고객센터에 내용이 전달되며,
				<br>고객센터에서 반품취소처리 및 고객환불을 진행합니다.
				<br>
				<br>*처리프로세스
				<br>1.접수
				<br>2.업체완료처리 --> 고객센터에 처리결과 전달
				<br>3.고객센터완료처리 --> 고객에게 처리결과 안내 및 메일발송
			<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- 출고시 유의사항 설명 -->
				<b>* 출고시 유의사항 도움말</b>
				<br>주문건 확인 후, 고객님이 주문관련 변경을 요청하셨을 경우,
				<br>출고시 유의사항으로 등록됩니다.
				<br>ex)배송지변경/상품변경/상품옵션변경
				<br>
				<br><font color="red">텐바이텐 고객센터에서 별도로 가능여부 확인을 위해 연락드립니다.</font>
			<% else %>

			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">고객명</td>
		<td><%= ioneas.FOneItem.FCustomerName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">고객아이디</td>
		<td><%= ioneas.FOneItem.FUserID %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">제목</td>
		<td><%= ioneas.FOneItem.FTitle %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">접수사유</td>
		<td><%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="50">
		<td bgcolor="<%= adminColor("tabletop") %>">접수내용</td>
		<td>
			<%
			tmpStr = replace(ioneas.FOneItem.Fcontents_jupsu,"<","&lt;")
			tmpStr = replace(tmpStr,">","&gt;")
			tmpStr = replace(tmpStr,VbCrlf,"<br>")
			%>
			<%= tmpStr %>
		</td>
	</tr>
	<% if (ioneasDetail.FResultCount>0) then %>
	<tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>">접수상품</td>
	    <td>
	        <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a">
	        <tr bgcolor="<%= adminColor("topbar") %>" align="center" height="25">
	            <td width="50"></td>
	            <td width="50">이미지</td>
	            <td width="50">상품코드</td>
	            <td>상품명<font color="blue">[옵션명]</font></td>
	            <td width="50">판매가</td>
	            <td width="40">수량</td>
	        </tr>
		<% if (IsChangeReturn) then %>
	    	<% for i=0 to ioneRefasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
	   	            <td>원주문</td>
		            <td><img src="<%= ioneRefasDetail.FItemList(i).FSmallImage %>" width="50"></td>
		            <td><%= ioneRefasDetail.FItemList(i).FItemID %></td>
		            <td align="left">
		            	<%= ioneRefasDetail.FItemList(i).Fitemname %>
		            	<% if ioneRefasDetail.FItemList(i).Fitemoptionname<>"" then %>
		            	<br>
		            	<font color="blue">[<%= ioneRefasDetail.FItemList(i).Fitemoptionname %>]</font>
		            	<% end if %>
		            </td>
		            <td align="right"><%= FormatNumber(ioneRefasDetail.FItemList(i).Fitemcost,0) %></td>
		            <td align="center"><%= ioneRefasDetail.FItemList(i).Fitemno %></td>
		        </tr>
	        <% next %>
	        <% for i=0 to ioneasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
    	            <td>└─&gt;</td>
		            <td><img src="<%= ioneasDetail.FItemList(i).FSmallImage %>" width="50"></td>
		            <td><%= ioneasDetail.FItemList(i).FItemID %></td>
		            <td align="left">
		            	<%= ioneasDetail.FItemList(i).Fitemname %>
		            	<% if ioneasDetail.FItemList(i).Fitemoptionname<>"" then %>
		            	<br>
		            	<font color="blue">[<%= ioneasDetail.FItemList(i).Fitemoptionname %>]</font>
		            	<% end if %>
		            </td>
		            <td align="right"><%= FormatNumber(ioneasDetail.FItemList(i).Fitemcost,0) %></td>
		            <td align="center"><%= ioneasDetail.FItemList(i).Fitemno %></td>
		        </tr>
	        <% next %>
        <% else %>
	        <% for i=0 to ioneasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
    	            <td>접수상품</td>
		            <td><img src="<%= ioneasDetail.FItemList(i).FSmallImage %>" width="50"></td>
		            <td><%= ioneasDetail.FItemList(i).FItemID %></td>
		            <td align="left">
		            	<%= ioneasDetail.FItemList(i).Fitemname %>
		            	<% if ioneasDetail.FItemList(i).Fitemoptionname<>"" then %>
		            	<br>
		            	<font color="blue">[<%= ioneasDetail.FItemList(i).Fitemoptionname %>]</font>
		            	<% end if %>
		            </td>
		            <td align="right"><%= FormatNumber(ioneasDetail.FItemList(i).Fitemcost,0) %></td>
		            <td align="center"><%= ioneasDetail.FItemList(i).Fitemno %></td>
		        </tr>
	        <% next %>
        <% end if %>
	        </table>
	    </td>
	</tr>
	<% end if %>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>배송CS 처리결과작성</b>
			&nbsp;&nbsp;
			*처리 내용 입력시 <font color=red>송장번호</font>등 상세내역을 기재해 주세요
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	<% if (receiveonly = "Y") then %><!-- (맞교환출고, 상품변경 맞교환출고) + (맞교환회수 등록된 경우) + (맞교환회수 입력) -->
		<td width="130" height="120" bgcolor="<%= adminColor("tabletop") %>">출고 처리내용</td>
		<td>
			<%= nl2br(ioneas.FOneItem.Fcontents_finish) %>
		</td>
	<% else %>
		<td width="130" bgcolor="<%= adminColor("tabletop") %>">처리내용</td>
		<td>
			<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
		</td>
	<% end if %>

		<td width="45%" rowspan="20" valign="top">
			<% if (ioneas.FOneItem.Fdivcd="A000") or (ioneas.FOneItem.Fdivcd="A100") then %> <!-- 맞교환 설명 -->
				<% if (receiveonly = "Y") then %>
					*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
					<br>(고객님께 오픈되는 정보가 아닙니다.)
					<br>
					<br><font color="red">고객 추가배송비가 있는경우 수령액을 꼭 입력 부탁드립니다.</font>
					<br>
					<br><font color="blue">*처리내용 입력요청사항</font>
					<br>기타내용 :
					<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
				<% else %>
					*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
					<br>(고객님께 오픈되는 정보가 아닙니다.)
					<br>
					<br><font color="red">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</font>
					<br>
					<br><font color="blue">*처리내용 입력요청사항</font>
					<br>출고일 :
					<br>기타내용 :
					<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
				<% end if %>
			<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- 누락재발송 설명 -->
				*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
				<br>(고객님께 오픈되는 정보가 아닙니다.)
				<br>
				<br><font color="red">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</font>
				<br>
				<br><font color="blue">*처리내용 입력요청사항</font>
				<br>출고일 :
				<br>기타내용 :
				<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
			<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- 반품 설명 -->
				*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
				<br>(고객님께 오픈되는 정보가 아닙니다.)
				<br>
				<br><font color="red">반품상품 입고 완료 후, 처리내용 입력과 함께 완료처리 부탁드립니다.</font>
				<br>
				<br><font color="blue">*처리내용 입력요청사항</font>
				<br>반품방법 : 고객선불 / 착불
				<br>반품사유 : 불량반품 / 고객반품
				<br>환불계좌 : 은행명 + 계좌번호 + 예금주명(고객님이 첨부한 경우)
				<br>기타내용 :
				<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
			<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- 출고시 유의사항 설명 -->
				*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
				<br>(고객님께 오픈되는 정보가 아닙니다.)
				<br>
				<br><font color="red">고객센터에서 요청한 출고유의사항에 대한 처리유무를 알려주시기 바랍니다.</font>
				<br>발송 후, 이 내용을 확인하셨을 경우에도, 미반영 출고로 완료처리 부탁드립니다.
			<% else %>

			<% end if %>

		</td>
	</tr>


	<%
	'[코드정리]
	'------------------------------------------------------------------------------
	'A008			주문취소
	'
	'A004			반품접수(업체배송)
	'A010			회수신청(텐바이텐배송)
	'
	'A001			누락재발송
	'A002			서비스발송
	'
	'A200			기타회수
	'
	'A000			맞교환출고
	'A100			상품변경 맞교환출고
	'
	'A009			기타사항
	'A006			출고시유의사항
	'A700			업체기타정산
	'
	'A003			환불
	'A005			외부몰환불요청
	'A007			카드,이체,휴대폰취소요청
	'
	'A011			맞교환회수(텐바이텐배송)
	'A012			맞교환반품(업체배송)

	'A111			상품변경 맞교환회수(텐바이텐배송)
	'A112			상품변경 맞교환반품(업체배송)
	%>

	<% if (receiveonly <> "Y") then %>
	<% if InStr(",A000,A100,A001,A002,A009,A006,A012,A004,", divcd) > 0 then %>
	<% if (divcd = "A004") then %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">반품사유</td>
		<td>
			<%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">반품배송비</td>
		<td>
			<% if (orefund.FOneItem.Frefunddeliverypay<>0) then %>
			환불시 반품배송비 <%= FormatNumber(orefund.FOneItem.Frefunddeliverypay*-1, 0) %> 원 차감후 환불
			<% else %>
			없음
			<% end if %>
		</td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">즉시완료 여부</td>
		<td>
			<% if (ioneas.FOneItem.FneedChkYN="F") then %>
			<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="F", "checked", "") %> > 고객센터 확인용
			<% else %>
			<input type="radio" id="needChkYN_N" name="needChkYN" value="N" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="N", "checked", "") %> > 즉시완료(고객센터 확인 불필요)
			<input type="radio" id="needChkYN_Y" name="needChkYN" value="Y" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="Y", "checked", "") %> > 고객센터 확인요청
			<% end if %>
		</td>
	</tr>
	<% If (divcd = "A000") Then %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">교환회수</td>
		<td>
			<input type="radio" id="needRefChkYN_N" name="needRefChkYN" value="N" <%= CHKIIF(receiveyn = "Y", "checked", "") %> > 즉시완료(상품회수완료 되었음)
			<input type="radio" id="needRefChkYN_Y" name="needRefChkYN" value="Y" > 상품회수 이전
		</td>
	</tr>
	<% end if %>
	<% end if %>
	<% end if %>

	<% if (receiveonly = "Y") then %>

		<!-- ============================ (맞교환출고, 상품변경 맞교환출고) + (맞교환회수 등록된 경우) + (맞교환회수 입력) -->
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" height="30">출고 운송장</td>
			<td>
				<%= DeliverDivCd2Nm(ioneas.FOneItem.FSongjangdiv) %>
				<%= ioneas.FOneItem.Fsongjangno %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">출고 상태</td>
			<td>
				<% if (chulgoyn = "Y") then %>
					출고완료
				<% elseif (chulgoyn = "N") then %>
					<font color="blue">출고이전</font>
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">회수 처리내용</td>
			<td>
				<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneRefas.FOneItem.Fcontents_finish %></textarea>
			</td>
		</tr>
		<% if InStr(",A000,A100,A001,A002,A009,A006,A012,A004,", refdivcd) > 0 then %>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">즉시완료 여부</td>
			<td>
				<% if (ioneRefas.FOneItem.FneedChkYN="F") then %>
				<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="F", "checked", "") %> > 고객센터 확인용
				<% else %>
				<input type="radio" id="needChkYN_N" name="needChkYN" value="N" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="N", "checked", "") %> > 즉시완료(고객센터 확인 불필요)
				<input type="radio" id="needChkYN_Y" name="needChkYN" value="Y" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="Y", "checked", "") %> > 고객센터 확인요청
				<% end if %>
			</td>
		</tr>
		<% end if %>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">택배접수</td>
			<td>
            	<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
				<%
				Select Case ioneRefas.FOneItem.FsongjangRegGubun
					Case "U"
						Response.Write("업체 접수")
					Case "C"
						Response.Write("고객직접접수")
					Case "T"
						Response.Write("상담사 접수")
					Case Else
						Response.Write ioneRefas.FOneItem.FsongjangRegGubun
				End Select
				%>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">운송장입력</td>
			<td>
            	<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
				<%= ioneRefas.FOneItem.FsongjangRegUserID %>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">관련운송장</td>
			<td>
				<% drawSelectBoxDeliverCompany "songjangdiv",ioneRefas.FOneItem.FSongjangdiv %>
				<input type="text" class="text" name="songjangno" value="<%= ioneRefas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
			</td>
		</tr>
		<% if (ioneas.FOneItem.Fdivcd="A100") then %>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="<%= adminColor("tabletop") %>">고객 추가배송비(예정)</td>
				<td>
					<input type="text" class="text_ro" name="customeraddbeasongpay" value="<%= ioneas.FOneItem.Fcustomeraddbeasongpay %>" size="10" ReadOnly >
					&nbsp;
		    	    <select class="select" name="customeraddmethod" class="text" disabled>
			    	    <option value="">선택
			    	    <option value="1" <% if (ioneas.FOneItem.Fcustomeraddmethod = "1") then %>selected<% end if %>>박스동봉
			    	    <option value="2" <% if (ioneas.FOneItem.Fcustomeraddmethod = "2") then %>selected<% end if %>>택배비 고객부담
			    	    <option value="9" <% if (ioneas.FOneItem.Fcustomeraddmethod = "9") then %>selected<% end if %>>환불액에서 차감
			    	    <option value="5" <% if (ioneas.FOneItem.Fcustomeraddmethod = "5") then %>selected<% end if %>>기타
		    	    </select>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="<%= adminColor("tabletop") %>">고객 추가배송비(확인)</td>
				<input type="hidden" name="customerreceiveyn" value="<%= ioneas.FOneItem.Fcustomerreceiveyn %>">
				<td>
					<input type="text" class="text" name="customerrealbeasongpay" value="<%= ioneas.FOneItem.Fcustomerrealbeasongpay %>" size="10"> * 박스동봉 인 경우 업체에서 확인한 금액
				</td>
			</tr>
		<% end if %>
		<!-- ============================ (맞교환출고, 상품변경 맞교환출고) + (맞교환회수 등록된 경우) + (맞교환회수 입력) -->

	<% else %>

		<!-- ============================ 이외 -->
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">택배접수</td>
			<td>
            	<% if ioneas.FOneItem.IsRequireSongjangNO then %>
				<%
				Select Case ioneas.FOneItem.FsongjangRegGubun
					Case "U"
						Response.Write("업체 접수")
					Case "C"
						Response.Write("고객직접접수")
					Case "T"
						Response.Write("상담사 접수")
					Case Else
						Response.Write ioneas.FOneItem.FsongjangRegGubun
				End Select
				%>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">운송장입력</td>
			<td>
            	<% if ioneas.FOneItem.IsRequireSongjangNO then %>
				<%= ioneas.FOneItem.FsongjangRegUserID %>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" height="25">관련운송장</td>
			<td>
				<% drawSelectBoxDeliverCompany "songjangdiv",ioneas.FOneItem.FSongjangdiv %>
				<input type="text" class="text" name="songjangno" value="<%= ioneas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
			</td>
		</tr>
		<% if (IsRefASExist) then %>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="130" height="120" bgcolor="<%= adminColor("tabletop") %>">회수 처리내용</td>
				<td>
					<%= nl2br(ioneRefas.FOneItem.Fcontents_finish) %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">택배접수</td>
				<td>
					<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
					<%
					Select Case ioneRefas.FOneItem.FsongjangRegGubun
						Case "U"
							Response.Write("업체 접수")
						Case "C"
							Response.Write("고객직접접수")
						Case "T"
							Response.Write("상담사 접수")
						Case Else
							Response.Write ioneRefas.FOneItem.FsongjangRegGubun
					End Select
					%>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">운송장입력</td>
				<td>
					<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
					<%= ioneRefas.FOneItem.FsongjangRegUserID %>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">회수 운송장</td>
				<td>
					<%= DeliverDivCd2Nm(ioneRefas.FOneItem.FSongjangdiv) %>
					&nbsp;
					<%= ioneRefas.FOneItem.Fsongjangno %>
				</td>
			</tr>
			<% if (ioneas.FOneItem.Fdivcd="A100") then %>
				<tr bgcolor="#FFFFFF">
					<td  height="30" bgcolor="<%= adminColor("tabletop") %>">고객 추가배송비(예정)</td>
					<td>
						<% if Not IsNull(ioneas.FOneItem.Fcustomeraddbeasongpay) then %>
							<%= FormatNumber(ioneas.FOneItem.Fcustomeraddbeasongpay, 0) %> 원
							(
							<% if (ioneas.FOneItem.Fcustomeraddmethod = "1") then %>
								박스동봉
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "2") then %>
								택배비 고객부담
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "9") then %>
								환불액에서 차감
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "5") then %>
								기타
							<% end if %>
							)
						<% end if %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td  height="30" bgcolor="<%= adminColor("tabletop") %>">고객 추가배송비(확인)</td>
					<td>
						<% if Not IsNull(ioneas.FOneItem.Fcustomerrealbeasongpay) then %>
							<%= FormatNumber(ioneas.FOneItem.Fcustomerrealbeasongpay, 0) %> 원
						<% end if %>
						&nbsp;
						 * 박스동봉 인 경우 업체에서 확인한 금액
					</td>
				</tr>
			<% end if %>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">회수 상태</td>
				<td>
					<% if (receiveyn = "Y") then %>
						회수완료
					<% elseif (receiveyn = "N") then %>
						<font color="blue">회수이전</font>
					<% end if %>
				</td>
			</tr>
		<% end if %>
		<!-- ============================ 이외 -->

	<% end if %>

	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>업체 추가 정산</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">회수배송비</td>
		<td>
			<%= FormatNumber(orefund.FOneItem.Frefunddeliverypay*-1, 0) %> 원
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">추가정산배송비</td>
		<td>
			<%= FormatNumber(ioneas.FOneItem.Fadd_upchejungsandeliverypay, 0) %> 원
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">추가정산사유</td>
		<td>
			<%= ioneas.FOneItem.Fadd_upchejungsancause %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">총정산배송비</td>
		<td>
			<b><%= FormatNumber((ioneas.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1), 0) %> 원</b>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="35" bgcolor="FFFFFF">
		<td colspan="15" align="center">

			<% if (IsRefASExist) and (receiveonly = "Y") then %>

				<% if ioneRefas.FOneItem.Fcurrstate="B007" then %>

				<% else %>
					<input type="button" class="button" value="회수완료처리" onclick="javascript:SaveFin(frm);">
				<% end if %>

			<% else %>

				<% if ioneas.FOneItem.Fcurrstate="B007" then %>

				<% else %>
				    <% if ((ioneas.FOneItem.Fdivcd = "A000") or (ioneas.FOneItem.Fdivcd = "A100")) and (IsRefASExist) then %>
					<input type="button" class="button" value="출고처리" onclick="javascript:SaveFin(frm);">
					<% else %>
					<input type="button" class="button" value="완료처리" onclick="javascript:SaveFin(frm);">
					<% end if %>
				<% end if %>

			<% end if %>

			<input type="button" class="button" value="목록보기" onClick="location.href='/designer/jumunmaster/upchecslist.asp';">
		</td>
	</tr>
</table>

<!-- 표 하단바 끝-->

<%
set ioneas = Nothing
set ioneasDetail = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
