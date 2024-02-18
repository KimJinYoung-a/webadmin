<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/oldmisendcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/upchebeasongcls.asp"-->
<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<%
'' XXXX 브랜드/ 어드민 공통사용
'' 핑거스
dim C_ADMIN_USER : C_ADMIN_USER = True

dim idx : idx= requestCheckVar(request("idx"),10)

dim omisend
set omisend = new COldMiSend
omisend.FRectDetailIDx = idx
omisend.getOneOldMisendItem

if (omisend.FResultCount<1) then
    response.write "검색결과가 없습니다."
    dbget.close() : response.end
end if

''업체인경우
if (Not C_ADMIN_USER) then
    if (LCase(omisend.FOneItem.FMakerid) <> LCASE(session("ssBctID"))) then
        response.write "권한이 없습니다."
        dbget.close() : response.end
    end if
end if

dim PreDispMail
PreDispMail = (omisend.FOneItem.isMisendAlreadyInputed) and (omisend.FOneItem.FMisendReason<>"05")


dim MisendReasonStr : MisendReasonStr = "03,02,08,09,04,10,07"
dim MisendReasonArr : MisendReasonArr = Split(MisendReasonStr, ",")
dim i, tmpStr

%>
<script language='javascript'>

function SetStockOut() {
    var itemSoldOutFlag = document.all.itemSoldOutFlag;
	var itemSoldOutContent = document.all.itemSoldOutContent;
	var itemSoldOutButton = document.all.itemSoldOutButton;

	if (itemSoldOutFlag.style.display == "none") {
        itemSoldOutFlag.style.display = "inline";
		itemSoldOutContent.style.display = "inline";
		if (itemSoldOutButton) {
			itemSoldOutButton.disabled = false;
		}
	} else {
        itemSoldOutFlag.style.display = "none";
		itemSoldOutContent.style.display = "none";
		if (itemSoldOutButton) {
			itemSoldOutButton.disabled = true;
		}
	}
}

function ShowHideObject(comp) {
    var frm = comp.form;
	var doc = document.all;
	var tmpObj;

	// 입고예정일
    var divipgodate = doc.divipgodate;

	// 품절출고불가
	var itemSoldOutFlag = doc.itemSoldOutFlag;
	var itemSoldOutContent = doc.itemSoldOutContent;

	// SMS/MAIL
	var SMSContentAll = doc.SMSContentAll;
	var MailContentAll = doc.MailContentAll;

	SMSContentAll.style.display = "none";
	MailContentAll.style.display = "none";

	<% for i = 0 to UBound(MisendReasonArr) %>
	doc.SMSContent<%= MisendReasonArr(i) %>.style.display = "none";
	doc.MailContent<%= MisendReasonArr(i) %>.style.display = "none";
	<% next %>

	if (divipgodate) {
		if (comp.value == "05") {
			divipgodate.style.display = "none";
		} else {
			divipgodate.style.display = "inline";
		}
	}

	if (comp.value == "05") {
		itemSoldOutFlag.style.display = "inline";
		itemSoldOutContent.style.display = "inline";

		SMSContentAll.style.display = "none";
		MailContentAll.style.display = "none";
	} else {
		itemSoldOutFlag.style.display = "none";
		itemSoldOutContent.style.display = "none";

		if (comp.value != "") {
			tmpObj = eval("doc.SMSContent" + comp.value);
			tmpObj.style.display = "inline";

			tmpObj = eval("doc.MailContent" + comp.value);
			tmpObj.style.display = "inline";

			SMSContentAll.style.display = "inline";
			MailContentAll.style.display = "inline";
		}
	}

	<% if (C_ADMIN_USER) then %>
		if (comp.value == "05") {
			frm.ckSendSMS.disabled = true;
			frm.ckSendEmail.disabled = true;
			frm.ckSendSMS.checked = false;
			frm.ckSendEmail.checked = false;
		} else {
			frm.ckSendSMS.disabled = false;
			frm.ckSendEmail.disabled = false;
			frm.ckSendSMS.checked = true;
			frm.ckSendEmail.checked = true;
		}
	<% end if %>
}

function MisendInput(){
    var frm = document.frmMisend;
    var today= new Date();
    //today = new Date(today.getYear(),today.getMonth(),today.getDate());  //오늘도 가능하도록
    today = new Date(<%=year(now())%>,<%=month(now())-1%>,<%=Day(now())%>);  //2016/09/08 수정.

    var inputdate;

    if (frm.MisendReason.value.length<1){
        alert('미출고 사유를 입력하세요.');
        frm.MisendReason.focus();
        return;
    }


    // 품절출고불가(05)
    if (frm.MisendReason.value != "05") {
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('출고 예정일을 입력하세요.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('출고 예정일은 오늘 이후날짜로 설정이 가능합니다.');
            ipgodate.focus();
            return;
        }

		/*
        if (frm.ckSendSMS && frm.ckSendEmail) {
        	if ((frm.ckSendSMS.checked != true) && (frm.ckSendEmail.checked != true)) {
				alert("SMS 와 메일발송 둘중 하나는 체크해야 합니다.");
				return;
        	}
        }
        */
    } else {
		// 품절등록시 변경가능 옵션 설정
		<% if (omisend.FOneItem.FItemoption <> "0000") then %>
		if (frm.reqaddstr) {
			var regExp = /변경가능 옵션 :[ \t]*\r?\n/;

			if(regExp.test(frm.reqaddstr.value) == true) {
				frm.reqaddstr.value = frm.reqaddstr.value.replace("변경가능 옵션 :", "변경가능 옵션 : 없음");

				alert('변경가능 옵션을 입력하세요.\n\n==>>> 미입력시 \"없음\" 으로 입력됩니다. <<<==');
				frm.reqaddstr.focus();
				return;
			}
		}
		<% end if %>
	}

	if (frm.MisendReason.value != "05") {
		frm.reqaddstr.value = "";
	}

    if (confirm('미출고 사유를 저장 하시겠습니까?')){
	    frm.action = "upchebeasong_Process.asp";
	    frm.submit();
	}
}

function MisendInputUpche() {
	var frm = document.frmMisend;

	if (confirm('품절등록 하시겠습니까?')){
	    frm.action = "upchebeasong_Process.asp";
	    frm.submit();
	}
}

function SetupObject() {
	<% if (C_ADMIN_USER) then %>
		ShowHideObject(frmMisend.MisendReason);
	<% end if %>
	popupResize(680);
}
window.onload = SetupObject;

</script>

<% if omisend.FResultCount>0 then %>
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMisend" method="post" action="upchebeasong_Process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="misendInputOne">
	<input type="hidden" name="detailidx" value="<%= omisend.FOneItem.Fidx %>">
	<input type="hidden" name="Sitemid" value="<%= omisend.FOneItem.FItemID %>">
	<input type="hidden" name="Sitemoption" value="<%= omisend.FOneItem.FItemOption %>">
	<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>미출고사유 입력</b>
	    </td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
    	<td width="130">상품코드</td>
    	<td width="480"><%= omisend.FOneItem.FItemID %>
    	    <% if (omisend.FOneItem.FCancelyn<>"N") then %>
				<b><font color="#CC3333">[취소주문]</font></b>
				<script language="javascript">alert("취소된 거래 입니다.");</script>
			<% else %>
			    <% if (omisend.FOneItem.FDetailCancelYn="Y") then %>
				    <b><font color="#CC3333">[취소상품]</font></b>
			    <% else %>
				    [정상주문]
			    <% end if%>
			<% end if %>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF">
	    <td>이미지</td>
	    <td><img src="<%= omisend.FOneItem.Fsmallimage %>" width="50" height="50"></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>상품명</td>
	    <td><%= omisend.FOneItem.FItemName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>옵션</td>
	    <td><%= omisend.FOneItem.FItemoptionName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>주문수량</td>
	    <td>
			<%= omisend.FOneItem.FItemcnt %>개
			<% if (C_ADMIN_USER) then %>
				(부족수량 <%= omisend.FOneItem.Fitemlackno %>)
			<% end if %>
	    </td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF">
	    <td>미출고사유</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>
				<%= omisend.FOneItem.getMiSendCodeName %>
				<% if omisend.FOneItem.isMisendAlreadyInputed and (omisend.FOneItem.FMisendReason <> "05") then %>
					<input type="button" class="button" value="품절전환" onClick="SetStockOut();">
					<input type="hidden" name="MisendReason" value="05">
				<% end if %>
	        <% else %>
				<select name="MisendReason" id="MisendReason" class="select" onChange="ShowHideObject(this);">
					<option value="">---------</option>
					<option value="03" <%= ChkIIF(omisend.FOneItem.FMisendReason="03","selected"," ") %> >출고지연</option>
					<option value="02" <%= ChkIIF(omisend.FOneItem.FMisendReason="02","selected"," ") %> >주문제작</option>
					<option value="08" <%= ChkIIF(omisend.FOneItem.FMisendReason="08","selected"," ") %> >수입</option>
					<option value="09" <%= ChkIIF(omisend.FOneItem.FMisendReason="09","selected"," ") %> >가구배송</option>
					<option value="04" <%= ChkIIF(omisend.FOneItem.FMisendReason="04","selected"," ") %> >예약배송</option>
					<option value="10" <%= ChkIIF(omisend.FOneItem.FMisendReason="10","selected"," ") %> >업체휴가</option>
					<option value="07" <%= ChkIIF(omisend.FOneItem.FMisendReason="07","selected"," ") %> >고객지정배송</option>
					<option value="">---------</option>
					<option value="05" <%= ChkIIF(omisend.FOneItem.FMisendReason="05","selected"," ") %> >품절출고불가</option>
					<option value="">---------</option>
				</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >상품 품절처리
			<input type="radio" name="itemSoldOut" value="S">상품 일시품절처리
			</span>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>품절출고불가상세<br>(고객센터전달사항)</td>
	    <td>
			<span id="itemSoldOutContent" name="itemSoldOutContent" style="display=<% if (omisend.FOneItem.FupcheRequestString = "") then %>none<% else %>inline<% end if %>" align="right" >
			<textarea class="textarea" name="reqaddstr" cols="65" rows="9" <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>readonly<% end if %> ><% if (omisend.FOneItem.FupcheRequestString = "") then %>부족수량 : <%= omisend.FOneItem.FItemcnt %>개
고객통화여부 : N
변경가능 옵션 :
기타 전달사항 :
<% else %><%= omisend.FOneItem.FupcheRequestString %><% end if %></textarea>
			</span>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>출고예정일</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>
	        	<%= omisend.FOneItem.FMisendIpgodate %>
	        <% else %>
				<div id="divipgodate" name="divipgodate">
					<input class="text" type="text" name="ipgodate" value="<%= omisend.FOneItem.FMisendIpgodate %>" size="10" maxlength="10">
					<a href="javascript:calendarOpen(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
				</div>
			<% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>고객안내여부</td>
	    <td>
	        <% if (C_ADMIN_USER) then %>
				<input name="ckSendSMS" type="checkbox" checked  >SMS발송<% if (omisend.FOneItem.FisSendSms="Y") then %>(Y)<% end if %>
				&nbsp;
				<input name="ckSendEmail" type="checkbox" checked  >MAIL발송<% if (omisend.FOneItem.FisSendEmail="Y") then %>(Y)<% end if %>
	        <% else %>
    	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendSms="Y","SMS발송완료 &nbsp; ","") %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendEmail="Y","MAIL발송완료 &nbsp; ","") %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendCall="Y","통화안내완료","") %>
    	        	<!-- 고객안내가 완료된 건은 미출고사유 및 출고예정일 수정 불가 -->
    	        <% else %>
        	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS발송
        	        &nbsp;
        	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL발송
    	        <% end if %>
    	    <% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td colspan="2">
	    	<font color="blue">
	    	미출고 사유가 출고지연 및 주문제작(수입)일 경우, 아래의 내용으로 고객님께 SMS와 메일이 발송됩니다.<br>
	    	고객님께 안내된 출고예정일을 꼭 지켜주시기 바라며, 변동사항이 생길경우, 고객센터로 연락 부탁드립니다.<br>
	    	</font>
	    	<font color="red">
	       	품절출고불가인 경우, 고객님께 SMS 및 메일이 발송되지 않으며, 텐바이텐고객센터에서<br>
	    	별도로 고객님께 연락을 드릴 예정입니다.
	    	</font>
	    </td>
	</tr>
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2" align="center">
	    <% if (C_ADMIN_USER) then %>
	        <input type="button" class="button" value="미출고 사유 저장" onclick="MisendInput();">
	    <% else %>
    	    <% if omisend.FOneItem.isMisendAlreadyInputed then %>
				<% if omisend.FOneItem.isMisendAlreadyInputed and (omisend.FOneItem.FMisendReason <> "05") then %>
					<input type="button" class="button" id="itemSoldOutButton" name="itemSoldOutButton" value="품절등록" onClick="MisendInputUpche();" disabled><br><br>
					(이외의 사유변경은 고객센터로 문의하세요.)
				<% else %>
					수정 불가
				<% end if %>
    	    <% else %>
    	    <input type="button" class="button" value="미출고 사유 저장" onclick="MisendInput();">
    	    <% end if %>
    	<% end if %>
	    </td>
	</tr>
	</form>
</table>

<p>

<!-- 출고지연/주문제작 선택시 아래 보이는 내용입니다. 사유선택시 실시간으로 보이도록 -->

<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td>
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>SMS 발송내용</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" id="SMSContentAll" style="display:none" >
	    <td>
        	<table width="100%" align="center" cellspacing="1" cellpadding="0" class="a" >
			<%
			for i = 0 to UBound(MisendReasonArr)
				tmpStr = GetMichulgoSMSString(MisendReasonArr(i))
				tmpStr = Replace(tmpStr, "[출고예정일]", "<span id='MaySendDate" + MisendReasonArr(i) + "' name='MaySendDate" + MisendReasonArr(i) + "'>" + CStr(CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD")) + "</span>")

				tmpStr = Replace(tmpStr, "[상품명]", DdotFormat(omisend.FOneItem.FItemName,16))
				tmpStr = Replace(tmpStr, "[상품코드]", omisend.FOneItem.FItemID)
			%>
			<tr bgcolor="#FFFFFF" id="SMSContent<%= MisendReasonArr(i) %>">
            	<td>
					<%= tmpStr %>
            	</td>
            </tr>
			<% next %>
            </table>
        </td>
    </tr>
</table>

<p>

<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td>
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>MAIL 발송내용</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" id="MailContentAll" style="display:none">
    	<td>
    		<!-- 메일 내용 시작 -->
    		<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>

						<!-- 컨텐츠 시작 -->
						<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="a">
						<tr>
							<td>
								<a href="http://www.thefingers.co.kr" target="_blank">
									<img src="http://image.thefingers.co.kr/2016/mail/img_logo.png" width="600" height="85" border="0" />
								</a>
							</td>
						</tr>
						<tr>
							<td style="border:7px solid #eeeeee;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
								<tr>
									<td><img src="http://fiximage.10x10.co.kr/web2008/mail/b01_img.gif" width="586"> </td>
								</tr>
								<tr>
									<td height="30" style="padding:0 15px 0 15px">
										<!-- 고객명 / 주문번호 -->
										<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
										<tr>
											<td class="black12px">

											</td>
											<td align="right" class="gray11px02">주문번호 : <span class="sale11px01"><%= omisend.FOneItem.FOrderserial %></span></td>
										</tr>
										<tr>
											<td height="3" colspan="2" class="black12px" style="padding:5px;" bgcolor="#99CCCC"></td>
										</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td style="padding:5px 15px 20px 15px">
										<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
										<%
										for i = 0 to UBound(MisendReasonArr)
											tmpStr = GetMichulgoMailString(MisendReasonArr(i))
										%>
										<tr bgcolor="#FFFFFF" id="MailContent<%= MisendReasonArr(i) %>">
											<td>
												<%= Replace(tmpStr, "\n", "<br>") %>
											</td>
										</tr>
										<% next %>

										<tr>
											<td>

												<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
												<tr>
													<td colspan="2" class="sky12pxb" style="padding: 10 0 5 0">*상품정보</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td width="150" height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" align="center" style="padding-top:2px;">상품</td>
													<td width="450"class="gray12px02" style="padding-left:10px;padding-top:2px;"><img src="<%= omisend.FOneItem.Fsmallimage %>" width="50" height="50"></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" align="center" style="padding-top:2px;">상품코드</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemID %> </td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">상품명</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemName %></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">옵션명</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemoptionName %></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">주문수량</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemcnt %>개</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td colspan="2" class="sky12pxb" style="padding: 20 0 5 0">*발송예정안내</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">발송(판매)자</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><%= omisend.FOneItem.getDlvCompanyName %></b></td>
													<!-- 텐바이텐 배송일 경우 텐바이텐 물류센터, 업체일경우, 업체회사명-->
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">발송예정일</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><span id="iMisendIpgodate2" name="iMisendIpgodate2"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span></b></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr id="iEMAILMENTNOTI1" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="07","none","inline") %>">
													<td colspan="2" class="gray12px02" style="padding: 5 0 5 0">
													* 발송예정일로부터 1~2일 후에 상품을 받아보실 수 있습니다.<br>
													</td>
												</tr>

												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td><img src="http://image.thefingers.co.kr/academy2009/mail/mail_bottom.gif" width="600" height="30" /></td>
							</tr>
							<tr>
								<td height="51" style="border-bottom:1px solid #eaeaea;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td style="padding-left:20px;"><img src="http://image.thefingers.co.kr/academy2009/mail/bottom_text.gif" width="245" height="26" /></td>
										<td width="128"><a href="http://www.thefingers.co.kr/cscenter/csmain.asp" onFocus="blur()" target="_blank"><img src="http://image.thefingers.co.kr/academy2009/mail/btn_cscenter.gif" width="108" height="31" border="0" /></a></td>
									</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td style="padding:10px 0 15px 0;line-height:17px;" class="gray11px02" class="a">
									(03086) 서울시 종로구 대학로12길 31 자유빌딩 5층 (주)텐바이텐 theFingers<br>
									대표이사 : 최은희  &nbsp;사업자등록번호 : 211-87-00620  &nbsp;통신판매업 신고번호 : 제 01-1968호  &nbsp;개인정보 보호 및 청소년 보호책임자 : 이문재<br>
									<span class="black11px">고객행복센터:TEL 1644-1557  &nbsp;E-mail:<a href="mailto:customer@thefingers.co.kr" class="link_black11pxb">customer@thefingers.co.kr</a> </span>
								</td>
							</tr>
							</table>
						<!-- 컨텐츠 끝 -->
					</td>
				</tr>
			</table>

    		<!-- 메일 내용 끝 -->
    	</td>
    </tr>
</table>


<% else %>
<table width="600">
<tr>
    <td align="center">취소된 상품이거나 해당 주문 내역이 없습니다.</td>
</tr>
</table>
<% end if %>

<%
set omisend = Nothing
%>
<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
