<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/misendcls.asp"-->

<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<%
''브랜드/ 어드민 공통사용

dim idx : idx= requestCheckVar(request("idx"),10)

dim omisend
set omisend = new COldMiSend
omisend.FRectDetailIDx = idx
omisend.getOneOldMisendItem

if (omisend.FResultCount<1) then
    response.write "검색결과가 없습니다."
    dbget.close() : response.end
end if

if (LCase(omisend.FOneItem.FMakerid) <> LCASE(session("ssBctID"))) then
    response.write "권한이 없습니다."
    dbget.close() : response.end
end if

dim PreDispMail
PreDispMail = (omisend.FOneItem.isMisendAlreadyInputed) and (omisend.FOneItem.FMisendReason<>"05")

%>
<script language='javascript'>
function getOnload(){
    // popupResize(640);
}
window.onload = getOnload;

function ShowDateBox(comp){
    var frm = comp.form;
    var iid = comp.id;
    var idiv = document.all.divipgodate;
    var isms = document.all.iSMSDISP;
    var iemail = document.all.iEMAILDISP;
    var isDPlusOver = true;
    var isold = document.all.itemSoldOutFlag

    document.all.iSMSDISP02.style.display = "none";
    document.all.iSMSDISP03.style.display = "none";
    document.all.iSMSDISP04.style.display = "none";
    document.all.iSMSDISP02_1.style.display = "none";
    document.all.iSMSDISP03_1.style.display = "none";
    document.all.iSMSDISP04_1.style.display = "none";

    document.all.iEMAILMENT02.style.display = "none";
    document.all.iEMAILMENT03.style.display = "none";
    document.all.iEMAILMENT04.style.display = "none";
    document.all.iEMAILMENT02_1.style.display = "none";
    document.all.iEMAILMENT03_1.style.display = "none";
    document.all.iEMAILMENT04_1.style.display = "none";

    if ((comp.value=="03")||(comp.value=="02")||(comp.value=="04")){
        idiv.style.display = "inline";
        isms.style.display = "inline";
        iemail.style.display = "inline";

        if ((frm.baljudate.value.length>0)&&(frm.ipgodate.value.length>0)){
            if (getDiffDay(frm.baljudate.value,frm.ipgodate.value)<2){
                isDPlusOver=false;
            }
        }

        if (comp.value=="03"){
            if (isDPlusOver){
                document.all.iSMSDISP03.style.display = "inline";
                document.all.iSMSDISP03_1.style.display = "none";

                document.all.iEMAILMENT03.style.display = "inline";
                document.all.iEMAILMENT03_1.style.display = "none";
            }else{
                document.all.iSMSDISP03.style.display = "none";
                document.all.iSMSDISP03_1.style.display = "inline";

                document.all.iEMAILMENT03.style.display = "none";
                document.all.iEMAILMENT03_1.style.display = "inline";
            }
        }else if(comp.value=="02"){
            if (isDPlusOver){
                document.all.iSMSDISP02.style.display = "inline";
                document.all.iSMSDISP02_1.style.display = "none";

                document.all.iEMAILMENT02.style.display = "inline";
                document.all.iEMAILMENT02_1.style.display = "none";
            }else{
                document.all.iSMSDISP02.style.display = "none";
                document.all.iSMSDISP02_1.style.display = "inline";

                document.all.iEMAILMENT02.style.display = "none";
                document.all.iEMAILMENT02_1.style.display = "inline";
            }
        }else if(comp.value=="04"){
            if (isDPlusOver){
                document.all.iSMSDISP04.style.display = "inline";
                document.all.iSMSDISP04_1.style.display = "none";

                document.all.iEMAILMENT04.style.display = "inline";
                document.all.iEMAILMENT04_1.style.display = "none";
            }else{
                document.all.iSMSDISP04.style.display = "none";
                document.all.iSMSDISP04_1.style.display = "inline";

                document.all.iEMAILMENT04.style.display = "none";
                document.all.iEMAILMENT04_1.style.display = "inline";
            }
        }
    }else{
        idiv.style.display = "none";
        isms.style.display = "none";
        iemail.style.display = "none";
    };

   //품절출고불가
   if (comp.value=="05") {
        isold.style.display = "inline";
   }else{
        isold.style.display = "none";
   }
}

function ipgodateChange(comp){
    var v = comp.value;
    if (v.length<10) v = "YYYY-MM-DD";

    document.getElementById("iMisendIpgodate02").innerHTML = v;
    document.getElementById("iMisendIpgodate02_1").innerHTML = v;
    document.getElementById("iMisendIpgodate03").innerHTML = v;
    document.getElementById("iMisendIpgodate03_1").innerHTML = v;
    document.getElementById("iMisendIpgodate04").innerHTML = v;
    document.getElementById("iMisendIpgodate04_1").innerHTML = v;

    document.getElementById("iMisendIpgodate2").innerHTML = v;

    ShowDateBox(frmMisend.MisendReason);
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


    //출고지연(03), 주문제작(02), 예약배송(04)
    if ((frm.MisendReason.value=="03")||(frm.MisendReason.value=="02")||(frm.MisendReason.value=="04")){
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('출고 예정일을 입력하세요.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('출고 예정일은 오늘 이후날짜로 설정이 가능합니다.');
            //ipgodate.focus();
            return;
        }
    }

    if (confirm('미출고 사유를 저장 하시겠습니까?')){
	    frm.action = "upchebeasong_Process.asp";
	    frm.submit();
	}
}

function getDiffDay(d1,d2){   // 두 날짜의 차이구함

  var v1=d1.split("-");
  var v2=d2.split("-");

  var a1=new Date(v1[0],v1[1],v1[2]);
  var a2=new Date(v2[0],v2[1],v2[2]);
  return parseInt((a2-a1)/(1000*3600*24));  //1000*3600*24 는 날의차이 만약 월의 차이를 구하고 싶다면 *30곱하면 월 12를 곱하면 년

}

</script>

<% if omisend.FResultCount>0 then %>
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMisend" method="post" action="upchebeasong_Process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="misendInputOne">
	<input type="hidden" name="detailidx" value="<%= omisend.FOneItem.Fidx %>">
	<input type="hidden" name="baljudate" value="<%= CHKIIF(IsNULL(omisend.FOneItem.Fbaljudate),"",Left(omisend.FOneItem.Fbaljudate,10)) %>">
	<input type="hidden" name="upcheconfirmdate" value="<%= CHKIIF(IsNULL(omisend.FOneItem.Fupcheconfirmdate),"",Left(omisend.FOneItem.Fupcheconfirmdate,10)) %>">

	<input type="hidden" name="Sitemid" value="<%= omisend.FOneItem.FItemID %>">
	<input type="hidden" name="Sitemoption" value="<%= omisend.FOneItem.FItemOption %>">

	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>미출고사유 입력</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td width="130">상품코드</td>
    	<td width="480"><%= omisend.FOneItem.FItemID %>

    	    <% if (omisend.FOneItem.FCancelyn<>"N") then %>
			<b><font color="#CC3333">[취소주문]</font></b>
			<script language='javascript'>alert('취소된 거래 입니다.');</script>
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
	    <td><img src="<%= omisend.FOneItem.Fsmallimage %>" width="60" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>상품명</td>
	    <td><%= omisend.FOneItem.FItemName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>옵션</td>
	    <td><%= omisend.FOneItem.FItemoptionName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>주문수량</td>
	    <td><%= omisend.FOneItem.FItemcnt %>개</td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td>미출고사유</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	        <%= omisend.FOneItem.getMiSendCodeName %>
	        <% else %>
	        <select name="MisendReason" id="MisendReason" class="select" onChange="ShowDateBox(this);">
				<option value="">---------</option>
				<option value="03" <%= ChkIIF(omisend.FOneItem.FMisendReason="03","selected"," ") %> >출고지연</option>
				<option value="05" <%= ChkIIF(omisend.FOneItem.FMisendReason="05","selected"," ") %> >품절출고불가</option>
				<option value="02" <%= ChkIIF(omisend.FOneItem.FMisendReason="02","selected"," ") %> >주문제작</option>
				<option value="04" <%= ChkIIF(omisend.FOneItem.FMisendReason="04","selected"," ") %> >예약배송</option>
				<!-- 텐바이텐배송 미출고사유와 통합했으면 합니다. -->
			</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >상품 품절처리
			<input type="radio" name="itemSoldOut" value="S">상품 일시품절처리
			</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>출고예정일</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	        <%= omisend.FOneItem.FMisendIpgodate %>
	        <% else %>
	        <div id="divipgodate" name="divipgodate" <%= ChkIIF(omisend.FOneItem.FMisendReason="03" or omisend.FOneItem.FMisendReason="02","style='display:inline'","style='display:none'") %> >
			    <input class="text" type="text" name="ipgodate" value="<%= omisend.FOneItem.FMisendIpgodate %>" size="10" maxlength="10" onChange="ipgodateChange(this);">
			    <a href="javascript:calendarOpen(frmMisend.ipgodate);ipgodateChange(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
			</div>
			<% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>고객안내여부</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	            <%= CHKIIF(omisend.FOneItem.FisSendSms="Y","SMS발송완료/","") %>
	            <%= CHKIIF(omisend.FOneItem.FisSendEmail="Y","MAIL발송완료/","") %>
	            <%= CHKIIF(omisend.FOneItem.FisSendCall="Y","통화안내완료","") %>
	        <!-- 고객안내가 완료된 건은 미출고사유 및 출고예정일 수정 불가 -->
	        <% else %>
    	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS발송
    	        &nbsp;
    	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL발송
	        <% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td colspan="2">
	    	<font color="blue">
	    	미출고 사유가 출고지연 및 주문제작일 경우, 아래의 내용으로 고객님께 SMS와 메일이 발송됩니다.<br>
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
		    <% if omisend.FOneItem.isMisendAlreadyInputed then %>
		    수정 불가
		    <% else %>
		    <input type="button" class="button" value="미출고 사유 저장" onclick="MisendInput();">
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
	<tr bgcolor="#FFFFFF" id="iSMSDISP" style="display:<%= chkIIF(PreDispMail,"inline","none") %>" >
	    <td>
        	<table width="610" align="center" cellspacing="1" cellpadding="0" class="a" >
        	<tr bgcolor="#FFFFFF" id="iSMSDISP02" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="02","inline","none") %>">
            	<td>
            		[더핑거스 출고지연안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품은 주문제작 상품으로 <span id="iMisendIpgodate02" name="iMisendIpgodate02"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP02_1" style="display:none">
            	<td>
            		[더핑거스 출고지연안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품이 <span id="iMisendIpgodate02_1" name="iMisendIpgodate02_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 감사합니다.
            	</td>
            </tr>
        	<tr bgcolor="#FFFFFF" id="iSMSDISP03" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="03","inline","none") %>">
            	<td>
            		[더핑거스 출고지연안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품이 <span id="iMisendIpgodate03" name="iMisendIpgodate03"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 쇼핑에 불편을 드려 죄송합니다.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP03_1" style="display:none">
            	<td>
            		[더핑거스 출고지연안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품이 <span id="iMisendIpgodate03_1" name="iMisendIpgodate03_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 감사합니다.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP04" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="04","inline","none") %>">
            	<td>
            		[더핑거스 출고예정안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품은 예약배송상품으로 <span id="iMisendIpgodate04" name="iMisendIpgodate4"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 감사합니다.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP04_1" style="display:none">
            	<td>
            		[더핑거스 출고예정안내]주문하신 상품중 <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)상품은 예약배송상품으로 <span id="iMisendIpgodate04_1" name="iMisendIpgodate04_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>에 발송될 예정입니다. 감사합니다.
            	</td>
            </tr>
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
	<tr bgcolor="#FFFFFF" id="iEMAILDISP" style="display:<%= chkIIF(PreDispMail,"inline","none") %>">
    	<td>
    		<!-- 메일 내용 시작 -->

    		<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>

						<!-- 컨텐츠 시작 -->
						<table width="600" border="0" align="center" cellspacing="0" cellpadding="0" class="a">
						<tr>
							<td><a href="http://www.thefingers.co.kr" target="_blank" onFocus="blur()">
								<img src="http://image.thefingers.co.kr/2016/mail/img_logo.png" width="600" height="85" border="0" /></a>
							</td>
						</tr>
						<tr>
							<td style="border:7px solid #eeeeee;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
								<tr>
									<td><img src="http://image.thefingers.co.kr/academy2010/email/sorry.gif" width="586"> </td>
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
										<tr id="iEMAILMENT03" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="03","inline","none") %>">
											<td>
												<!-- 출고지연일 경우 D+2 -->
												안녕하세요.   고객님<br>
												고객님께서 주문하신 상품이 발송이 지연될 예정입니다.<br>
												아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>
												고객행복센터로 연락 부탁드립니다.<br>
												쇼핑에 불편을 드린 점 진심으로 사과드리며, 기분좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT03_1" style="display:none">
											<td>
												<!-- 출고지연일 경우 D+0/1 -->
												안녕하세요.   고객님<br>
												고객님께서 주문하신 상품의 출고안내 메일입니다.<br>
												아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>
												고객행복센터로 연락 부탁드립니다.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT02" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="02","inline","none") %>">
										    <td>
												<!-- 주문제작 경우 D+2 -->
												안녕하세요.  고객님<br>
												고객님께서 주문하신 상품은 주문 후 제작되는 상품으로<br>
												일반상품과 달리 주문제작기간이 소요되는 상품입니다.<br>
												아래와 같이 발송 예정일을 안내해드리오니,<br>
												판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT02_1" style="display:none">
										    <td>
												<!-- 주문제작 경우 D+0/1 -->
												안녕하세요.  고객님<br>
												고객님께서 주문하신 상품의 출고안내 메일입니다.<br>
												아래와 같이 발송예정일을 안내해 드립니다.<br>
												판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT04" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="04","inline","none") %>">
										    <td>
												<!-- 예약상품 경우 D+2 -->
												안녕하세요.  고객님<br>
												고객님께서 주문하신 상품의 출고안내 메일입니다.<br>
                                                주문하신 상품은 예약배송상품으로 아래 발송예정일에 발송될 예정이며,<br>
                                                부득이한 사정으로 상품취소를 원하시는 경우,<br>
                                                고객행복센터로 연락 부탁드립니다.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT04_1" style="display:none">
										    <td>
												<!-- 예약상품 경우 D+0/1 -->
												안녕하세요.  고객님<br>
												고객님께서 주문하신 상품의 출고안내 메일입니다.<br>
                                                주문하신 상품은 예약배송상품으로 아래 발송예정일에 발송될 예정이며,<br>
                                                부득이한 사정으로 상품취소를 원하시는 경우,<br>
                                                고객행복센터로 연락 부탁드립니다.<br>

											</td>
										</tr>
										<tr id="iEMAILMENT05" style="display:none">
										    <td>
										        <!-- 품절 출고불가일 경우 --- 이건 업체에서는 발송 안함 더핑거스 고객센터에서만 발송 멘트 나중에 추가-->
										    </td>
										</tr>
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
													<td width="450"class="gray12px02" style="padding-left:10px;padding-top:2px;"><img src="<%= omisend.FOneItem.Fsmallimage %>" width="60" ></td>
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
													<!-- 더핑거스 배송일 경우 더핑거스 물류센터, 업체일경우, 업체회사명-->
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">발송예정일</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><span id="iMisendIpgodate2" onClick="ipgodateChange(frmMisend.ipgodate);" name="iMisendIpgodate2"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span></b></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
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
								(03086) 서울시 종로구 대학로12길 31 자유빌딩 5층 (주)텐바이텐 더핑거스<br>
								대표이사 : 최은희  &nbsp;사업자등록번호:211-87-00620  &nbsp;통신판매업 신고번호 : 제 01-1968호  &nbsp;개인정보 보호 및 청소년 보호책임자 : 이문재<br>
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->