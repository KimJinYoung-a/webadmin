<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  현장수령주문
' History : 2012.05.10 서동석 생성
'			2012.05.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
dim orderserial , aplot , userid , memodispyn
	orderserial = requestCheckVar(request("orderserial"),11)
	aplot = requestCheckVar(request("aplot"),1)

memodispyn = FALSE

dim ojumun
set ojumun = new CJumunMaster
	ojumun.FRectOrderSerial = orderserial

	if (ojumun.FRectOrderSerial<>"") then
	    ojumun.SearchJumunList
	end if

if orderserial <> "" then
	if ojumun.FTotalCount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('해당되는 주문건이 없습니다');"
		response.write "	self.close()"
		response.write "</script>"
		dbget.close()	:	response.End
	else
		userid = ojumun.FMasterItemList(0).fuserid
		memodispyn = TRUE
	end if
end if

Dim isValidOrder
	isValidOrder = (ojumun.FResultCount=1)

dim ix ,sellsum
sellsum = 0

dim IsCanNotView, ValidUpcheItem : ValidUpcheItem = False

if (isValidOrder) then
	ojumun.SearchJumunDetail orderserial
end if

'for ix=0 to ojumun.FJumunDetail.FDetailCount-1
'    if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then
'	    if UCase(ojumun.FJumunDetail.FJumunDetailList(ix).FMakerid)=UCase(session("ssBctID")) and (ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y") and (ojumun.FJumunDetail.FJumunDetailList(ix).Fcancelyn<>"Y") then
'            if (ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate<3) or IsNULL(ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate) then
'                IsCanNotView = true
'            end if
'            ValidUpcheItem = True
'        end if
'    end if
'next

''사은품정보 추가 : 상품 발주 시 생성됨.
dim oGift
set oGift = new COrderGift

if (isValidOrder)  then
    oGift.FRectOrderSerial = orderserial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist
end if

dim isFinishValid     : isFinishValid= false
dim isAlreadyFinished : isAlreadyFinished= false

if (isValidOrder) then
	isFinishValid = (ojumun.FMasterItemList(0).FIpkumdiv>3) and (ojumun.FMasterItemList(0).FIpkumdiv<8)
	isFinishValid = isFinishValid and (ojumun.FMasterItemList(0).FCancelyn="N")
end if

Dim iNoticsStr : iNoticsStr=""

if (isValidOrder) and (Not isFinishValid) then
	if (ojumun.FMasterItemList(0).FIpkumdiv>7) then
		iNoticsStr = "이미 처리 완료된 주문 입니다."
		isAlreadyFinished = true
	elseif (ojumun.FMasterItemList(0).FIpkumdiv<4) then
		iNoticsStr = "결제이전 주문건 입니다."
	elseif (ojumun.FMasterItemList(0).FCancelyn<>"N") then
		iNoticsStr = "취소된 주문 입니다."
	elseif (ojumun.FMasterItemList(0).Fjumundiv<>"7") then
		iNoticsStr = "현장수령 주문건이 아닙니다."
	end if
end if

if (isValidOrder) and (isFinishValid) then
	if date() <> ojumun.FMasterItemList(0).freqdate then
		iNoticsStr = iNoticsStr & "\n수령(예정)일이 "&ojumun.FMasterItemList(0).Freqdate&" 일인 주문입니다"
	end if
end if
%>

<script language='javascript'>

var finishvalid=<%= LCASE(isFinishValid) %>

//완료처리
function reSearchAct(comp){
	//alert('사용기간이 아닙니다');
	//return;

	var frm = comp.form;
	if ((frm.orderserial.value=='<%= orderserial %>')&&(finishvalid)){
		if (!confirm('완료처리 하시겠습니까?')){
			return;
		}

		frm.action='/admin/apps/siteReceive/siteReceive_process.asp';
		frm.method='post';
		frm.mode.value='siteReceivefinsh';
		frm.target='view';
		frm.submit();

	}else{
		frm.action="";
		frm.method="get";
		frm.submit();
	}
}

//영수증 출력
function plotReceipt(){
<% if (isValidOrder) then	%>
	var iplot = document.iSrp350plot;

	iplot.prtPrinterName("SRP-350");

	if (iplot.isPrinterExists()!=true){
		alert('SRP-350 설치 후 사용요망');
		return;
	}

	var y = 0;
	var ygap = 28;
	iplot.beginPrint();
	iplot.prtFontName("돋움");
	iplot.prtFontSize(8);
	iplot.prtTextAlign(0);
	iplot.prtFontStyle(-1);
	iplot.prtFontStyle(0);

	iplot.prtDrawImage(200,0,'logo.gif');

	x = 20;
	y = 130;
	iplot.prtTextOut(x,y,"-------  [ 현장 수령 확인증 - 고객용 ]  -------");
	y += ygap +10;
	iplot.prtTextOut(x,y,"주문번호 : <%= orderserial %>" );
	y += ygap;
	iplot.prtTextOut(x,y,"수령일시 : <%= now() %>");
	y += ygap;
	iplot.prtTextOut(x,y,"수령인 : <%= ojumun.FMasterItemList(0).FReqName %>");
	y += ygap;
	y += ygap;
	iplot.prtTextOut(x,y,"품목");
	//iplot.prtTextOut(x+240,y,"단가");
	iplot.prtTextOut(x+400,y,"수량");
	//iplot.prtTextOut(x+440,y,"금액");
	y += ygap;
	iplot.prtTextOut(x,y,"----------------------------------------------");
	ygap = 26;
	y += ygap;
	iplot.prtFontSize(7);

	<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>

	<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fcancelyn ="N" then %>
		iplot.prtTextOut(x,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemname %>');

		<% if ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName <> "" then %>
			iplot.prtTextOut(x+180,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>');
		<% end if %>

		<% if ojumun.FJumunDetail.FJumunDetailList(ix).FitemNo<>1 then %>
			iplot.prtFontSize(10);
	    <% end if %>
		iplot.prtTextOut(x+404,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).FitemNo %>');
		iplot.prtFontSize(7);
		y += ygap;
	<% end if %>
	<% end if %>
	<% next %>
	iplot.prtTextOut(x,y,"----------------------------------------------");
	ygap = 28;
	y += ygap;
	iplot.prtTextOut(x,y+70,"고객 서명 :");
	iplot.prtTextOut(x,y+100,"----------------------------------------------");
	iplot.prtTextOut(x,y+70,".");
	iplot.endPrint();

	iplot.beginPrint();
	iplot.prtFontName("돋움");
	iplot.prtFontSize(8);
	iplot.prtTextAlign(0);
	iplot.prtFontStyle(-1);
	iplot.prtFontStyle(0);
	x = 20;
	y = 0;
	iplot.prtTextOut(x,y,"-------  [ 현장 수령 확인증 - 매장용 ]  -------");
	y += ygap +10;
	iplot.prtTextOut(x,y,"주문번호 : <%= orderserial %>" );
	y += ygap;
	iplot.prtTextOut(x,y,"수령일시 : <%= now() %>");
	y += ygap;
	iplot.prtTextOut(x,y,"수령인 : <%= ojumun.FMasterItemList(0).FReqName %>");
	y += ygap;
	y += ygap;
	iplot.prtTextOut(x,y,"품목");
	//iplot.prtTextOut(x+240,y,"단가");
	iplot.prtTextOut(x+400,y,"수량");
	//iplot.prtTextOut(x+440,y,"금액");
	y += ygap;
	iplot.prtTextOut(x,y,"----------------------------------------------");
	ygap = 26;
	y += ygap;
	iplot.prtFontSize(7);
	<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fcancelyn ="N" then %>
		iplot.prtTextOut(x,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemname %>');

		<% if ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName <> "" then %>
			iplot.prtTextOut(x+180,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>');
		<% end if %>

		<% if ojumun.FJumunDetail.FJumunDetailList(ix).FitemNo<>1 then %>
			iplot.prtFontSize(10);
	    <% end if %>
		iplot.prtTextOut(x+404,y,'<%= ojumun.FJumunDetail.FJumunDetailList(ix).FitemNo %>');
		iplot.prtFontSize(7);
		y += ygap;
	<% end if %>
	<% end if %>
	<% next %>
	iplot.prtTextOut(x,y,"----------------------------------------------");
	ygap = 28;
	y += ygap;
	iplot.prtTextOut(x,y+70,"고객 서명 :");
	iplot.prtTextOut(x,y+100,"----------------------------------------------");
	iplot.prtTextOut(x,y+ygap+100,".");
	iplot.endPrint();

<% end if %>
}

//온로드이벤트
function GetOnLoad(){
	<% if (iNoticsStr<>"") and aplot <> "Y" then %>
		alert('<%= iNoticsStr %>');
	<% end if %>

	document.frmAct.orderserial.focus();
	document.frmAct.orderserial.select();

	<% if aplot = "Y" then %>
		plotReceipt()
	<% end if %>
}

window.onload = GetOnLoad;

//cs메모 입력&수정
function csmemoreg(orderserial,id,userid){
	var csmemoreg = window.open('/common/cscenter/cscenter_memo.asp?orderserial='+orderserial+'&userid='+userid+'&id='+id,'csmemoreg','width=800,height=350,scrollbars=yes,resizable=yes');
	csmemoreg.focus();
}

</script>

<iframe id="view" name="view" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="200" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>현장 수령 주문 상세</b></font>
				</td>
				<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">

				</td>

			</tr>
		</table>
	</td>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" method="post" action="siteReceive_process.asp">
<input type="hidden" name="mode">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
    	<font style="font-size:16px;font-weight:bold">주문번호</font> :
    	<input type="text" name="orderserial" value="<%= orderserial %>" size="14" maxlength="16" onKeyPress="if (event.keyCode == 13) { reSearchAct(this);return false;}" style="font-size:18px;font-weight:bold">

		<% if (isValidOrder) and ( isFinishValid) then %>
			<input type="button" value="완료처리" onclick="reSearchAct(this)" class="button">
	    <% end if %>
	</td>

</tr>
</form>
<% if (isValidOrder) then	%>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td width="225" bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FOrderSerial %></td>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">아이디</td>
	<td width="225" bgcolor="#FFFFFF"><%= printUserId(ojumun.FMasterItemList(0).fuserid,2,"**") %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">결제방식</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).JumunMethodName %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
	<td bgcolor="#FFFFFF"><font color="<%= ojumun.FMasterItemList(0).IpkumDivColor %>"><%= ojumun.FMasterItemList(0).IpkumDivName %></font></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">주문일</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FRegDate %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">입금일</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FIpkumDate %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">취소여부</td>
	<td bgcolor="#FFFFFF">
		<font color="<%= ojumun.FMasterItemList(0).CancelYnColor %>"><%= ojumun.FMasterItemList(0).CancelYnName %></font>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>"><b>수령(예정)일</b></td>
	<td bgcolor="#FFFFFF">
		<%= ojumun.FMasterItemList(0).Freqdate %>
	</td>
</tr>
<tr><td colspan="4"  bgcolor="#777777"></td></tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">구매자</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyName %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">구매자전화</td>
	<td bgcolor="#FFFFFF"><%= printUserId(ojumun.FMasterItemList(0).FBuyPhone,3,"*") %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">구매자핸드폰</td>
	<td bgcolor="#FFFFFF"><%= printUserId(ojumun.FMasterItemList(0).FBuyHp,3,"*") %></td>
</tr>
<tr><td colspan="4"  bgcolor="#777777"></td></tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqName %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인전화</td>
	<td bgcolor="#FFFFFF"><%= printUserId(ojumun.FMasterItemList(0).FReqPhone,3,"*") %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인핸드폰</td>
	<td bgcolor="#FFFFFF"><%= printUserId(ojumun.FMasterItemList(0).FReqHp,3,"*") %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
	<td colspan="3" bgcolor="#FFFFFF">
	<%= nl2br(ojumun.FMasterItemList(0).FComment) %>
	</td>
</tr>
<% If Not IsNULL(ojumun.FMasterItemList(0).Fbeadaldate) then %>
	<tr><td colspan="4"  bgcolor="#777777"></td></tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><b>출고완료일</b></td>
		<td bgcolor="#FFFFFF" colspan=3>
			<%= ojumun.FMasterItemList(0).fbeadaldate %>
		</td>
	</tr>
<% end if %>

<% If Not IsNULL(ojumun.FMasterItemList(0).Fcardribbon) then %>
	<% If ojumun.FMasterItemList(0).Fcardribbon <> 3 then %>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">카드 리본 구분</td>
			<td colspan="3" bgcolor="#FFFFFF">
			<% If ojumun.FMasterItemList(0).Fcardribbon = 1 then %>카드<% elseIf ojumun.FMasterItemList(0).Fcardribbon = 2 then %>리본<% else %> 없음<% End if %>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">카드리본메세지</td>
			<td colspan="3" bgcolor="#FFFFFF">
			<% if ojumun.FMasterItemList(0).Ffromname<>"" then %>
			From.<%= nl2br(ojumun.FMasterItemList(0).Ffromname) %><br>
			<% End if %>
			<%= nl2br(ojumun.FMasterItemList(0).Fmessage) %>
			</td>
		</tr>
	<% End if %>
<% End if %>
<% else %>
<tr bgcolor="#FFFFFF" height="30" >
	<td colspan="4" align="center">
		<% if (orderserial<>"") then %>
			검색 결과가 없습니다.
		<% else %>
			주문 번호를 입력 하세요.
		<% end if %>
	</td>
</tr>
<% End if %>
</table>

<br>
<% if (isValidOrder) then	%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
    	<b>주문상품정보</b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">상품코드</td>
	<td width="50">이미지</td>
	<td>상품명<font color="blue">[옵션명]</font></td>
	<td width="35">수량</td>
	<td width="50">판매가격</td>
	<td width="35">배송<br>구분</td>
	<td width="35">취소<br>여부</td>
</tr>
<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
<% sellsum = sellsum + ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost*ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %></td>
	<td><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>" target="_blank"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></a></td>
	<td align="left">
		<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %>
		<br>
		<% if ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName <> "" then %>
			<font color="blue">[<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>]</font>
		<% end if %>
	</td>
	<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %></td>
	<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %></td>
	<td>
		<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y" then %>
		<font color="red">업체</font>
		<% else %>
		10x10
		<% end if %>
	</td>
	<td>
		<font color="<%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateColor %>"><%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateStr %></font>
	</td>
</tr>
<% if ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail <> "" then %>
<tr bgcolor="#FFFFFF">
	<td colspan="7"><%= nl2BR(ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail) %></td>
</tr>
<% end if %>
<% end if %>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td>합계</td>
	<td colspan="4" align="right"><%= FormatNumber(sellsum,0) %></td>
	<td colspan="2"></td>
</tr>
</table>

<p>
<% if oGift.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td width="50" align="center" >사은품</td>
    <td>
    <% for ix=0 to oGift.FResultCount -1 %>
        [<%= oGift.FItemList(ix).Fevt_name %>] <%= oGift.FItemList(ix).GetEventConditionStr %><br>
    <% next %>
    </td>
</tr>
</table>
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="50" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">

	</td>
	<td colspan="15" align="right">
    	<input type="button" value="재 출력" onClick="plotReceipt();" <%= CHKIIF(isAlreadyFinished,"","disabled") %> class="button">
    	<!--<input type="button" value="test" onClick="plotReceipt();" class="button">-->
	</td>
</tr>
</table>
<% end if %>
<!-- 표 하단바 끝-->

<% if memodispyn then %>
<Br>

<!-- 관련메모 시작-->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
	        <td>
	        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>관련메모</b>
	        </td>
	        <td align="right">
				<input type="button" class="button" value="메모신규입력" onclick="csmemoreg('<%= orderserial %>','','<%= userid %>');">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<iframe id="i_history_memo" name="i_history_memo" onload="this.style.height=this.contentWindow.document.body.scrollHeight;" src="/common/cscenter/iframeHistory.asp?userid=<%= userid %>&orderserial=<%= orderserial %>&id=" width="100%" scrolling="auto" frameborder="0"></iframe>
	</td>
</tr>
</table>
<% end if %>

<%
set ojumun = Nothing
set oGift = Nothing
%>

<script language="javascript">drawSrp350PlotOcx('iSrp350plot','1,0,0,1');</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->