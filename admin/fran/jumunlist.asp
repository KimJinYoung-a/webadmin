<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 주문서관리
' History : 2009.04.07 서동석 생성
'			2010.06.03 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid ,designer, statecd, baljucode ,notipgo, minusjumun, shopdiv, tot_jumunsuplycash, tot_totalsuplycash, tot_jumunbuycash, tot_totalbuycash
dim yyyy1,mm1 ,dd1,yyyy2,mm2,dd2, totaljumunsellcash ,i ,fromDate ,toDate , datefg
dim itemgubun, itemid, itemoption
dim tplgubun, totalyn, popupyn
dim totaljumunforeign_sellcash, totaljumunforeign_suplycash, totalforeign_suplycash
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	designer = request("designer")
	statecd  = request("statecd")
	baljucode= request("baljucode")
	notipgo = request("notipgo")
	minusjumun = request("minusjumun")
	shopdiv = request("shopdiv")
	shopid = request("shopid")
	page = request("page")
	if page="" then page=1
	datefg = request("datefg")
	tplgubun = request("tplgubun")
	popupyn = request("popupyn")
	itemgubun = request("itemgubun")
	itemid = request("itemid")
	itemoption = request("itemoption")


totalyn="Y"

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))-1
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim osheet
set osheet = new COrderSheet
	osheet.FRectFromDate = fromDate
	osheet.FRectToDate = toDate
	osheet.frectdatefg = datefg
	osheet.FCurrPage = page
	osheet.Fpagesize=20

	if baljucode<>"" then
		osheet.FRectBaljuCode = baljucode
	else
		osheet.FRectBaljuid = shopid
		osheet.FRectStatecd = statecd
		osheet.FRectMakerid = designer
		osheet.FRectDivCodeArr = "('501','502','503')"
		osheet.FRectNotIpgoOnly = notipgo
		osheet.FRectMinusOnly = minusjumun
		osheet.FRectshopdiv = shopdiv
	end if

	osheet.FtplGubun = tplgubun
	osheet.frecttotalyn = totalyn

	if designer<>"" then
		osheet.FRectItemGubun = itemgubun
		osheet.FRectItemID = itemid
		osheet.FRectItemOption = itemoption
		osheet.GetOrderSheetListByBrand
	else
		osheet.GetOrderSheetList
	end if

%>

<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
<script type="text/javascript">

function downloadOrder(masteridx, baljucode, shopid) {
	var popwin = window.open("/common/popOrderSheet_foreign_excel.asp?masteridx=" + masteridx + "&baljucode=" +baljucode + "&shopid=" +shopid,"ExcelOfflineOrderSheet","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function MakeJumun(){
	location.href="jumuninput.asp";
}

function PopSegumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('세금일 : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value = "segumil";
			frm.submit();
		}
	};
}

function PopIpgumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('입금일 : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value="ipkumil";
			frm.submit();
		}
	};
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v){
	window.open('popshopjumunsheet2.asp?idx=' + v + '&xl=on');
}

function MakeReJumun(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('출고 예정일 : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('미배송 주문서를 작성 하시겠습니까?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "remijumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function MakeReturn(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('출고 예정일 : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('반품 주문서를 작성 하시겠습니까?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "returnjumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function jsAddSheet(idx) {
	var ifrm = document.getElementById("ifrm");
	var frm = opener.document.frmMaster;
	var shopid = frm.shopid.value;
	var suplyer = frm.suplyer.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	ifrm.src = "doshopjumun.asp?idx=" + idx + "&mode=cpsheet&shopid=" + shopid + "&suplyer=" + suplyer + "&cwflag=" + cwflag;
}

function popSelectTargetShop(dftShopid,param1,param2){
    var popwin=window.open('/common/offshop/popShopSelect.asp?shopid='+dftShopid+'&param1='+param1+'&param2='+param2,'popShopSelect','width=400,height=200,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popRetShopid(ishopid,param1,param2){
    MakeDuplicateJumun(ishopid,param1);
}

function MakeDuplicateJumun(cpbaljuid,iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('출고 예정일 : ' + frmMaster.datestr.value + ' OK?')){ return };

    if (cpbaljuid.length<1){
        alert('출고처가 선택되지 않았습니다.');
        return
    }

	if (confirm('복사 주문서를 작성 하시겠습니까?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "duplicatejumun";
		frmMaster.cpbaljuid.value = cpbaljuid;
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function Popbalju(){
	var frm = document.frmlist;
	var idxarr = "";
	for (var i=0;i<frm.elements.length;i++){
		if ((frm.elements[i].name=="ck_all") && (frm.elements[i].checked)){
        	idxarr = idxarr + frm.elements[i].value + ",";
      	}
	}
	if (idxarr==""){
		alert('주문서를 선택하세요.');
		return;
	}else{
		frm.idxarr.value= idxarr;
		frm.target="_blank";
		frm.action="popoffbaljulist.asp"
		frm.submit();
	}
}

function jsStockMove() {
	var frm = document.frm;
	if (frm.shopid.value == "") {
		alert("먼저 매장(출발지)을 선택하세요.\n\n(검색조건에서 ShopID 를 입력 후 검색하세요.)");
		return;
	}

	var pop = window.open('pop_jumun_move.asp?menupos=' + frm.menupos.value + '&shopid=' + frm.shopid.value,'jsStockMove','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop.focus();
}

function jsCheckAll(obj) {
    var frm = document.frmlist;
    var checked = obj.checked;

	for (var i = 0; i < frm.elements.length; i++) {
        if (frm.elements[i].name == "ck_all") {
            frm.elements[i].checked = checked;
            AnCheckClick(frm.elements[i]);
        }
	}
}

function jsStockMoveSheet() {
    var frm = document.frmlist;
    var idxArr = '';
    var idx, statecd;

	for (var i = 0; i < frm.elements.length; i++) {
        if ((frm.elements[i].name == "ck_all") && (frm.elements[i].checked == true)) {
            // statecd, chk
            idx =  frm.elements[i].id.substring(3);
            statecd = document.getElementById('statecd' + idx).value;
            if (statecd >= '7') {
                alert('출고완료 내역이 있습니다.\n\n출고이전 전환 후 재고이동 가능합니다.');
                return;
            }

            idxArr = idxArr + ',' + idx;
        }
	}

    if (idxArr == '') {
        alert('선택된 주문이 없습니다.');
        return;
    }

    var pop = window.open('pop_jumun_move_by_sheet.asp?menupos=' + document.frm.menupos.value + '&idx=' + idxArr,'jsStockMoveSheet','width=800,height=300,scrollbars=yes,resizable=yes');
	pop.focus();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function pop_limiitcheck(alinkcode){
	var pop_limiitcheck = window.open('/admin/fran/poplimitcheckipgoNew.asp?alinkcode='+alinkcode,'pop_limiitcheck','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_limiitcheck.focus();
}

function smssendreg(masteridx){
	var popwin = window.open('/admin/fran/jumun_smssendreg.asp?masteridx='+masteridx+'&paymentgroup=ORDER','regsmssend','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="research" value="on">
		<input type="hidden" name="page" value="1">
		<input type="hidden" name="popupyn" value="<%= popupyn %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left">
				* 주문코드 : <input type="text" id="idBaljuCode" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8" onKeyUp="if (event.keyCode == 13) { frmsubmit(''); }">
				&nbsp;&nbsp;
				* ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				&nbsp;&nbsp;
				* 주문상태 :
				<% drawstatecd "statecd", statecd, " onchange='frmsubmit("""");'" %>
				<br>
				* 브랜드포함 : <% drawSelectBoxDesignerwithName "designer", designer %>
				&nbsp;&nbsp;
				* 날짜기준 :
				<% drawipgo_datefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			</td>

			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
			</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="left">
				<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >출고미처리만
     			&nbsp;&nbsp;
                <!--
     			<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >마이너스주문만
                -->
                * 주문구분 :
                <select class="select" name="minusjumun">
                    <option value="">전체</option>
                    <option value="N" <%= CHKIIF(minusjumun="N", "selected", "") %>>정상주문</option>
                    <option value="Y" <%= CHKIIF(minusjumun="Y", "selected", "") %>>반품주문</option>
                </select>
     			&nbsp;&nbsp;
     			* SHOP구분 :
     			<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >전체
				<input type="radio" name="shopdiv" value="j" <% if shopdiv="j" then response.write "checked" %> >직영
				<input type="radio" name="shopdiv" value="f" <% if shopdiv="f" then response.write "checked" %> >가맹점
				&nbsp;
				3PL 구분 : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
				&nbsp;
				* 상품코드 :
				<% if (designer <> "") then %>
				<input type="text" class="text" name="itemgubun" value="<%= itemgubun %>" size="2" maxlength="2">
				<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="16">
				<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4" maxlength="4">
				<% else %>
				브랜드입력시 검색가능
				<% end if %>
			</td>
		</tr>
	</form>
</table>
<!-- 검색 끝 -->

<br>

	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
		<tr>
			<td align="Left">
				<% If (popupyn <> "Y") Then %>
				<input type="button" class="button" value="주문서작성" onClick="MakeJumun();">
				<% End If %>
			</td>
			<td align="right">
                <input type="button" class="button" value="선택주문 재고이동" onClick="jsStockMoveSheet()">
				<input type="button" class="button" value="재고이동" onClick="jsStockMove()">
				<!--		<input type="button" class="button" value="선택주문 발주서보기" onClick="Popbalju()">	-->
			</td>
		</tr>
	</table>
	<!-- 액션 끝 -->

<form name="frmlist" method="post" style="margin:0px;">
<input type=hidden name="idxarr">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22">
		검색결과 : <b><%= osheet.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= osheet.FTotalpage %></b>
		<% if totalyn="Y" then %>
		<% if osheet.FResultCount >0 then %>
		&nbsp;
		합계 :
		총주문(소비자가) <b><%= FormatNumber(osheet.total_jumunsellcash,0) %></b>
		/ 총주문(공급가) <b><%= FormatNumber(osheet.total_jumunsuplycash,0) %></b>
		/ 총확정(공급가) <b><%= FormatNumber(osheet.total_totalsuplycash,0) %></b>
		<% end if %>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" onClick="jsCheckAll(this)"></td>
	<td>주문코드</td>
	<td>공급자</td>
	<td>공급받는자</td>
	<td>주문당시<br>화폐단위</td>
	<td>등록자<br>처리자<!--<br>주문자(SHOP)--></td>
	<td>주문상태</td>
	<!--<td>wholesale<br>결제상태</td>-->
	<!--<td>세금일<br>입금일</td>-->
	<td>주문일/<br>입고(요청)일</td>
	<!-- <td width=60>구분</td> -->
	<td>총주문액<br>(소비자가)</td>
	<td>총주문액<br>(공급가)</td>
	<td>확정금액<br>(공급가)</td>
	<td>총주문액<br>(매입가)</td>
	<td>확정금액<br>(매입가)</td>
	<td>해외총주문액<br>(소비자가)</td>
	<td>해외총주문액<br>(공급가)</td>
	<td>해외확정금액<br>(공급가)</td>
	<td>출고일</td>
	<td>송장번호</td>
	<td>내역서</td>
	<td>내역서(엑셀)</td>
	<td>
		<% If (popupyn <> "Y") Then %>
		바코드
		<% Else %>
		주문서<br />추가
		<% End If %>
	</td>
</tr>
<% if osheet.FResultCount >0 then %>
<% for i=0 to osheet.FResultcount-1 %>
<%
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
tot_jumunsuplycash = tot_jumunsuplycash + osheet.FItemList(i).Fjumunsuplycash
tot_totalsuplycash   = tot_totalsuplycash + osheet.FItemList(i).Ftotalsuplycash
tot_jumunbuycash = tot_jumunbuycash + osheet.FItemList(i).Fjumunbuycash
tot_totalbuycash   = tot_totalbuycash + osheet.FItemList(i).Ftotalbuycash
totaljumunforeign_sellcash = totaljumunforeign_sellcash + osheet.FItemList(i).fjumunforeign_sellcash
totaljumunforeign_suplycash = totaljumunforeign_suplycash + osheet.FItemList(i).fjumunforeign_suplycash
totalforeign_suplycash = totalforeign_suplycash + osheet.FItemList(i).ftotalforeign_suplycash
%>
<tr bgcolor="#FFFFFF">
    <input type="hidden" id="statecd<%= osheet.FItemList(i).Fidx %>" value="<%= osheet.FItemList(i).Fstatecd %>">
	<td width=16 rowspan=2><input type="checkbox" id="chk<%= osheet.FItemList(i).Fidx %>" name="ck_all" value="<%= osheet.FItemList(i).Fidx %>" onClick="AnCheckClick(this);"></td>
	<td rowspan=2 align="center"><a href="jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&oshopid=<%= shopid %>&ostatecd=<%= statecd %>&odesigner=<%= designer %>"><%= osheet.FItemList(i).Fbaljucode %></a></td>

	<% if osheet.FItemList(i).Ftargetid<>"10x10" then %>
	<td rowspan=2 align="center"><b><%= osheet.FItemList(i).Ftargetid %></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% else %>
	<td rowspan=2 align="center"><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% end if %>

	<td rowspan=2 align="center"><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)<!--<br>(<%= osheet.FItemList(i).Fregname %>)--></td>
	<td rowspan=2 align="center">
		<%= osheet.FItemList(i).fcurrencyUnit %>

		<% if osheet.FItemList(i).fsitename<>"" then %>
			<Br><%= osheet.FItemList(i).fsitename %>
		<% end if %>
	</td>
	<td rowspan=2 align="center">
		<%= osheet.FItemList(i).Fregname %><br>
		<%= osheet.FItemList(i).Ffinishname %>
	</td>
	<td rowspan=2 align="center">
		<font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font>
		<br><%= osheet.FItemList(i).FAlinkCode %>
	</td>
	<!--<td rowspan=2 align="center">
		<%'= osheet.FItemList(i).getOrderpaymentstatus %>
		<% 'if osheet.FItemList(i).fsitename="WSLWEB" then %>
			<% 'if osheet.FItemList(i).fsmssenddate<>"" and not(isnull(osheet.FItemList(i).fsmssenddate)) then %>
				<br>문자발송:
				<br><%'= left(osheet.FItemList(i).fsmssenddate,10) %>
				<br><%'= mid(osheet.FItemList(i).fsmssenddate,12,22) %>
			<% 'else %>
				<br><input type="button" onclick="smssendreg('<%'= osheet.FItemList(i).Fidx %>')" value="문자발송" class="button">
			<% 'end if %>
		<% 'end if %>
	</td>-->
	<!--<td align="center">
			<% 'if IsNULL(osheet.FItemList(i).Fsegumdate) then %>
			<div align="right"><a href="javascript:PopSegumil(frmMaster,'<%'= osheet.FItemList(i).Fidx %>',frmMaster.datestr);"><img src="/images/calicon.gif" border=0></a></div>
			<% 'else %>
			<a href="javascript:PopSegumil(frmMaster,'<%'= osheet.FItemList(i).Fidx %>',frmMaster.datestr);"><%'= osheet.FItemList(i).Fsegumdate %></a>
			<% 'end if %>
			</td>-->
	<td align="center"><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
	<!-- <td align="center"><%= osheet.FItemList(i).GetDivCodeName %></td> -->
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).fjumunforeign_sellcash,2) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).fjumunforeign_suplycash,2) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).ftotalforeign_suplycash,2) %></td>
	<td align="center"><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
	<td align="center"><%= Left(osheet.FItemList(i).Fsongjangno,10) %></td>
	<td rowspan=2 align="center" width=40>
		<!--
				<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a> <a href="javascript:ExcelSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexcel.gif" width=21 border=0></a>
			-->

		<a href="javascript:ViewOfflineOrderSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a>
	</td>
	<td rowspan=2 align="center" width=210>
		<input type="button" onclick="ExcelOfflineOrderSheet('<%= osheet.FItemList(i).FIdx %>');" value="물류코드" class="button">
		<input type="button" onclick="ExcelOfflineOrderSheetpublic('<%= osheet.FItemList(i).FIdx %>');" value="범용" class="button">
		<input type="button" onclick="downloadOrder('<%= osheet.FItemList(i).FIdx %>','<%= osheet.FItemList(i).Fbaljucode %>','<%= osheet.FItemList(i).Fbaljuid %>');" value="상품목록" class="button">
		<%
		'/출고완료인거
		if osheet.FItemList(i).Fstatecd = "7" then
		%>
		<%
		'/반품주문인거
		if osheet.FItemList(i).Ftotalsellcash < 0 then
		%>
		<input type="button" onclick="pop_limiitcheck('<%= osheet.FItemList(i).FAlinkCode %>');" value="한정" class="button">
		<% end if %>
		<% end if %>
	</td>
	<td rowspan=2 align="center">
		<% If (popupyn <> "Y") Then %>
			<input type="button" class="button" value="출력" onclick="printbarcode_off('JUMUN', '', '', '', '', '', '<%= osheet.FItemList(i).Fidx %>', '', '');">
		<% Else %>
			<input type="button" class="button" value="추가" onclick="jsAddSheet(<%= osheet.FItemList(i).Fidx %>);">
		<% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<!--<td align="center">
			<% 'if IsNULL(osheet.FItemList(i).Fipkumdate) then %>
			<div align="right"><a href="javascript:PopIpgumil(frmMaster,'<%'= osheet.FItemList(i).Fidx %>',frmMaster.datestr);"><img src="/images/calicon.gif" border=0></a></div>
			<% 'else %>
			<a href="javascript:PopIpgumil(frmMaster,'<%'= osheet.FItemList(i).Fidx %>',frmMaster.datestr);"><%'= osheet.FItemList(i).Fipkumdate %></a>
			<% 'end if %>
			</td>-->
	<td align="center">
		<% if IsNULL(osheet.FItemList(i).FIpgodate) then %>
		<%= Left(osheet.FItemList(i).Fscheduledate,10) %>
		<% else %>
		<%= Left(osheet.FItemList(i).FIpgodate,10) %>
		<% end if %>
	</td>
	<td colspan=9><font color="#777777"><%= DdotFormat(osheet.FItemList(i).Fbrandlist,30) %></font></td>
	<td>
		<a href="javascript:MakeReJumun('<%= osheet.FItemList(i).Fidx %>')">재작성</a>
		<% if (C_ADMIN_AUTH) then %> <!--  and (osheet.FItemList(i).FStateCD=" ") -->
		&nbsp;
		<a href="javascript:MakeReturn('<%= osheet.FItemList(i).Fidx %>')">반품</a>

		<br><a href="javascript:popSelectTargetShop('','<%= osheet.FItemList(i).Fidx %>','')">복사</a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td></td>
	<td align="center">총계</td>
	<td colspan=7></td>
	<td align="right"><b><%= formatNumber(totaljumunsellcash,0) %></b></td>
	<td align="right"><b><%= formatNumber(tot_jumunsuplycash,0) %></b></td>
	<td align="right"><b><%= formatNumber(tot_totalsuplycash,0) %></b></td>
	<td align="right"><b><%= formatNumber(tot_jumunbuycash,0) %></b></td>
	<td align="right"><b><%= formatNumber(tot_totalbuycash,0) %></b></td>
	<td align="right"><b><%= formatNumber(totaljumunforeign_sellcash,2) %></b></td>
	<td align="right"><b><%= formatNumber(totaljumunforeign_suplycash,2) %></b></td>
	<td align="right"><b><%= formatNumber(totalforeign_suplycash,2) %></b></td>
	<td colspan=5></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=22 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22" align="center">
		<% if osheet.HasPreScroll then %>
		<a href="javascript:frmsubmit('<%= osheet.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
		[pre]
		<% end if %>

		<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
		<% if i>osheet.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:frmsubmit('<%= i %>');">[<%= i %>]</a>
		<% end if %>
		<% next %>

		<% if osheet.HasNextScroll then %>
		<a href="javascript:frmsubmit('<%= i %>');">[next]</a>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>
</form>
<form name="frmMaster" method=post action="doshopjumun.asp" target="_blank">
	<input type="hidden" name="idx" value="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="datestr" value="">
	<input type="hidden" name="cpbaljuid" value="">
</form>

<iframe id="ifrm" border="0" scrolling="no" class="frame" width="0" height="0"></iframe>
<%
set osheet = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
