<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 개별 입출고 리스트
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->

<%
dim chargeid ,shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2
dim fromDate,toDate,tmpdate ,page,notipgo,datesearchtype ,oipchul , moveipchulyn
dim totcnt, totsum1, totsum2 ,i
	page = request("page")
	chargeid = request("chargeid")
	shopid = request("shopid")
	notipgo = request("notipgo")
	datesearchtype = request("datesearchtype")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	moveipchulyn = request("moveipchulyn")

if page="" then page=1
if datesearchtype="" then datesearchtype="scheduledt"

if (C_IS_SHOP) then
	'직영/가맹점
	shopid = C_STREETSHOPID

else
	if (C_IS_Maker_Upche) then
		chargeid = session("ssBctId")
	else
		if not(C_ADMIN_USER) then
		else
		end if
	end if
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))-1, Cstr(day(now())))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then
    toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))+1, Cstr(day(now())))
    tmpdate = dateAdd("d",toDate,-1)
    yyyy2 = Cstr(Year(tmpdate))
    mm2 = Cstr(Month(tmpdate))
    dd2 = Cstr(day(tmpdate))
else
    toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

set oipchul = new CShopIpChul
	oipchul.FPageSize = 50
	oipchul.FCurrPage = page
	oipchul.FRectDatesearchtype = datesearchtype
	oipchul.FRectStartDay = CStr(fromDate)
	oipchul.FRectEndDay = CStr(toDate)
	oipchul.FRectChargeId = chargeid
	oipchul.FRectShopId = shopid
	oipchul.FRectNotIpgo = notipgo
	oipchul.FRectmoveipchulyn = moveipchulyn
	oipchul.GetIpChulMasterList
%>

<script language='javascript'>

function popsimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=400,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function ReSearch(page){
	frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function IpgoStateChange(v){
	alert('관리자만 수정 가능한 메뉴입니다.');
	var popwin = window.open('/common/offshop/pop_offipgostatechange.asp?idx=' + v,'pop_offipgostatechange','width=480,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function DelThis(v,shopid,chargeid){
	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		document.frmipchul.shopid.value=shopid;
		document.frmipchul.chargeid.value=chargeid;	
		document.frmipchul.idx.value=v;
		document.frmipchul.submit();
	}
}

function ReThis(v,comp){
	var ret = confirm('미 입고로 전환 하시겠습니까?');

	if (ret){
		document.frmipchul.mode.value="miipgo";
		document.frmipchul.idx.value=v;
		document.frmipchul.submit();
	}
}

function IpThis(v,comp,shopid,chargeid){

	if (!confirm('입고 확인 후에는 내역수정이 불가능 합니다.\n내역이 차이가 있을경우 업체에 연락하여 수정 후 진행하시기 바랍니다.\n\n - 진행하시겠습니까?(다음창에서 입고일을 선택하세요)')) return;
	if (!calendarOpen4(comp,'입고일','')) return;

	var ret = confirm('입고일 : ' + comp.value + '\n입고 확인 하시겠습니까?');

	if (ret){
		document.frmipchul.shopid.value=shopid;
		document.frmipchul.chargeid.value=chargeid;
		document.frmipchul.mode.value="ipgook";
		document.frmipchul.idx.value=v;
		document.frmipchul.execdate.value = comp.value;

		document.frmipchul.submit();
	}
}

//입고 요청
function ReqIpChulInput(){
	var chargeid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (chargeid==''){
		alert('공급처를 먼저 선택해 주세요');
		frm.chargeid.focus();
		return;
	}

	document.location = "/common/offshop/shop_ipchulinput.asp?menupos=<%= menupos %>&chargeid=" + chargeid + "&shopid=" + shopid + "&isreq=Y";
}

function ipChulInput(){
	var chargeid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (chargeid==''){
		alert('공급처를 먼저 선택해 주세요');
		frm.chargeid.focus();
		return;
	}

	document.location = "/common/offshop/shop_ipchulinput.asp?menupos=<%= menupos %>&chargeid=" + chargeid + "&shopid=" + shopid ;
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?idx=' + v,'ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheetXL(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?idx=' + v + '&xl=on','ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopBarCodePrint(v){
	document.iiframe.location.href = "/common/offshop/iframebarcode.asp?idxlist=" + v;
}

function SelBarCodePrt(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 내역이 없습니다.');
		return;
	}

	var idxArr="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				idxArr = idxArr + frm.idx.value + ","
			}
		}
	}

	if (idxArr.substr(idxArr.length-1,1)==","){
		idxArr = idxArr.substr(0,idxArr.length-1);
	}
	PopBarCodePrint(idxArr);
}

function SelImagePrt(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 내역이 없습니다.');
		return;
	}

	var idxArr="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				idxArr = idxArr + frm.idx.value + ","
			}
		}
	}

	if (idxArr.substr(idxArr.length-1,1)==","){
		idxArr = idxArr.substr(0,idxArr.length-1);
	}
	var popwin;
	popwin = window.open('popshopImagelist.asp?idx=' + idxArr,'shopitem','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function sendSMSEmail(idesigner,iidx){
	var popwin = window.open("/admin/offshop/popupchejumunsms_off.asp?designer=" + idesigner + "&idx=" + iidx,"popupchejumunsms","width=600 height=500,scrollbars=yes,resizabled=yes");
	popwin.focus();
}

//매장 재고 이동
function ipchulmove(isreq){
	var makerid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (makerid==''){
		alert('공급처를 먼저 선택해 주세요');
		frm.chargeid.focus();
		return;
	}

	var popwin = window.open('/common/offshop/shop_ipchuldetail_move.asp?menupos=<%= menupos %>&isreq='+isreq+'&firstshopid='+shopid+'&makerid='+makerid,'popwin','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();		
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		공급처 :<% drawSelectBoxDesignerwithName "chargeid",chargeid %>		
		매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		매장재고이동:<% Call drawSelectBoxUsingYN("moveipchulyn",moveipchulyn) %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >입고대기전체보기
		<select name="datesearchtype">
		<option value="scheduledt" <% if datesearchtype="scheduledt" then response.write "selected" %> >입고예정일
		<option value="execdt" <% if datesearchtype="execdt" then response.write "selected" %> >입고일
		</select>
		 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>	
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 업체에서 매장으로 직접 입출고 하는경우 사용하는 메뉴입니다. (업체 특정인경우만 사용가능)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;- 텐바이텐 물류센터로 입고하는경우 물류센터에서 입출고 내역을 입력합니다.<br>
			<% 
			'/업체가 아닐경우
			if not(C_IS_Maker_Upche) then
				'/직영점 이거나 본사일경우
				if getoffshopdiv(shopid) = "1" or C_ADMIN_USER then
			%>
				&nbsp;&nbsp;&nbsp;&nbsp;- <font color="red">매장재고이동</font>인경우 출발매장(<font color="red">마이너스주문</font>)과 도착매장(<font color="red">입고주문</font>)에 주문이 각각 생성됩니다<br>
			<%
				end if
			end if
			%>
			&nbsp;&nbsp;&nbsp;&nbsp;- <font color="red">반품</font>인경우 수량을 <font color="red">마이너스</font>로 입력합니다.<br>			
		※ 입고상태 :<br>
			&nbsp;&nbsp;&nbsp;&nbsp;1. <b>입고대기</b> - 업체에서 매장으로 상품을 보낸상태입니다.(내역수정가능)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;2. <b>매장 입고확인</b> - 매장에서 입고를 확인한 상태입니다.(내역수정불가)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;3. <b>입고확정(업체 입고확인)</b> - 매장 입고확인 후 업체에서 입고 확인한 상태입니다.(내역수정불가)<br>
	</td>
	<td align="right">
	    <input type="button" class="button" value="선택내역바코드출력" onclick="SelBarCodePrt()">
	    <!-- <input type="button" value="선택내역이미지출력" onclick="SelImagePrt()"> -->	
	    <input type="button" class="button" value="입고 요청 입력  [발주서 작성]" onclick="ReqIpChulInput()">
	    <input type="button" class="button" value="입고/반품 입력" onclick="ipChulInput()">
		<% 
		'/업체가 아닐경우
		if not(C_IS_Maker_Upche) then
			'/직영점 이거나 본사일경우
			if getoffshopdiv(shopid) = "1" or C_ADMIN_USER then
		%>
				<input type="button" onclick="ipchulmove('M');" class="button" value="재고이동">
		<% 
			end if
		end if
		%>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oipchul.FResultCount %></b> <%= Page %>/<%= oipchul.FTotalPage %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>입출코드</td>
	<td><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td>입력</td>
	<td>발주</td>
	<td>공급처</td>
	<td>가맹점</td>
	<td>판매가</td>
	<td>공급가</td>
	<td>마진</td>
	<td>입고<br>예정일</td>
	<td>입고일</td>
	<td>입고상태</td>
	<td>내역<br>수정</td>
	<td>입고<br>내역서</td>
	<td>바코드<br>출력</td>
	<td>삭제</td>
	<td>입고<br>확인</td>
	<td>비고</td>
</tr>
<% if oipchul.FResultCount > 0 then %>
<% for i=0 to oipchul.FResultcount -1 %>
<%
totcnt = totcnt + 1
totsum1 = totsum1 + oipchul.FItemList(i).FTotalSellcash
totsum2 = totsum2 + oipchul.FItemList(i).FTotalSuplycash
%>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="idx" value="<%= oipchul.FItemList(i).FIdx %>">
<tr bgcolor="#FFFFFF" align="center">
    <td ><%= oipchul.FItemList(i).FIdx %></td>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><font color="<%= oipchul.FItemList(i).getInputAreaColor %>"><%= oipchul.FItemList(i).getInputAreaStr %></font></td>
	<td> 
		<% if (oipchul.FItemList(i).FisbaljuExists="Y") then %>
			발주
		<% elseif (oipchul.FItemList(i).FisbaljuExists="M") then %>
			재고이동
		<% end if %>
	</td>
	<td ><a href="javascript:popsimpleBrandInfo('<%= oipchul.FItemList(i).FChargeID %>');"><%= oipchul.FItemList(i).FChargeID %></a></td>
	<td ><a href="/common/offshop/shop_ipchuldetail.asp?idx=<%= oipchul.FItemList(i).FIdx %>&menupos=<%= menupos %>"><%= oipchul.FItemList(i).FShopName %></a></td>
	<td align="right"><%= FormatNumber(oipchul.FItemList(i).FTotalSellcash,0) %></td>
	<td align="right"><%= FormatNumber(oipchul.FItemList(i).FTotalSuplycash,0) %></td>
	<td align=center>
	<% if oipchul.FItemList(i).FTotalSellcash<>0 then %>
	<%= 100-CLng(oipchul.FItemList(i).FTotalSuplycash/oipchul.FItemList(i).FTotalSellcash*100*100)/100 %> 
	<% end if %>
	</td>
	<td align=center><%= oipchul.FItemList(i).FScheduleDt %></td>
	<td align=center><%= oipchul.FItemList(i).FExecDt %></td>
	<td >
	<input type=hidden name=yyyymmdd>
    	<% if (C_ADMIN_AUTH) or (C_OFF_AUTH) or (session("ssBctId") = "sangmi")  then %>
    	<!-- 관리자만 상태변경가능 -->
		<a href="javascript:IpgoStateChange('<%= oipchul.FItemList(i).FIdx %>')"><font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font></a>
		<% else %>
		<font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font>
		<% end if %>
	</td>
	<td ><a href="/common/offshop/shop_ipchuldetail.asp?idx=<%= oipchul.FItemList(i).FIdx %>&menupos=<%= menupos %>"><img src="/images/icon_search.jpg" border="0" width="16"></a></td>
	<td align="center">
		<a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexplorer.gif" border="0" width="21"></a>
		<a href="javascript:PopIpgoSheetXL('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexcel.gif" border="0" width="21"></a>
	</td>
	<td align="center"><a href="javascript:PopBarCodePrint('<%= oipchul.FItemList(i).FIdx %>');"><img src="/images/icon_print02.gif" border="0" ></a></td>
	<td align="center">
		<% if (oipchul.FItemList(i).FStatecd>=7) then %>

		<% else %>
			<a href="javascript:DelThis('<%= oipchul.FItemList(i).FIdx %>','<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>')">x</a>
		<% end if %>
	</td>
	<td>
		<% if (oipchul.FItemList(i).FStatecd>7) then %>
		<% else %>
			<input type="button" value="확정" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>')" class="button">
		<% end if %>
	</td>
	<td width=200>
		<%
		'/업체가 아닐경우
		if not(C_IS_Maker_Upche) then
			if (oipchul.FItemList(i).fsendsms="N") and (isnull(oipchul.FItemList(i).FExecDt)) and (oipchul.FItemList(i).Fstatecd < 0) then
		%>
				<input type="button" class="button" value="SMS" onclick="sendSMSEmail('<%= oipchul.FItemList(i).FChargeID %>','<%= oipchul.FItemList(i).Fidx %>')">
		<%
			end if
		end if

		if oipchul.FItemList(i).fipchulmoveidx <> "" then
		%>
			<br>관련재고이동입출코드 : <%= oipchul.FItemList(i).fipchulmoveidx %>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">총 <%= FormatNumber(totcnt,0) %>건</td>
	<td align="right"><%= FormatNumber(totsum1,0) %></td>
	<td align="right"><%= FormatNumber(totsum2,0) %></td>
	<td colspan="10"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
		<% if oipchul.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= oipchul.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + oipchul.StartScrollPage to oipchul.FScrollCount + oipchul.StartScrollPage - 1 %>
			<% if i>oipchul.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if oipchul.HasNextScroll then %>
			<a href="javascript:ReSearch('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<form name="frmipchul" method="post" action="/common/offshop/shopipchul_process.asp">
	<input type="hidden" name="mode" value="delmaster">
	<input type="hidden" name="idx" value="">
	<input type=hidden name="execdate" >
	<input type=hidden name="shopid" >
	<input type=hidden name="chargeid" >		
</form>
<iframe name="iiframe" src="" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
</table>

<%
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->