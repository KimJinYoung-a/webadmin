<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim dateback

nowdate = Left(CStr(now()),10)


yyyy1   = requestCheckVar(request("yyyy1"), 32)
mm1     = requestCheckVar(request("mm1"), 32)
dd1     = requestCheckVar(request("dd1"), 32)
yyyy2   = requestCheckVar(request("yyyy2"), 32)
mm2     = requestCheckVar(request("mm2"), 32)
dd2     = requestCheckVar(request("dd2"), 32)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm2-1, dd2)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim cknodate
cknodate = requestCheckVar(request("cknodate"), 32)

dim page
dim ojumun

page = requestCheckVar(request("page"), 32)
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FPageSize = 200
ojumun.FCurrPage = page
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerDateBaljuList

dim ix,iy
%>
<script type='text/javascript'>

function ViewOrderDetail(frm){
	var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}

function ViewItem(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
}


function NextPage(ipage){
	document.frmsearch.page.value= ipage;
	document.frmsearch.submit();
}


function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.orderserial.length>1){
		for(i=0;i<frm.orderserial.length;i++){
			frm.orderserial[i].checked = comp.checked;
			AnCheckClick(frm.orderserial[i]);
		}
	}else{
		frm.orderserial.checked = comp.checked;
		AnCheckClick(frm.orderserial);
	}
}

function CheckNBaljusu(){
	var frm = document.frmbalju;
	var pass = false;

    if(!frm.orderserial.length) {
		pass = frm.orderserial.checked;
	} else {
    	for (var i=0;i<frm.orderserial.length;i++){
    	    pass = (pass||frm.orderserial[i].checked);
    	}
    }

	if (!pass) {
		alert("선택 주문이 없습니다.");
		return;
	}

	var ret = confirm("[ 개인정보 취급 주의 안내 ]\n\n[ 수집 및 저장 단계]\n● 배송, A/S 등의 목적으로 다운로드 받은 고객의 개인정보는 관리 담당자 지정, 엑셀 파일에 암호설정등의 조치를 통해 인가되지 않은 제 3자 등에게 불법 유출되지 않도록 철저한 관리가 필요 합니다.\n\n[ 이용 및 폐기 단계 ]\n● 고객의 개인정보는 배송 및 A/S 등의 수집시 고지한 범위를 넘어서 제 3자에게 제공하거나 판매 목적의 광고, 할인 판매 안내등의 영업 목적으로 사용은 법적으로 금지되어 있으며 이를 어길 경우 법에 따라 처벌을 받습니다.\n● 보유하고 있는 개인정보 중 업무 목적이 달성되어 더 이상 필요 없을 경우 관련 파일 또는 인쇄된 문서를 폐기하여 주시기 바랍니다.\n\n\n상기 내용을 이해하였고 개인정보 유출 사고가 발생되지 않도록 주의 하겠습니다.\n\n===============================================================\n\n\n현재 조회조건으로 다운로드하려면 확인을 누르고 잠시 기다려주세요.\n조회결과가 많을 경우 다운로드가 오래 걸릴 수 있습니다.\n\n\n선택 주문을 확인 하시겠습니까?");

	if (ret){
 		frm.action="selectbaljulist.asp";
		frm.submit();

	}
}

function CheckNBaljusuNew(){
	var frm = document.frmbalju;
	var pass = false;

    if(!frm.orderserial.length) {
		pass = frm.orderserial.checked;
	} else {
    	for (var i=0;i<frm.orderserial.length;i++){
    	    pass = (pass||frm.orderserial[i].checked);
    	}
    }

	if (!pass) {
		alert("선택 주문이 없습니다.");
		return;
	}

	var ret = confirm("[ 개인정보 취급 주의 안내 ]\n\n[ 수집 및 저장 단계]\n● 배송, A/S 등의 목적으로 다운로드 받은 고객의 개인정보는 관리 담당자 지정, 엑셀 파일에 암호설정등의 조치를 통해 인가되지 않은 제 3자 등에게 불법 유출되지 않도록 철저한 관리가 필요 합니다.\n\n[ 이용 및 폐기 단계 ]\n● 고객의 개인정보는 배송 및 A/S 등의 수집시 고지한 범위를 넘어서 제 3자에게 제공하거나 판매 목적의 광고, 할인 판매 안내등의 영업 목적으로 사용은 법적으로 금지되어 있으며 이를 어길 경우 법에 따라 처벌을 받습니다.\n● 보유하고 있는 개인정보 중 업무 목적이 달성되어 더 이상 필요 없을 경우 관련 파일 또는 인쇄된 문서를 폐기하여 주시기 바랍니다.\n\n\n상기 내용을 이해하였고 개인정보 유출 사고가 발생되지 않도록 주의 하겠습니다.\n\n===============================================================\n\n\n현재 조회조건으로 다운로드하려면 확인을 누르고 잠시 기다려주세요.\n조회결과가 많을 경우 다운로드가 오래 걸릴 수 있습니다.\n\n\n선택 주문을 확인 하시겠습니까?");

	if (ret){
 		frm.action="selectbaljulistNew.asp";
		frm.submit();

	}
}

</script>


<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
	<input type="hidden" name="page" value="1" />
	<input type="hidden" name="menupos" value="<%= menupos %>" />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left" bgcolor="#FFFFFF">
				<input type="radio" name="" value="" checked />출고요청 주문리스트
				<!-- <input type="radio" name="" value="">요청이전 주문리스트(주문접수 포함) -->
			</td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="검색" onClick="document.frm.submit();" />
			</td>
		</tr>
	</table>
</form>


<!-- 검색 시작 -->
<form name="frmsearch" method="get" action="">
	<input type="hidden" name="page" value="1" />
	<input type="hidden" name="menupos" value="<%= menupos %>" />
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px;">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
    		<input type="button" class="button" value="전체선택" onclick="document.frmbalju.chkAll.checked=true;switchCheckBox(document.frmbalju.chkAll)" />
			&nbsp;
			<input type="button" class="button" value="선택주문확인" onclick="CheckNBaljusu()" />
			&nbsp;
			<input type="button" class="button" value="선택주문확인(New)" onclick="CheckNBaljusuNew()" />
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frmbalju" method="post" style="margin:0;">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top:10px;">
		<tr bgcolor="FFFFFF">
			<td height="25" colspan="15">
				검색결과 : <b><% = ojumun.FTotalCount %></b>
				&nbsp;
				페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    		<td width="30"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
			<td width="70">주문번호</td>
			<td width="55">주문자</td>
			<td width="55">수령인</td>
			<td width="50">상품코드</td>
			<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
			<td width="30">수량</td>
			<td width="65">주문일</td>
			<td width="65">입금확인일</td>
			<td width="65">출고기준일<!--주문통보일--></td>
			<td width="40">경과일</td>
		</tr>
		<% if ojumun.FresultCount<1 then %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center">[검색결과가 없습니다.]</td>
		</tr>
		<% else %>

		<% for ix=0 to ojumun.FresultCount-1 %>
		<tr align="center" class="a" bgcolor="#FFFFFF">
			<td>
				<!-- detail Index -->
				<input type="checkbox" name="orderserial"  onClick="AnCheckClick(this);" value="<% =ojumun.FMasterItemList(ix).Fidx %>">
			</td>
			<td><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
			<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
			<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
			<td><%= ojumun.FMasterItemList(ix).FitemID %></td>
			<td align="left">
				<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
				<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
				<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
				<% end if %>
			</td>
			<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).Fregdate %>"><%= left(ojumun.FMasterItemList(ix).Fregdate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= left(ojumun.FMasterItemList(ix).FIpkumdate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FMasterItemList(ix).Fbaljudate %>"><%= left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></acronym></td>
			<td>
				<% if IsNULL(ojumun.FMasterItemList(ix).Fbaljudate) then %>
				D+0
				<% elseif datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) )>2 then %>
				<font color="red"><b>D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
				<% else %>
				D+<%= datediff("d",(left(ojumun.FMasterItemList(ix).Fbaljudate,10)) , (left(now,10)) ) %>
				<% end if %>
			</td>
		</tr>

		<% next %>
		<% end if %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
				<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
				[pre]
				<% end if %>
				<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
				<% next %>

				<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	</table>
</form>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
