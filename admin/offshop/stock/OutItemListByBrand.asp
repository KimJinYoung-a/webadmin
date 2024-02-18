<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 정리대상 상품 포함 브랜드
' History : 2011.08 서동석 생성
'			2020.06.02 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shopstockClearCls.asp"-->
<%

dim shopid, makerid,  research, params, errExist,ipchulcode
dim cType, CLDiv, LstYYYYMM
shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
research     = RequestCheckVar(request("research"),2)
cType      = RequestCheckVar(request("cType"),10)
CLDiv      = RequestCheckVar(request("CLDiv"),10) ''재고 기준
LstYYYYMM  = RequestCheckVar(request("LstYYYYMM"),7)
errExist   = RequestCheckVar(request("errExist"),10)
ipchulcode = RequestCheckVar(request("ipchulcode"),10)
params       = "shopid="&shopid&"&makerid="&makerid&"&research="&research&"&cType="&cType&"&CLDiv="&CLDiv&"&LstYYYYMM="&LstYYYYMM&"&errExist="&errExist&"&ipchulcode="&ipchulcode


if (CLDiv="") then CLDiv="C"

if not(C_ADMIN_USER or C_IS_OWN_SHOP) then
    response.write "권한이 없습니다."
    dbget.close() : response.end
end if

dim part_sn
part_sn = session("ssAdminPsn")

dim oOffStock
set oOffStock = new CShopStockClear
oOffStock.FRectShopID   = shopid
oOffStock.FRectMakerID  = makerid
if (CLDiv="L") then
    oOffStock.FRectLastYYYYMM = LstYYYYMM
end if
oOffStock.FRectOnlyerrExist = errExist
oOffStock.GetShopStockClearBrandDetail

Dim i, TotErrrealcheckno, TotSampleNo
Dim mayStockDate : mayStockDate = Left(now(),10)
Dim iclearTypeName
IF (cType="C") then
    iclearTypeName="재고조정"
ELSEIF (cType="M") then
    iclearTypeName="오차조정"    ''사용안함.
ELSEIF (cType="L") then
    iclearTypeName="오차로스처리"
ELSEIF (cType="S") then
    iclearTypeName="샘플"
ENd IF

'// 어디서 쓰이는지 모름, skyer9, 2016-03-14
if (cType <> "L") and (cType <> "S") then
	response.write "시스템팀 문의!!"
	dbget.close()
	response.end
end if


if (CLDiv="L") and (LstYYYYMM<>"") then
    mayStockDate = Left(dateAdd("d",-1,dateAdd("m",1,(LstYYYYMM+"-01"))),10)
end if


Dim sqlStr, ArrList1, ArrList2

sqlStr = " select top 5 SD.YYYYMM,SD.comm_cd,C.comm_name,SD.defaultmargin,SD.defaultSuplymargin"
sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_shop_designer SD"
sqlStr = sqlStr & " 	left Join db_jungsan.dbo.tbl_jungsan_comm_code C"
sqlStr = sqlStr & " 	on SD.comm_cd=C.comm_cd"
sqlStr = sqlStr & " where shopid='"&shopid&"'"
sqlStr = sqlStr & " and makerid='"&makerid&"'"
sqlStr = sqlStr & " order by yyyymm desc"

rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
IF Not (rsget.EOF OR rsget.BOF) THEN
	ArrList1 = rsget.getRows()
END IF
rsget.Close

sqlStr = " select top 5 m.yyyymm,d.gubuncd,C.comm_name,sum(d.suplyprice*d.itemno) "
sqlStr = sqlStr & " from db_jungsan.dbo.tbl_off_jungsan_master m"
sqlStr = sqlStr & " 	Join db_jungsan.dbo.tbl_off_jungsan_detail d"
sqlStr = sqlStr & " 	on m.idx=d.masteridx"
sqlStr = sqlStr & " 	left Join db_jungsan.dbo.tbl_jungsan_comm_code C"
sqlStr = sqlStr & " 	on d.gubuncd=C.comm_cd"
sqlStr = sqlStr & " where m.makerid='"&makerid&"'"
sqlStr = sqlStr & " and d.shopid='"&shopid&"'"
sqlStr = sqlStr & " group by m.yyyymm,d.gubuncd,C.comm_name"
sqlStr = sqlStr & " order by m.yyyymm desc"

rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
IF Not (rsget.EOF OR rsget.BOF) THEN
	ArrList2 = rsget.getRows()
END IF
rsget.Close

dim cnt
dim currItemNo

dim job_sn
	job_sn = session("ssAdminPOsn")
%>
<script languag='javascript'>
function popShopCurrentStock(shopid,barcode){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&barcode=' + barcode ,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function reCalcuLoss(comp,i){
    var frm = comp.form;

    if (frm.cksel.length){
        frm.SUBTTLrealcheckErrRemain[i].value = frm.realcheckErr[i].value*1+frm.AssignrealcheckErr[i].value*1;
        frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
    }else{
        frm.SUBTTLrealcheckErrRemain.value = frm.realcheckErr.value*1+frm.AssignrealcheckErr.value*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
    }

    summaryTotal(frm);
}

function summaryTotal(frm){
return;
    var ttlsum = 0;
    var itemcnt = 0;
    var remaincnt = 0;
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemcnt+=frm.AssignrealcheckErr[i].value*1;
                remaincnt+=frm.SUBTTLrealcheckErrRemain[i].value*1;
                ttlsum+=frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
    }else{
        if (frm.cksel.checked){
            itemcnt+=frm.AssignrealcheckErr.value*1;
            remaincnt+=frm.SUBTTLrealcheckErrRemain.value*1;
            ttlsum+=frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
        }
    }
    frm.TTLrealcheckErr.value = itemcnt;
    frm.TTLrealcheckErrRemain.value = remaincnt;
    frm.TTLshopsuplycash.value = ttlsum;
}

function chkALL(comp){
    var frm = comp.form;
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (!frm.cksel[i].disabled){
                if (frm.AssignrealcheckErr[i].value*1!=0){
                    frm.cksel[i].checked=comp.checked;
                    AnCheckClick(frm.cksel[i]);
                }

                if (!comp.checked){
                    frm.cksel[i].checked=comp.checked;
                    AnCheckClick(frm.cksel[i]);
                }
            }
        }
    }else{
        if (!frm.cksel.disabled){
            if (frm.AssignrealcheckErr.value*1!=0){
                frm.cksel.checked=comp.checked;
                AnCheckClick(frm.cksel);
            }

            if (!comp.checked){
                frm.cksel.checked=comp.checked;
                AnCheckClick(frm.cksel);
            }
        }
    }
    summaryTotal(frm);
}

function AssignErrLoss(){
    <% if C_ADMIN_AUTH or C_OFF_AUTH or C_Relationship_Part or C_MngPart then %>
	<% elseif C_IS_SHOP then %>
		<% 
		if C_IS_OWN_SHOP then
			'/job_sn 3:본부장 , 6:점장/매니저 , 10:Chief manager , 11:Manager
			if job_sn = 3 or job_sn = 6 or job_sn = 10 or job_sn = 11 then
		%>
			<% else %>
				alert('점장 권한 부터 사용 가능합니다.');
				return;
			<% end if %>
		<% end if %>
	<% else %>
        alert('현재 관리자 권한만 가능합니다.');
        return;
    <% end if %>

    var frm = document.frmArr;
    frm.lossDate.value = document.frmStockDt.stockdate.value;

    if (!chkExesits(frm.cksel)){
        alert('선택 내역이 없습니다.');
        return;
    }

    if (document.frmStockDt.losstype.value.length<1){
        alert('로스타입을 선택하세요.');
        return;
    }

    if (document.frmStockDt.losstype.value=="L"){
        if (!confirm('로스처리(정산반영)으로 선택한경우 로스처리반영단가로 정산됩니다.(특정상품의 경우) 계속하시겠습니까?')){
            return;
        }
    }

    if (confirm('매장 로스 출고 반영 하시겠습니까?')){
        frm.losstype.value=document.frmStockDt.losstype.value;
        frm.mode.value="lossact";
        frm.submit();
    }
}

function chkExesits(comp){
    var frm = comp.form;

    if (comp.length){
        for (var i=0;i<comp.length;i++){
            if (comp[i].checked){
                return true;
            }
        }
    }else{
        if (comp.checked){
            return true;
        }
    }
    return false;
}

function AssignMeaipPro(){
   var frm = document.frmArr;
   var pro = document.frmStockDt.assignPro.value;
   if (!chkExesits(frm.cksel)){
        alert('선택 내역이 없습니다.');
        return;
   }

   if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                frm.shopsuplycash[i].value = frm.Orgshopsuplycash[i].value*1*pro/100*1;
                frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
   }else{
        frm.cksel.checked=comp.checked;
        frm.shopsuplycash.value = frm.Orgshopsuplycash.value*1*pro/100*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
   }

   summaryTotal(frm);
}

function AssignMeaipProbySell(){
   var frm = document.frmArr;
   var pro = document.frmStockDt.assignProSell.value;
   if (!chkExesits(frm.cksel)){
        alert('선택 내역이 없습니다.');
        return;
   }

   if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                frm.shopsuplycash[i].value = frm.OrgshopSellcash[i].value*1*pro/100*1;
                frm.SUBTTLshopsuplycash[i].value = frm.AssignrealcheckErr[i].value*1*frm.shopsuplycash[i].value*1;
            }
        }
   }else{
        frm.cksel.checked=comp.checked;
        frm.shopsuplycash.value = frm.OrgshopSellcash.value*1*pro/100*1;
        frm.SUBTTLshopsuplycash.value = frm.AssignrealcheckErr.value*1*frm.shopsuplycash.value*1;
   }

   summaryTotal(frm);
}

function fnNoDispIpChul(comp){
    var e = document.getElementsByName("NII");

    for (i=0;i<e.length;i++){
        if (comp.checked){
            e[i].style.display="none";
        }else{
            e[i].style.display="inline";
        }
    }
}

function chgComp(comp){
    if (comp.value=="L"){
        comp.form.LstYYYYMM.style.background="#FFFFFF";
        comp.form.LstYYYYMM.readOnly=false;

    }else{
        comp.form.LstYYYYMM.style.background="#CCCCCC";
        comp.form.LstYYYYMM.readOnly=true;
    }
}

function researchFrm(){
    if ((document.frm.CLDiv[1].checked)&&(document.frm.LstYYYYMM.value.length!=7)){
        alert('월말재고 년 월을 YYYY-MM 형식으로 입력하세요.');
        document.frm.LstYYYYMM.focus();
        return;
    }

    document.frm.submit();
}

function chkConfirm(comp){
    document.frmArr.ckAll.checked = comp.checked;
    var frm = document.frmArr;
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            frm.cksel[i].disabled=false;
            frm.cksel[i].checked=comp.checked;
        }
    }
}

function refreshStockArr(){
    if (confirm('관리자 메뉴 선택 상품 재고를 새로고침 하시겠습니까?')){
        //
        document.frmArr.mode.value="stockupArr";
        document.frmArr.submit();
    }
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="shopid" value="<%= shopid %>">
    <input type="hidden" name="cType" value="<%=cType%>">

	<tr height=30 align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" colspan="3">
		    매장 : <%= shopid %> &nbsp;&nbsp;
    		브랜드ID : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="20" maxlength="32"> &nbsp;&nbsp;

			<input type="radio" name="CLDiv" value="C" <%=CHKIIF(CLDiv="C","checked","") %> onClick="chgComp(this)">현재고 기준
    		<input type="radio" name="CLDiv" value="L" <%=CHKIIF(CLDiv="L","checked","") %> onClick="chgComp(this)">월말재고 기준

    		<input type="text" name="LstYYYYMM" value="<%= LstYYYYMM %>" size="7" maxlength="7" <%= CHKIIF(CLDiv="C","style='background=#CCCCCC'  readonly","") %> >
    		(YYYY-MM)
    		&nbsp;&nbsp;&nbsp;
    		<input type="checkbox" name="errExist" <%=CHKIIF(errExist="on","checked","") %> >재고<!--(시스템 or 실사)/-->오차/샘플 존재상품

    		&nbsp;&nbsp;&nbsp;
    		입출코드 : <input type="text" class="text" name="ipchulcode" size="10" value="<%=ipchulcode%>">
		</td>
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="researchFrm()">
		</td>
	</tr>
	<tr height=30 align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% if (CLDiv="C") then %>
		    재고 기준 : <b>현재</b>
		    <% else %>
		    재고 기준 : <b><%=LstYYYYMM %> 월말재고</b>
		    <% end if %>

		    <% if (C_ADMIN_AUTH) or (C_OFF_AUTH) or (C_MngPart) then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    		<input type="checkbox" name="ck1" onclick="chkConfirm(this);"> 	<input type="button" value="재고새로고침" onClick="refreshStockArr()">
    		<% end if %>
		</td>
		<td rowspan="3" width="300" >
		<%
		    if IsArray(ArrList1) then
		        cnt = UBound(ArrList1,2)+1
		    else
		        cnt = 0
		    end if

		%>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<% for i=0 to cnt-1 %>
		<tr>
		    <td><%=ArrList1(0,i) %></td>
		    <td><%=ArrList1(1,i) %></td>
		    <td><%=ArrList1(2,i) %></td>
		    <td><%=ArrList1(3,i) %></td>
		    <td><%=ArrList1(4,i) %></td>
		</tr>
		<% next %>
		</table>
        </td>

        <td rowspan="3" width="300" >
		<%
		    if IsArray(ArrList2) then
		        cnt = UBound(ArrList2,2)+1
		    else
		        cnt = 0
		    end if
		%>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<% for i=0 to cnt-1 %>
		<tr>
		    <td><%=ArrList2(0,i) %></td>
		    <td><%=ArrList2(1,i) %></td>
		    <td><%=ArrList2(2,i) %></td>
		    <td><%=ArrList2(3,i) %></td>
		</tr>
		<% next %>
		</table>
        </td>
	</tr>
	<tr height=40 align="center" bgcolor="#FFFFFF" >
		<td align="left">
			작업구분 : <b><%= iclearTypeName %></b> <br>
			<% if (cType="L") then %>
			1. 시스템 재고와 실사 재고의 차이 (즉, 오차)를 0으로 만듦<br>
			2. 출고처(<b>shopitemloss</b>) 에 출고등록함.<br>
			3. 업체 정산에 반영
			<% elseif (cType="S") then %>
			1. 실사 재고와 유효 재고의 차이 (즉, 샘플)를 0으로 만듦<br>
			2. 출고처(<b>shopitemsample</b>) 에 출고등록함.<br>
			3. 업체 정산에 반영
			<% else %>
			1. 시스템 재고와 실사 재고의 차이 (즉 오차)를 0으로 만듦 (기존 총 오차*-1 입력)<br>
			2. 기존 총 오차를 출고 처리 하여 (출고처 : <b>shopstockmodify</b>) 시스템 재고와 실사재고 차이를 없앰<br>
			3. 업체 정산에 반영 안되며, 재고 수량만 변경됨.
			<% end if %>
		</td>
	</tr>
	<!--
	<tr height=40 align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<input type="checkbox" name="noDispIpChul" onClick="fnNoDispIpChul(this)"> 입출고/판매 없는내역 표시 안함.
	    </td>
	</tr>
	-->
	</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
    <form name="frmStockDt">
	<tr height="30">
		<td align="left">
		    반영단가 비율 본사 매입가의 <input type="text" class="text" name="assignPro" value="100" size="3">%
			<input type="button" class="button" value="반영" onClick="AssignMeaipPro();">
			반영단가 비율 본사 판매가의 <input type="text" class="text" name="assignProSell" value="100" size="3">%
			<input type="button" class="button" value="반영" onClick="AssignMeaipProbySell();">

		</td>
		<td align="right">
		    <%=iclearTypeName%> 반영일
		    <input type="text" class="text" name="stockdate" value="<%= mayStockDate %>" size=11 readonly ><a href="javascript:calendarOpen(document.frmStockDt.stockdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		    (재고자산과 관련있음)
		    &nbsp;&nbsp;
		    로스타입
		    <select name="losstype" class="select">
				<option value="">선택</option>
				<% if (cType="S") then %>
				<option value="S">샘플폐기(정산미반영)</option>
				<% else %>
				<option value="M">로스처리(정산미반영)</option>
				<option value="L">로스처리(정산반영)</option>
				<% end if %>
		    </select>
			<input type="button" class="button" name="stock_sheet_print" value="선택상품 <%=iclearTypeName%> 출고 반영" onclick="AssignErrLoss();">
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->

<p>

	* 최대 <font color="red">500건</font>까지만 표시됩니다.

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmArr" method="post" action="OutItemListByBrand_Process.asp">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="lossDate" value="">
<input type="hidden" name="cType" value="<%=cType%>">
<input type="hidden" name="CLDiv" value="<%=CLDiv%>">
<input type="hidden" name="params" value="<%=params%>">
<input type="hidden" name="losstype" value="">
<input type="hidden" name="mode" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="ckAll" onClick="chkALL(this);"></td>
	<td width="80">상품코드</td>
	<td width="200">상품명 <font color="blue">[옵션명]</font></td>
	<td width="70">월<br>정산구분</td>
	<td width="70">현 판매가</td>
	<!-- <td width="70">매장 매입가</td> -->
	<td width="70">본사 매입가</td>
	<td width="60">물류입고</td>
	<td width="60">업체입고</td>
	<td width="60">총판매</td>
	<td width="60">정산수량</td>
	<td width="50">시스템<br>재고</td>
	<td width="50">오차</td>
	<td width="50">실사<br>재고</td>
	<td width="50">샘플</td>
	<td width="50">유효<br>재고</td>
	<td width="100"><%=iclearTypeName%><br>반영수</td>
	<td width="100"><%=iclearTypeName%><br>반영단가</td>
	<td width="40">남는<br>오차</td>
	<td >합계</td>
</tr>
<% for i=0 to oOffStock.FResultcount -1 %>
<%
TotErrrealcheckno = TotErrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno
TotSampleNo = TotSampleNo + oOffStock.FItemList(i).Ferrsampleitemno

if (cType="S") then
	currItemNo = oOffStock.FItemList(i).Ferrsampleitemno
else
	currItemNo = oOffStock.FItemList(i).Ferrrealcheckno
end if

%>
<tr bgcolor="#FFFFFF" id="<%=CHKIIF(oOffStock.FItemList(i).isIpChulNotExists,"NII","") %>" name="<%=CHKIIF(oOffStock.FItemList(i).isIpChulNotExists,"NII","") %>">
    <td>
		<input type="checkbox" name="cksel" value="<%= i %>" <%= CHKIIF(oOffStock.FItemList(i).IsCheckAvail,"","disabled") %> onClick="reCalcuLoss(this,<%= i %>);AnCheckClick(this);">
    	<input type="hidden" name="itemgubun" value="<%= oOffStock.FItemList(i).Fitemgubun %>">
    	<input type="hidden" name="itemid" value="<%= oOffStock.FItemList(i).Fitemid %>">
    	<input type="hidden" name="itemoption" value="<%= oOffStock.FItemList(i).Fitemoption %>">
    	<input type="hidden" name="shopitemprice" value="<%= oOffStock.FItemList(i).Fshopitemprice %>">
    	<input type="hidden" name="shopbuyprice" value="<%= oOffStock.FItemList(i).Fshopsuplycash %>"> <!-- 매입가로 넣음 -->
    	<input type="hidden" name="OrgshopSellcash" value="<%= oOffStock.FItemList(i).Fshopitemprice %>">
    </td>
    <td><a href="javascript:popShopCurrentStock('<%= shopid %>','<%= oOffStock.FItemList(i).getBarcode %>');"><%= oOffStock.FItemList(i).getBarcode %></a></td>
    <td><%= oOffStock.FItemList(i).Fshopitemname %>
    <% if oOffStock.FItemList(i).Fshopitemoptionname<>"" then %>
        <font color="blue">[<%= oOffStock.FItemList(i).Fshopitemoptionname %>]</font>
    <% end if %>
    </td>
    <td></td>
    <td align="right">
    <%= FormatNumber(oOffStock.FItemList(i).Fshopitemprice,0) %></td>
    <!-- <td align="right"><%= FormatNumber(oOffStock.FItemList(i).Fshopbuyprice,0) %></td> -->
    <td align="right"><%= FormatNumber(oOffStock.FItemList(i).Fshopsuplycash,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Flogicsipgono+oOffStock.FItemList(i).Flogicsreipgono,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Fbrandipgono+oOffStock.FItemList(i).Fbrandreipgono,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).FttlSellno,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).FjungsanCNT,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Fsysstockno,0) %></td>
	<td align="center"><%= FormatNumber(oOffStock.FItemList(i).Ferrrealcheckno,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Frealstockno,0) %></td>
    <td align="center"><%= FormatNumber(oOffStock.FItemList(i).Ferrsampleitemno,0) %></td>
	<td align="center"><%= FormatNumber((oOffStock.FItemList(i).Frealstockno + oOffStock.FItemList(i).Ferrsampleitemno),0) %></td>
    <td align="center">
		<input type="hidden" name="realcheckErr" value="<%= currItemNo %>">
		<input type="text" name="AssignrealcheckErr" value="<%= currItemNo*-1 %>" class="text" size="5"  style="text-align=center" onKeyUp="reCalcuLoss(this,<%= i %>)">
	</td>
	<td align="center">
		<input type="hidden" name="Orgshopsuplycash" value="<%= oOffStock.FItemList(i).fshopsuplycash %>">
		<input type="text" name="shopsuplycash" value="<%= oOffStock.FItemList(i).fshopsuplycash %>" class="text" size="9"  style="text-align=right" onKeyUp="reCalcuLoss(this,<%= i %>)"> <!-- 정산반영액 shopbuyprice 필드에 넣음 -->
	</td>
	<td align="center"><input type="text" name="SUBTTLrealcheckErrRemain" value="0" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
	<td align="center"><input type="text" name="SUBTTLshopsuplycash" value="<%= FormatNumber(oOffStock.FItemList(i).fshopsuplycash*currItemNo*-1,0) %>" class="text" size="9"  style="text-align=right;border=0" READONLY ></td>
</tr>
<% next %>
<tr bgcolor="#DDFFFF">
    <td>합계</td>
    <td align="center"><%=oOffStock.FResultcount%> 건</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="center"><%= FormatNumber(TotErrrealcheckno,0) %></td>
    <td></td>
	<td align="center"><%= FormatNumber(TotSampleNo,0) %></td>
	<td></td>
    <td align="center"><input type="text" name="TTLrealcheckErr" value="" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
    <td></td>
    <td align="center"><input type="text" name="TTLrealcheckErrRemain" value="" class="text" size="5"  style="text-align=center;border=0" READONLY ></td>
    <td align="center"><input type="text" name="TTLshopsuplycash" value="" class="text" size="9"  style="text-align=right;border=0" READONLY ></td>
</tr>
</form>
</table>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
