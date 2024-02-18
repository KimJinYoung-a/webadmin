<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품수정
' History : 서동석 생성
'			2018.06.02 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitem, deliverfixday, mwdiv, deliverytype, purchaseType, deliverarea, purchaseTypedefalut
dim makerid
dim saleCode, saleName
dim chkMWAuth 'mw 변경가능한 권한인지 체크

itemid = request("itemid")
makerid = request("makerid")
menupos = request("menupos")
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

if oitem.FTotalCount>0 then
	purchaseTypedefalut = oitem.FOneItem.fpurchaseType		' 구매유형
	'purchaseType = oitem.FOneItem.fpurchaseType		' 구매유형

	' 구매유형이 해외직구 일경우 강제 고정
	if purchaseType="9" then
		deliverfixday = "G"	' 해외직구
		mwdiv = "U"
		deliverarea = ""

		' 업체(무료)배송 일경우
		if oitem.FOneItem.Fdeliverytype="2" then
			deliverytype = oitem.FOneItem.Fdeliverytype
		else
			deliverytype = "9"
		end if
	else
		deliverfixday = oitem.FOneItem.Fdeliverfixday	' 해외직구
		mwdiv = oitem.FOneItem.Fmwdiv
		deliverarea = oitem.FOneItem.Fdeliverarea
		deliverytype = oitem.FOneItem.Fdeliverytype
	end if
end if

'==============================================================================
''업체 기본계약 구분
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
dim sqlStr
sqlStr = "select defaultmargine, maeipdiv as defaultmaeipdiv, "
sqlStr = sqlStr + " IsNULL(defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit,"
sqlStr = sqlStr + " IsNULL(defaultDeliverPay,0) as defaultDeliverPay,"
sqlStr = sqlStr + " IsNULL(defaultDeliveryType,'') as defaultDeliveryType"
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c where userid='" & oitem.FOneItem.Fmakerid & "'"
rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        defaultmargin           = rsget("defaultmargine")
        defaultmaeipdiv         = rsget("defaultmaeipdiv")
        defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
        defaultDeliverPay       = rsget("defaultDeliverPay")
        defaultDeliveryType     = rsget("defaultDeliveryType")
    end if
rsget.close

'==============================================================================
'세일마진
dim sailmargine, orgmargine, margine

''수정
if oitem.FOneItem.Fsailprice<>0 and oitem.FOneItem.Fsailsuplycash<>0 then
	sailmargine = Formatnumber((1-(CDbl(oitem.FOneItem.Fsailsuplycash)/CDbl(oitem.FOneItem.Fsailprice)))*100,0)
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 and oitem.FOneItem.Forgsuplycash<>0 then
	orgmargine = Formatnumber((1-(CDbl(oitem.FOneItem.Forgsuplycash)/CDbl(oitem.FOneItem.Forgprice)))*100,0)
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 and oitem.FOneItem.Fbuycash<>0 then
	margine = Formatnumber((1-(CDbl(oitem.FOneItem.Fbuycash)/CDbl(oitem.FOneItem.Fsellcash)))*100,0)
else
	margine = 0
end if

'==============================================================================
Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- 업체선택 --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c order by userid"
'   query1 = query1 + " where isusing='Y' order by userid desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "' "&tmp_str&">" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


'mw 변경가능 권한인지 체크
chkMWAuth = False
IF (Not oitem.FOneItem.FisCurrStockExists) or C_ADMIN_AUTH  THEN chkMWAuth = True ''

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 업체마진자동입력
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatinclude, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatinclude = frm.vatinclude[0].checked;

	if (frm.sailyn[0].checked == true) {
	    // 정상가격
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

    	if (imargin.length<1){
    		alert('마진을 입력하세요.');
    		frm.margin.focus();
    		return;
    	}

    	if (isellcash.length<1){
    		alert('판매가를 입력하세요.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (!IsDouble(imargin)){
    		alert('마진은 숫자로 입력하세요.');
    		frm.margin.focus();
    		return;
    	}

    	if (!IsDigit(isellcash)){
    		alert('판매가는 숫자로 입력하세요.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);  //parseInt-> round로 변경
			imileage = parseInt(isellcash*0.005) ;
    	}else{
    		ibuycash = isellcash - Math.round(isellcash*imargin/100);  //parseInt-> round로 변경
			imileage = parseInt(isellcash*0.005) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // 세일가격
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;
		isellcash = frm.sellcash.value;

    	if (isailmargin.length<1){
    		alert('세일마진을 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('세일판매가를 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('세일마진은 숫자로 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('세일판매가는 숫자로 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatinclude==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);   //parseInt-> round로 변경
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
			if (parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10>=40){
				imileage = parseInt(0) ;
			}
			else{
				imileage = parseInt(isailprice*0.005) ;
			}
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - Math.round(isailprice*isailmargin/100);   //parseInt-> round로 변경
    		isailsuplycashvat = 0;
			if (parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10>=40){
				imileage = parseInt(0) ;
			}
			else{
				imileage = parseInt(isailprice*0.005) ;
			}
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//할인율 계산
	if (frm.sailyn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "할인율: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
	}
}

// ============================================================================
// 저장하기
function fnSubmitSave() {
	if (document.itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		document.itemreg.designer.focus();
		return;
	}

    if (validate(document.itemreg)==false) {
        return;
    }

    if (document.itemreg.sailyn[0].checked == true) {
        // 정상가격
        if (Math.round((document.itemreg.sellcash.value*1) * (document.itemreg.margin.value*1) / 100) != ((document.itemreg.sellcash.value*1) - (document.itemreg.buycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[소비자가*마진 = 공급가]");
    		document.itemreg.sellcash.focus();

    		if (!confirm('마진율로 계산 할 수 없을때 공급가만 입력하면 마진율은 공급가에 맞춰 계산됩니다. \n계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (document.itemreg.mileage.value*1 > document.itemreg.sellcash.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            document.itemreg.mileage.focus();
            return;
        }

        <% if oitem.FOneItem.Fitemdiv<>"09" then %>
        if (document.itemreg.sellcash.value*1 < 0 || document.itemreg.sellcash.value*1 >= 20000000){
			alert("판매 가격은 20,000,000만원 미만으로 등록 가능합니다.");
			document.itemreg.sellcash.focus();
			return;
		}
		<% end if %>

    } else {
        // 할인가격
        if (Math.round((document.itemreg.sailprice.value*1) * (document.itemreg.sailmargin.value*1) / 100) != ((document.itemreg.sailprice.value*1) - (document.itemreg.sailsuplycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[할인소비자가*할인마진 = 할인공급가]");
    		document.itemreg.sailprice.focus();

    		if (!confirm('계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (document.itemreg.mileage.value*1 > document.itemreg.sailprice.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            document.itemreg.mileage.focus();
            return;
        }

        <% if oitem.FOneItem.Fitemdiv<>"09" then %>
        if (document.itemreg.sailprice.value*1 < 0 || document.itemreg.sailprice.value*1 >= 20000000){
			alert("판매 가격은 20,000,000만원 미만으로 등록 가능합니다.");
			document.itemreg.sailprice.focus();
			return;
		}
		<% end if %>
    }


    //세일가격이 정상가격 보다 클 수 없음.
    if (document.itemreg.sailprice.value*1>document.itemreg.sellcash.value*1){
        alert('세일가격이 정상가보다 클 수 없습니다.');
        return;
    }

    if (document.itemreg.sailsuplycash.value*1>document.itemreg.buycash.value*1){
        alert('세일매입가가 정상 매입가보다 클 수 없습니다.');
        return;
    }

	// 원래입력된 판매가보다 수정된 판매가의 차이가 많이 날때 확인 메시지
	if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.2)%>) {
		if(!confirm("\n\n\n\n입력하신 소비자가 수정하기 전의 가격보다 매우 많이 차이납니다(80%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sellcash.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.4)%>) {
		if(!confirm("\n\n\n\n입력하신 소비자가 수정하기 전의 가격보다 매우 많이 차이납니다(60%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sellcash.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.6)%>) {
		if(!confirm("\n\n\n\n입력하신 소비자가 수정하기 전의 가격보다 매우 많이 차이납니다(40%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sellcash.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sellcash.value<<%=fix(oitem.FOneItem.Fsellcash*0.8)%>) {
		if(!confirm("\n\n\n\n입력하신 소비자가 수정하기 전의 가격보다 매우 많이 차이납니다(20%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsellcash,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sellcash.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	}

	<% if oitem.FOneItem.Fsailyn="Y" then %>
	// 원래입력된 할인가보다 수정된 할인가의 차이가 많이 날때 확인 메시지
	if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.2)%>) {
		if(!confirm("\n\n\n\n입력하신 할인가가 수정하기 전의 가격보다 매우 많이 차이납니다(80%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sailprice.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.4)%>) {
		if(!confirm("\n\n\n\n입력하신 할인가가 수정하기 전의 가격보다 매우 많이 차이납니다(60%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sailprice.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.6)%>) {
		if(!confirm("\n\n\n\n입력하신 할인가가 수정하기 전의 가격보다 매우 많이 차이납니다(40%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sailprice.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	} else if(document.itemreg.sailprice.value<<%=fix(oitem.FOneItem.Fsailprice*0.8)%>) {
		if(!confirm("\n\n\n\n입력하신 할인가가 수정하기 전의 가격보다 매우 많이 차이납니다(20%이상).\n\n수정전 가격 [ <%=formatNumber(oitem.FOneItem.Fsailprice,0)%> ]원 → 입력하신 가격 [ "+plusComma(document.itemreg.sailprice.value)+" ]원\n\n\n입력하신 내용이 정확합니까?\n\n\n\n")) {
			return;
		}
	}
	<% end if %>

	// 할인율 검사(50%이상 경고)
	if (document.itemreg.sailyn[1].checked == true) {
		if(((document.itemreg.sellcash.value-document.itemreg.sailprice.value)/document.itemreg.sellcash.value*100)>50) {
			if(!confirm("\n\n할인율이 매우 높게 설정되어있습니다.\n\n입력하신 내용이 정확합니까?")) {
				return;
			}
		}
	}

	// 업체 할인분담율 체크 (분담율 50%이상 설정 불가)
	if (document.itemreg.sailyn[1].checked == true && document.itemreg.mwdiv.value!="M") {
		var limitMarPrc = document.itemreg.orgsuplycash.value-((document.itemreg.orgprice.value-document.itemreg.sailprice.value)*0.5);
		var limitMarPer = (document.itemreg.sailprice.value-limitMarPrc)/document.itemreg.sailprice.value*100;
		if(parseInt(limitMarPrc)>parseInt(document.itemreg.sailsuplycash.value)) {
			if(!confirm('업체 할인 분담율이 50%를 넘습니다. (최대할인마진 : '+limitMarPer+'%)\n\n입력하신 내용이 정확합니까?')){;
				return;
			}
		}
	}

    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((document.itemreg.defaultFreeBeasongLimit.value*1>0) && (document.itemreg.defaultDeliverPay.value*1>0))||(document.itemreg.defaultDeliveryType.value=="9") )){
        if (document.itemreg.deliverytype[3].checked){
            alert('배송 구분을 확인해주세요. 개별배송 업체가 아닙니다.');
            return;
        }
    }

//    //업체착불배송 : 조건배송도 착불설정가능 - 삭제 2015.05.22
//    if (!(document.itemreg.defaultDeliveryType.value=="7")||(document.itemreg.defaultDeliveryType.value=="9"))&&(document.itemreg.deliverytype[4].checked)){
//        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
//        document.itemreg.deliverytype[4].focus();
//        return;
//    }

    if ((document.itemreg.deliverytype[1].checked)||(document.itemreg.deliverytype[3].checked)||(document.itemreg.deliverytype[4].checked)){
    	if(document.itemreg.mwdiv.length>0){
	        if ((document.itemreg.mwdiv[0].checked)||(document.itemreg.mwdiv[1].checked)){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
    	}else{
    		if ((document.itemreg.mwdiv.value=="M")||(document.itemreg.mwdiv.value=="W")){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
    	}
     //   if (document.itemreg.deliverOverseas.checked){
     //       alert('텐바이텐 배송일 경우에만 해외배송을 하실 수 있습니다.');
     //       return;
    //    }
    }
    if(document.itemreg.mwdiv.length>0){
	    if (document.itemreg.mwdiv[2].checked){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
	    }
	}else{
		 if (document.itemreg.mwdiv.value=="U"){
	        if ((document.itemreg.deliverytype[0].checked)||(document.itemreg.deliverytype[2].checked)){
	            alert('배송 구분을 확인해주세요. 매입 구분과 일치하지 않습니다..');
	            return;
	        }
	    }
	}
	if(document.itemreg.deliverfixday[1].checked) {
		if(document.itemreg.freight_min.value<=0||document.itemreg.freight_max.value<=0) {
            alert('화물배송 비용을 입력해주세요.');
            document.itemreg.freight_min.focus();
            return;
		}
	}

	// 배송방법 해외직구 체크
	<% if purchaseTypedefalut="9" then %>
		if (itemreg.deliverfixday[3].checked == false){
			alert('해외직구 브랜드 입니다. 해외직구로 선택해 주세요.')
			return;
		}
	<% end if %>
	if (itemreg.deliverfixday[3].checked == true){
		if (itemreg.mwdiv[2].checked == false){
			alert('해외직구는 업체배송만 선택 가능 합니다.');
			return;
		}
		if ( !(itemreg.deliverytype[1].checked == true || itemreg.deliverytype[3].checked == true) ){
			alert('해외직구는 업체무료배송과 업체조건배송만 선택 가능 합니다.');
			return;
		}
		if (itemreg.deliverarea[0].checked == false){
			alert('해외직구는 전국배송만 선택 가능 합니다.');
			return;
		}
	}

	if(document.itemreg.orderMinNum.value<1||document.itemreg.orderMinNum.value>32000) {
        alert('최소판매수는 1~32,000 범위의 숫자로 입력해주세요.');
        document.itemreg.orderMinNum.focus();
        return;
	}
	if(document.itemreg.orderMaxNum.value<1||document.itemreg.orderMaxNum.value>32000) {
        alert('최대판매수는 1~32,000 범위의 숫자로 입력해주세요.');
        document.itemreg.orderMaxNum.focus();
        return;
	}
	if(parseInt(document.itemreg.orderMinNum.value)>parseInt(document.itemreg.orderMaxNum.value)) {
        alert('최대판매수보다 최소판매수가 클 수 없습니다.');
        document.itemreg.orderMinNum.focus();
        return;
	}

	if((document.itemreg.sellyn[0].checked||document.itemreg.sellyn[1].checked)&&(document.itemreg.isusing[1].checked)) {
        alert('판매여부와 사용여부를 확인해주세요.\n\n※사용하지 않는 상품은 판매중을 선택할 수 없습니다.');
        return;
	}

	// 선착순 또는 Just1Day 선택이면 저장 전 확인
	if(document.itemreg.availPayType[0].checked||document.itemreg.availPayType[1].checked) {
		if(!confirm("선착순 결제 상품을 선택하셨습니다.\n그대로 진행하시겠습니까?")){	
			return;
		}
	}

    //==================================================================================

    if(confirm("상품을 올리시겠습니까?") == true){
        document.itemreg.deliverytype[0].disabled=false;
		document.itemreg.deliverytype[1].disabled=false;
		document.itemreg.deliverytype[2].disabled=false;
        document.itemreg.deliverytype[3].disabled=false;
        document.itemreg.deliverytype[4].disabled=false;
        document.itemreg.submit();
    }

}

function SubmitSave() {
	//텐배 옵션 추가금 체크 (정태훈 2020-01-29)
	var deliverOverseas="", mwdiv="";
	if(document.itemreg.deliverOverseas.checked){
		deliverOverseas="Y";
	}
	if(document.itemreg.mwdiv.length>0){
		mwdiv = $("[name=mwdiv]:checked").val();
	} else {
		mwdiv = $("[name=mwdiv]").val();
	}

    $.ajax({
        type: "POST",
        url: "/admin/itemmaster/ajaxItemOptionPriceCheck.asp",
        data: "itemid=<%=itemid%>&mwdiv="+mwdiv+"&deliverOverseas="+deliverOverseas,
        cache: false,
        success: function(message){
			if(message=="1"){
				alert('텐바이텐 배송의 경우 옵션 추가금액을 사용할 수 없습니다.');
				return;
			}
			else if(message=="2"){
				alert('해외배송을 하는 경우 옵션 추가금액을 사용할 수 없습니다.');
				return;
			}
			else{
				fnSubmitSave();
			}
        },
        error: function(err) {
           	alert('배송 구분을 확인해주세요.');
			return;
        }
    });
}

function TnGoClear(frm){
	frm.buycash.value = "";
	frm.mileage.value = "";
}

// 배송방법
function TnCheckFixday(frm) {
	if(frm.deliverfixday[0].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="none";
	} else if(frm.deliverfixday[1].checked) {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=true;
		frm.deliverarea[2].disabled=true;
		document.getElementById("lyrFreightRng").style.display="";

	// 해외직구
	} else if(frm.deliverfixday[3].checked) {
		frm.mwdiv[2].checked=true;
		frm.deliverarea[0].checked=true;

		document.getElementById("lyrFreightRng").style.display="none";
	} else {
		frm.deliverarea[0].checked=true;
		frm.deliverarea[1].disabled=false;
		frm.deliverarea[2].disabled=false;
		document.getElementById("lyrFreightRng").style.display="none";
	}
}

// 배송구분
function TnCheckUpcheDeliverYN(frm){
	if (frm.deliverytype[0].checked || frm.deliverytype[2].checked){
		if(frm.mwdiv.length>0){
			if (frm.mwdiv[2].checked){
				alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
				frm.mwdiv[0].checked=true;
			}
		}else{
			if (frm.mwdiv.value=="U"){
				alert("매입위탁 구분이 업체일 경우\n배송구분을 텐바이텐 배송으로 선택 하실 수 없습니다!!\n매입위탁구분을 확인해주세요!!");
				frm.mwdiv.value="M";
			}
		}
	}
	else if(frm.deliverytype[1].checked || frm.deliverytype[3].checked){
	//else if(frm.deliverytype[1].checked ){
		if(frm.mwdiv.length>0){
			if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
				alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
				frm.mwdiv[2].checked=true;
			}
		}else{
			if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
				alert("매입위탁 구분이 매입이나 위탁일 경우\n배송구분을  업체배송으로 선택 하실 수 없습니다!!!\n매입위탁구분을 확인해주세요!!");
				frm.mwdiv.value="U";
			}
		}
	}
}

function TnChkIsUsing(frm) {
	if(frm.isusing[0].checked) {
		frm.sellyn[0].disabled=false;
		frm.sellyn[1].disabled=false;
	} else {
		if(frm.sellyn[0].checked||frm.sellyn[1].checked) {
			alert("사용여부를 사용안함으로 선택하셨습니다.\n판매여부가 [판매안함]으로 자동설정됩니다.");
		}
		frm.sellyn[2].checked=true;
		frm.sellyn[0].disabled=true;
		frm.sellyn[1].disabled=true;
	}
}

function TnCheckSailYN(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

// 매입위탁구분
function TnCheckUpcheYN(frm){
if(frm.mwdiv.length>0){
	if (frm.mwdiv[0].checked || frm.mwdiv[1].checked){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
		frm.deliverytype[3].disabled=true;  //업체개별배송(9)
		frm.deliverytype[4].disabled=true;  //업체착불배송(7)
		//frm.deliverOverseas.checked=true;	// 해외배송체크 -> To 넣은사람. 이거때문에 체크빼고 수정했는데 계속 체크가 되서 주석처리했음.
	}
	else if(frm.mwdiv[2].checked){

	    // 배송구분 지정(업체배송)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// 기본 체크
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// 업체착불배송 기본 체크
	    }else{
	        frm.deliverytype[1].checked=true;	// 기본 체크
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;

        <%
        ' 해외직구 일경우
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //업체착불배송(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //업체착불배송(7)
		<% end if %>

       // frm.deliverOverseas.checked=false;	// 해외배송체크해제
	}
}else{
	if (frm.mwdiv.value=="M" || frm.mwdiv.value=="W"){
		frm.deliverytype[0].checked=true;	// 기본체크
		// 배송구분 지정(텐바이텐)
		frm.deliverytype[0].disabled=false;
		frm.deliverytype[1].disabled=true;
		frm.deliverytype[2].disabled=false;
		frm.deliverytype[3].disabled=true;  //업체개별배송(9)
		frm.deliverytype[4].disabled=true;  //업체착불배송(7)
		//frm.deliverOverseas.checked=true;	// 해외배송체크
	}
	else if(frm.mwdiv.value=="U"){
	    // 배송구분 지정(업체배송)
	    if ((frm.defaultFreeBeasongLimit.value*1>0)&&(frm.defaultDeliverPay.value*1>0)){
	        frm.deliverytype[3].checked=true;	// 기본 체크
	    }else if(frm.defaultDeliveryType.value=="7"){
	        frm.deliverytype[4].checked=true;	// 업체착불배송 기본 체크
	    }else{
	        frm.deliverytype[1].checked=true;	// 기본 체크
	    }

		frm.deliverytype[0].disabled=true;
		frm.deliverytype[1].disabled=false;
		frm.deliverytype[2].disabled=true;
        frm.deliverytype[3].disabled=false;

         <%
        ' 해외직구 일경우
        if deliverfixday="G" then
        %>
        	frm.deliverytype[4].disabled=true;  //업체착불배송(7)
        <% else %>
			frm.deliverytype[4].disabled=false;  //업체착불배송(7)
		<% end if %>

        frm.deliverOverseas.checked=false;	// 해외배송체크해제
	}
}

	if (frm.deliverytype[1].checked==true || frm.deliverytype[3].checked==true){
		frm.deliverfixday[3].disabled=false;	// 해외직구
	}
}

function CheckSailEnDisabled(frm){
	if (frm.sailyn[0].checked == true) {
	    // 정상가격
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.buycash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailsuplycash.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // 세일가격
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.buycash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailsuplycash.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>상품 가격/판매 정보 수정</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<br><b>등록된 상품의 가격 및 판매 정보를 수정합니다.</b>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><br>기본정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="ItemPriceInfo">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">

<!-- 업체 기본 계약 구분 -->
<input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">

<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<tr align="left">
<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <%= oitem.FOneItem.Fitemid %>
	  &nbsp;&nbsp;&nbsp;&nbsp;
	  <input type="button" value="미리보기" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체ID :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%=oitem.FOneItem.FMakerid %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fitemname %></td>
</tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><br>가격정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">가격설정 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center">
			<td height="25" width="90" bgcolor="#DDDDFF">선택</td>
			<td width="100" bgcolor="#DDDDFF">소비자가</td>
			<td width="100" bgcolor="#DDDDFF">공급가</td>
			<td width="100" bgcolor="#DDDDFF">마진</td>
			<td bgcolor="#DDDDFF">&nbsp;</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(document.itemreg)" value="N" <% if oitem.FOneItem.Fsailyn = "N" then response.write "checked" %>> 정상가격</label></td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(document.itemreg);">원
			<% else %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(document.itemreg);">원
			<% end if %>
			</td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">원
			<% else %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">원
			<% end if %>
			</td>
			<% if oitem.FOneItem.Fsailyn = "N" then %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][마진]" value="<%= margine %>">%
			</td>
			<% else %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][마진]" value="<%= orgmargine %>">%
			</td>
			<% end if %>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="공급가 자동계산" class="button" onclick="CalcuAuto(document.itemreg);">
			</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(document.itemreg)" value="Y" <% if oitem.FOneItem.Fsailyn = "Y" then response.write "checked" %>> 할인가격</label></td>
			<input type="hidden" name="sailpricevat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailprice" maxlength="16" size="8" class="text" id="[on,on,off,off][할인소비자가]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(document.itemreg);">원
			</td>
			<input type="hidden" name="sailsuplycashvat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailsuplycash" maxlength="16" size="8" class="text" id="[on,on,off,off][할인공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">원
			</td>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailmargin" maxlength="32" size="5" class="text" id="[on,off,off,off][할인마진]" value="<%= sailmargine %>">%
			</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="공급가 자동계산" class="button" onclick="CalcuAuto(document.itemreg);">
				<%
					dim itemSalePer : itemSalePer=0
					if oitem.FOneItem.Fsailyn="Y" then
						itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
						itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
					end if
				%>
				<span id="lyrPct"><% if itemSalePer>0 then %>할인율: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
			</td>
		</tr>
		<%
			'// 할인코드 접수
			Call oitem.FOneItem.getSeleCode(saleCode, saleName)
			if Not(saleCode="" or isNull(saleCode)) then
		%>
		<tr height="25">
			<td bgcolor="#F8F8FA" align="center">해당할인정보</td>
			<td colspan="4" bgcolor="#F8F8FA"><a href="/admin/shopmaster/sale/saleReg.asp?sC=<%=saleCode%>&menupos=290" target="blank">[<b><%=saleCode%></b>] <%=saleName%></a></td>
		</tr>
		<% end if %>
		</table>
		<br>
		- 공급가는 <b>부가세 포함가</b>입니다.<br>
		- 소비자가(할인가)와 마진(할인마진)을 입력하고 [공급가자동계산] 버튼을 누르면 공급가와 마일리지가 자동계산됩니다.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">마일리지 :</td>
	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" class="text" id="[on,on,off,off][마일리지]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" onclick="TnGoClear(this.form);" <% if oitem.FOneItem.Fvatinclude = "Y" then response.write "checked" %>>과세</label>
		<label><input type="radio" name="vatinclude" value="N" onclick="TnGoClear(this.form);" <% if oitem.FOneItem.Fvatinclude = "N" then response.write "checked" %>>면세</label>
	</td>
</tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left"><br>판매정보</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">매입위탁구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% IF chkMWAuth THEN %>
		<label><input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "M" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >매입</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "W" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >위탁</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <% if mwdiv = "U" then response.write "checked" %>>업체배송</label>
		&nbsp;&nbsp; - 매입위탁구분에 따라 배송구분이 달라집니다. 배송구분을 확인해주세요.
		<%ELSE%>
		<%= fnColor(mwdiv,"mw") %>
		<input type="hidden" name="mwdiv" value="<%=mwdiv%>">
		<%END IF%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "1" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >텐바이텐배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "2" then response.write "checked" %>>업체(무료)배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "4" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >텐바이텐무료배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "9" then response.write "checked" %>>업체조건배송(개별 배송비부과)</label>
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if deliverytype = "7" then response.write "checked" %> <%=chkIIF(deliverfixday="G" ," disabled","")%> >업체착불배송</label>
		<% if deliverytype = "6" then %>
		<label><input type="radio" name="deliverytype" value="6" onclick="TnCheckUpcheDeliverYN(this.form);" checked <%=chkIIF(deliverfixday="G" ," disabled","")%> ><font color="darkred">현장수령</font></label>
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송방법 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" <%=chkIIF(Trim(deliverfixday)="" or IsNull(deliverfixday),"checked","")%> <%=chkIIF(purchaseTypedefalut="9"," disabled","")%> onclick="TnCheckFixday(this.form)">택배(일반)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" <%=chkIIF(deliverfixday="X","checked","")%> <%=chkIIF(purchaseTypedefalut="9" ," disabled","")%> onclick="TnCheckFixday(this.form)">화물</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" <%=chkIIF(deliverfixday="C","checked","")%> <%=chkIIF(purchaseTypedefalut="9"," disabled","")%> onclick="TnCheckFixday(this.form)">플라워지정일</label>
		<label><input type="radio" name="deliverfixday" value="G" <%=chkIIF(deliverfixday="G","checked","")%> <%=chkIIF(mwdiv<>"U" or (deliverytype <> "2" and purchaseTypedefalut <> "9")," disabled","")%> onclick="TnCheckFixday(this.form)">해외직구</label>
		<label><input type="radio" name="deliverfixday" value="L" <%=chkIIF(deliverfixday="L","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv<>"08"," disabled","")%> onclick="TnCheckFixday(this.form)">클래스</label>
		<span id="lyrFreightRng" style="display:<%=chkIIF(deliverfixday="X","","none")%>;">
			<br />&nbsp;
			반품/교환 시 화물배송 비용(편도) :
			최소 <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">원 ~
			최대 <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">원
		</span>
		<br>&nbsp;<font color="red">(플라워 상품인 경우만 수도권배송, 서울배송, 플라워지정일 옵션이 사용가능합니다.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송지역 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(deliverarea)="" or IsNull(deliverarea),"checked","")%>>전국배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" <%=chkIIF(deliverarea="C","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >수도권배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" <%=chkIIF(deliverarea="S","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> >서울배송</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF(oitem.FOneItem.FdeliverOverseas="Y","checked","")%> <%=chkIIF(deliverfixday="G" ," disabled","")%> title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">포장가능여부 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fpojangok %> <!-- 읽기전용 포장 여부 수정은 다른곳에서 popup 으로. -->
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">재입고예정일 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="reipgodate" class="text" id="[off,off,off,off][재입고예정일]" size="10" value="<%= oitem.FOneItem.FreipgoDate %>" maxlength="10">
		<a href="javascript:calendarOpen(document.itemreg.reipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		<a href="javascript:ClearVal(document.itemreg.reipgodate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">최소/최대 판매수 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		최소
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최소판매수]" value="<%= oitem.FOneItem.ForderMinNum %>">
		/ 최대
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최대판매수]" value="<%= oitem.FOneItem.ForderMaxNum %>">
		(한 주문에 판매 제한 수)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>판매함</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>일시품절</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>판매안함<label>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">사용여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="Y","checked","")%>>사용함</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="N","checked","")%>>사용안함</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">선착순 결제 상품 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
	<label><input type="radio" name="availPayType" value="9" <%=chkIIF(oitem.FOneItem.FavailPayType="9","checked","")%>>선착순</label>
		<label><input type="radio" name="availPayType" value="8" <%=chkIIF(oitem.FOneItem.FavailPayType="8","checked","")%>>저스트원데이</label>
		<label><input type="radio" name="availPayType" value="0" <%=chkIIF(oitem.FOneItem.FavailPayType="0","checked","")%>>일반</label>
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
      <input type="button" value="저장하기" class="button" onClick="SubmitSave()">
      <input type="button" value="취소하기" class="button" onClick="self.close()">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->

<p>
<script language='javascript'>
// 매입위탁구분 및 배송구분세팅
TnCheckUpcheYN(document.itemreg);
for (var i = 0; i < document.itemreg.elements.length; i++) {
    if (document.itemreg.elements[i].name == "deliverytype") {
        if (document.itemreg.elements[i].value == "<%= deliverytype %>") {
            document.itemreg.elements[i].checked = true;
        }
    }
}

// 세일
CheckSailEnDisabled(document.itemreg);
</script>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->