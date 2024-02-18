<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%
'' divcd    divname
'' A008     주문취소
'' A020     전체취소 - 주문취소로 통합
'' A021     부분취소 - 주문취소로 통합

'' A004     반품접수
'' A006     출고시유의사항
'' A009     기타사항(Memo)
'' A010     회수신청  -  텐바이텐 배송만 가능?
'' A011     맞교환회수

'' A700     업체기타정산

dim i, id, mode, divcd, orderserial
dim ckAll
id          = RequestCheckvar(request("id"),10)
mode        = RequestCheckvar(request("mode"),16)
divcd       = RequestCheckvar(request("divcd"),4)
orderserial = RequestCheckvar(request("orderserial"),16)
ckAll       = RequestCheckvar(request("ckAll"),10)


dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


''정보가 없을경우 신규등록으로 간주함
if (ocsaslist.FResultCount<1) then
    set ocsaslist.FOneItem = new CCSASMasterItem
    ocsaslist.FOneItem.FId=0
    ocsaslist.FOneItem.Fdivcd = divcd
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    orderserial = ocsaslist.FOneItem.Forderserial
end if


''등록인지 수정인지 여부
dim IsRegState
IsRegState = (ocsaslist.FOneItem.FId=0)



''주문 마스타
dim oordermaster
set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' 과거 6개월 이전 내역 검색
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


dim ocsOrderDetail
set ocsOrderDetail = new CCSASList
ocsOrderDetail.FRectCsAsID = ocsaslist.FOneItem.FId
ocsOrderDetail.FRectOrderSerial = orderserial

if (oordermaster.FRectOldOrder = "on") then
    ocsOrderDetail.FRectOldOrder = "on"
end if

''접수 상태에서는 전체 주문목록 / 수정, 완료상태에서는 접수된 내역만 보여줌
if (IsRegState) then
    ocsOrderDetail.GetOrderDetailByCsDetail
else
    ocsOrderDetail.GetCsDetailList
end if


''환불정보
dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = ocsaslist.FOneItem.FId
orefund.GetOneRefundInfo

if (orefund.FOneItem.Fencmethod = "TBT") then
	''orefund.FOneItem.Frebankaccount = Decrypt(orefund.FOneItem.FencAccount)
elseif (orefund.FOneItem.Fencmethod = "PH1") then
	orefund.FOneItem.Frebankaccount = orefund.FOneItem.Fdecaccount
end if

if (ocsaslist.FOneItem.FId <> 0) and ((ocsaslist.FOneITem.FDeleteyn = "Y") or (mode = "finished")) then
	if DateDiff("m", ocsaslist.FOneItem.Fregdate, Now) > 3 then
		orefund.FOneItem.Frebankaccount = ""
		orefund.FOneItem.Frebankownername = ""
		orefund.FOneItem.Frebankname = ""
	end if
end if


''수정 가능 상태
dim IsEditState
IsEditState = (Not IsRegState) and ((mode="editreginfo") or (mode="editrefundinfo"))

''완료처리 가능
dim IsFinishProcState
IsFinishProcState = (Not IsRegState) and (mode="finishreginfo")

''완료상태인지
dim IsStateFinished
IsStateFinished = (ocsaslist.FOneItem.FCurrState="B007")

''업체처리완료상태인지
dim IsUpcheConfirmState
IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")

''detail's distinct id
dim distinctid

''접수 불가시 메세지
dim JupsuInValidMsg

if (Left(orderserial,1)<>"A") and (oordermaster.FResultCount<1) and (mode<>"editrefundinfo") then
    response.write "<br><br>!!! 과거 주문내역이거나 주문 내역이 없습니다. - 관리자 문의 요망"
    dbget.close()	:	response.End
end if

''접수 가능 여부 ''주문내역이 없을경우 체크.
dim IsJupsuProcessAvail

if (oordermaster.FResultCount>0) then
    IsJupsuProcessAvail = ocsaslist.FOneItem.IsAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
else
    IsJupsuProcessAvail = false
end if


'' 배송비, 배송옵션
dim baesongmethodstr,orgbeasongpay

'' 원주문 상품금액
dim orgitemcostsum

'' 접수상품 합계금액
dim regitemcostsum

dim isDefaultCheckedItem,isAllchecked

''기 접수된 CS건 있는지 확인
dim oOldcsaslist
set oOldcsaslist = New CCSASList
oOldcsaslist.FRectNotCsID     = id
oOldcsaslist.FRectOrderserial = orderserial
oOldcsaslist.GetCSASMasterList

dim ExistsRegedCSCount
ExistsRegedCSCount = oOldcsaslist.FResultCount


''취소 관련정보 Display여부
dim IsCancelInfoDisplay
IsCancelInfoDisplay = ((IsRegState) or (orefund.FResultCount>0))
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A000")       '' 맞교환
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A001")       '' 누락
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A002")       '' 서비스
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A009")       '' 기타메모
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A003")       '' 환불접수
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A005")       '' 외부몰환불접수
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A006")       '' 출고시 유의사항
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A700")       '' 업체 기타정산


''환불 관련  표시 :
dim IsReFundInfoDisplay
if (oordermaster.FResultCount>0) then
    IsReFundInfoDisplay = ocsaslist.FOneItem.IsRefundProcessRequire(oordermaster.FOneItem.Fipkumdiv,oordermaster.FOneItem.FCancelyn)
else
    IsReFundInfoDisplay = false
end if

IsReFundInfoDisplay = (IsReFundInfoDisplay and IsJupsuProcessAvail)
IsReFundInfoDisplay = IsReFundInfoDisplay or (divcd="A003") or (divcd="A005")
IsReFundInfoDisplay = IsReFundInfoDisplay or (orefund.FResultCount>0)

''기타정산 표시 :
dim IsUpCheAddJungsanDisplay
IsUpCheAddJungsanDisplay = (divcd="A004") or (divcd="A700") or (divcd="A000") ''반품접수, 업체 기타정산

''상품 목록 display 여부
dim IsItemDetailDisplay
IsItemDetailDisplay = True

if (divcd="A003") or (divcd="A005") then
    IsItemDetailDisplay = False
end if

%>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript'>
var IsCancelProcess = <%= LCase(ocsaslist.FOneItem.IsCancelProcess) %>;
var IsReturnProcess = <%= LCase(ocsaslist.FOneItem.IsReturnProcess) %>;
var IsRefundProcess = <%= LCase(ocsaslist.FOneItem.IsRefundProcess) %>;
var IsServiceDeliverProcess= <%= LCase(ocsaslist.FOneItem.IsServiceDeliverProcess) %>;

var CDEFAULTBEASONGPAY = 2000;
var Fdivcd = "<%= divcd %>";

<% if RequestCheckvar(request("finishtype"),10)<>"" then %>
function FinishActType(finishtype){
    if (finishtype=="1"){
        PopCSSMSSend('<%= oordermaster.FOneItem.Freqhp %>','<%= orderserial %>','<%= oordermaster.FOneItem.Fuserid %>','텐바이텐입니다. 고객님 환불이 완료되었습니다. 즐거운 하루 되세요 감사합니다.^^*')
    }
}

FinishActType('<%= RequestCheckvar(request("finishtype"),10) %>');
<% end if %>
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
<% if (IsFinishProcState) then %>
    alert('수정창에서 수정해주세요. - 완료처리시 수정불가');
    return;
<% end if %>
    var params = "?gubun01=" + value_gubun01 + "&gubun02=" + value_gubun02 + "&name_gubun01=" + name_gubun01 + "&name_gubun02=" + name_gubun02 + "&name_gubun01name=" + name_gubun01name + "&name_gubun02name=" + name_gubun02name +"&name_frm=" + name_frm + "&targetDiv=" + targetDiv;
    initializeURL("/cscenter/action/ajax_cs_gubun_select.asp" + params);
    initializeReturnFunction("processAjaxCSGubunSelect(" + targetDiv + ")");
    initializeErrorFunction("onErrorAjaxCSGubunSelect()");
    startRequest();
}

function processAjaxCSGubunSelect(targetDiv) {
    eval(targetDiv).innerHTML = xmlHttp.responseText;
}

function onErrorAjaxCSGubunSelect() {
    alert("데이터를 읽는 도중에 에러가 발생했습니다. 잠시후 다시 시도해보시기 바랍니다.[CODE:" + xmlHttp.status + "]");
}

function colseCausepop(targetDiv){
    eval(targetDiv).innerHTML = "";
}

function delGubun(name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
    eval("document." + name_frm + "." + name_gubun01).value = "";
    eval("document." + name_frm + "." + name_gubun02).value = "";
    eval("document." + name_frm + "." + name_gubun01name).value = "";
    eval("document." + name_frm + "." + name_gubun02name).value = "";

    eval(targetDiv).innerHTML = "";
}


function selectGubun(value_gubun01,value_gubun02,value_gubun01name,value_gubun02name,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
<% if  (IsFinishProcState) then %>
    alert('수정창에서 수정해주세요. - 완료처리시 수정불가');
    return;
<% end if %>
    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";

    //마스타에서 선택한 경우 전체 선택된 Detail에 세팅
    if (targetDiv=="causepop"){
        for (var i=0;i<frm.elements.length;i++){
            var e = frm.elements[i];

            if ((e.type=="checkbox")&&(e.checked)&&(e.name=="orderdetailidx")){
                setDetailCause(e.value,value_gubun01,value_gubun02,value_gubun01name,value_gubun02name, name_frm);
            }
        }
    }

    //완료시에는 금액 수정 불가
    //업체개별배송 체크
    CheckUpcheDeliverPay(frm);

    //배송비 체크
    CheckDeliverPay(frmaction);

    CalculateAndApplyItemCostSum(frmaction);

}

function setDetailCause(idx, value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_frm) {
        var ogubun01 = eval(name_frm + ".gubun01_" + idx);
        ogubun01.value = value_gubun01;

        var ogubun02 = eval(name_frm + ".gubun02_" + idx);
        ogubun02.value = value_gubun02;

        var ogubun01name = eval(name_frm + ".gubun01name_" + idx);
        ogubun01name.value = value_gubun01name;

        var ogubun02name = eval(name_frm + ".gubun02name_" + idx);
        ogubun02name.value = value_gubun02name;
}



function reloadMe(comp){
    var divcd = comp.value;
    var mode  = "<%= mode %>";
    var orderserial = "<%= orderserial %>";
    document.location = "?mode=" + mode + "&divcd=" + divcd + "&orderserial=" + orderserial;
}


function CheckMaxItemNo(obj, maxno) {
    if (obj.value*1 > maxno*1) {
        alert("주문수량 이상으로 상품수량을 수정할수 없습니다.");
        obj.value = maxno;
    }

    <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
    if (obj.value <0){
        alert("0개 미만 으로  수정할수 없습니다. ");
        obj.value = maxno;
    }
    <% else %>
    if (obj.value <1){
        alert("0개 이하로  수정할수 없습니다. 상품선택을 해지해주세요.");
        obj.value = maxno;
    }
    <% end if %>
}

function CheckSelect(comp){
    var chkidx = comp.value;
    var frm = document.frmaction;

    if (comp.name!="Deliverdetailidx"){
        if (comp.checked){
            eval("frm.gubun01_" + chkidx).value = frm.gubun01.value;
            eval("frm.gubun02_" + chkidx).value = frm.gubun02.value;
            eval("frm.gubun01name_" + chkidx).value = frm.gubun01name.value;
            eval("frm.gubun02name_" + chkidx).value = frm.gubun02name.value;
        }else{
            delGubun("gubun01_" + chkidx,"gubun02_" + chkidx,"gubun01name_" + chkidx,"gubun02name_" + chkidx,frm.name,causepop);
        }
    }

    //업체개별배송 체크
    CheckUpcheDeliverPay(frm);

    //배송비 체크
    CheckDeliverPay(frm);

    CalculateAndApplyItemCostSum(frm);

    //선택 브랜드 체크
    DispCheckedUpcheID(frmaction);
}

function GetCheckedUpcheBeasongPay(frm){
    var retVal = 0;
    if (!frm.Deliverdetailidx) return retVal ;

    if (frm.Deliverdetailidx.length>1){
        for (var i=0;i<frm.Deliverdetailidx.length;i++){
            if (frm.Deliverdetailidx[i].checked){
                retVal += frm.Deliveritemcost[i].value*1;
            }
        }
    }else{
        if (frm.Deliverdetailidx.checked){
            retVal = frm.Deliveritemcost.value*1;
        }
    }

    return retVal;
}

function CheckUpcheDeliverPay(frm){
    var upbeaMakerid;
    var itemMakerid;
    var NotCheckExists, isCheckValExists;
    var value_gubun02 = frm.gubun02.value;

    if (!frm.Deliverdetailidx) return;
    if (!frm.orderdetailidx) return;

    if ((!IsCancelProcess)&&(!IsReturnProcess)) return;

    if (frm.Deliverdetailidx.length>1){
        for (var i=0;i<frm.Deliverdetailidx.length;i++){
            NotCheckExists=false;
            isCheckValExists=false;
            upbeaMakerid = frm.DeliverMakerid[i].value;
            //텐배송비
            if (upbeaMakerid.length<1){
                if (frm.orderdetailidx.length>1){
                    for (var j=0;j<frm.orderdetailidx.length;j++){
                        if ((frm.odlvtype[j].value=="1")||(frm.odlvtype[j].value=="4")){
                            isCheckValExists = true;
                            NotCheckExists = (NotCheckExists)||(!frm.orderdetailidx[j].checked)||((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
                        }
                    }
                }else{
                    if ((frm.odlvtype.value=="1")||(frm.odlvtype.value=="4")){
                        isCheckValExists = true;
                        NotCheckExists = (!frm.orderdetailidx.checked)||((frm.orderdetailidx.checked == true) && (frm.itemno.value != frm.regitemno.value));
                    }
                }
                frm.Deliverdetailidx[i].checked = ((!NotCheckExists)&&(isCheckValExists));
                //반품 프로세스고 단순변심이거나
                if ((IsReturnProcess)&&((value_gubun02=="CD01")||(value_gubun02==""))){
                    frm.Deliverdetailidx[i].checked = false;
                }
                AnCheckClick(frm.Deliverdetailidx[i]);
            }else{
                if (frm.orderdetailidx.length>1){
                    for (var j=0;j<frm.orderdetailidx.length;j++){
                        itemMakerid = frm.makerid[j].value;
                        if (upbeaMakerid==itemMakerid){
                        	isCheckValExists = true;
                            NotCheckExists = (NotCheckExists)||(!frm.orderdetailidx[j].checked)||((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
                        }
                    }
                }else{
                    itemMakerid = frm.makerid.value;
                    if (upbeaMakerid==itemMakerid){
                    	isCheckValExists = true;
                        NotCheckExists = (!frm.orderdetailidx.checked)||((frm.orderdetailidx.checked == true) && (frm.itemno.value != frm.regitemno.value));
                    }
                }
                frm.Deliverdetailidx[i].checked = ((!NotCheckExists)&&(isCheckValExists));
                //반품 프로세스고 단순변심이거나
                if ((IsReturnProcess)&&((value_gubun02=="CD01")||(value_gubun02==""))){
                    frm.Deliverdetailidx[i].checked = false;
                }
                AnCheckClick(frm.Deliverdetailidx[i]);
            }
        }
    }else{
        makerAllChecked=false;
        isCheckValExists=false;
        upbeaMakerid = frm.DeliverMakerid.value;
        //텐배송비
        if (upbeaMakerid.length<1){
            if (frm.orderdetailidx.length>1){
                for (var j=0;j<frm.orderdetailidx.length;j++){
                    if ((frm.odlvtype[j].value=="1")||(frm.odlvtype[j].value=="4")){
                        isCheckValExists = true;
                        // 상품이 선택 안되어 있거나, 체크되어 있고 등록상품수와 취소상품수가 같지 않은 경우
                        NotCheckExists = (NotCheckExists)||(!frm.orderdetailidx[j].checked)||((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
                    }
                }
            }else{
                if ((frm.odlvtype.value=="1")||(frm.odlvtype.value=="4")){
                    isCheckValExists = true;
                    NotCheckExists = (!frm.orderdetailidx.checked)||((frm.orderdetailidx.checked == true) && (frm.itemno.value != frm.regitemno.value));
                }
            }
            frm.Deliverdetailidx.checked = ((!NotCheckExists)&&(isCheckValExists));
            //반품 프로세스고 단순변심이거나
            if ((IsReturnProcess)&&((value_gubun02=="CD01")||(value_gubun02==""))){
                frm.Deliverdetailidx.checked = false;
            }
            AnCheckClick(frm.Deliverdetailidx);
        }else{
            if (frm.orderdetailidx.length>1){
                for (var j=0;j<frm.orderdetailidx.length;j++){
                    itemMakerid = frm.makerid[j].value;
                    if (upbeaMakerid==itemMakerid){
                    	isCheckValExists = true;
                        NotCheckExists = (NotCheckExists)||(!frm.orderdetailidx[j].checked)||((frm.orderdetailidx[j].checked == true) && (frm.itemno[j].value != frm.regitemno[j].value));
                    }
                }
            }else{
                itemMakerid = frm.makerid.value;
                if (upbeaMakerid==itemMakerid){
                	isCheckValExists = true;
                    NotCheckExists = (!frm.orderdetailidx.checked)||((frm.orderdetailidx.checked == true) && (frm.itemno.value != frm.regitemno.value));
                }
            }

            frm.Deliverdetailidx.checked = ((!NotCheckExists)&&(isCheckValExists));
            //반품 프로세스고 단순변심이거나
            if ((IsReturnProcess)&&((value_gubun02=="CD01")||(value_gubun02==""))){
                frm.Deliverdetailidx.checked = false;
            }

            AnCheckClick(frm.Deliverdetailidx);
        }
    }

}

//배송비 체크
function CheckDeliverPay(frm){


    var allselected = IsAllSelected(frm);
    var value_gubun02 = frm.gubun02.value;

    //취소Process
    if (IsCancelProcess){
        if (allselected){
            //원 배송비 전체 환급 :체크 사용안함.
            //frm.ckbeasongpayAssign.checked = true;

            frm.milereturn.checked = true;
            frm.couponreturn.checked = true;
            //frm.allatsubtract.checked = true;
        }else{
            //원 배송비 전체 환급 :체크 사용안함.
            //frm.ckbeasongpayAssign.checked = false;

            frm.milereturn.checked = false;
            frm.couponreturn.checked = false;
            ////frm.allatsubtract.checked = false;
        }
    //반품Process
    }else if (IsReturnProcess){
        //원 배송비 전체 환급 :체크 사용안함.
        /*
        if ((allselected)&&(value_gubun02!="CD01")){
            frm.ckbeasongpayAssign.checked = true;
        }else{
            frm.ckbeasongpayAssign.checked = false;
        }
        */

        //회수배송비 관련
        if (value_gubun02=="CD01"){
            if (frm.divcd.value=="A010"){
                frm.ckreturnpay.checked = true;
            }else{
                frm.ckreturnpay.checked = false;
            }
        }else{
            frm.ckreturnpay.checked = false;
        }

    }


}

function ChangeReturnMethod(comp){
    if (comp==undefined) return;

    var returnmethod = comp.value;


    //CalculateAndApplyItemCostSum;

    document.all.refundinfo_R007.style.display = "none";
    //document.all.refundinfo_R020.style.display = "none";
    document.all.refundinfo_R050.style.display = "none";
    //document.all.refundinfo_R080.style.display = "none";
    document.all.refundinfo_R100.style.display = "none";
    document.all.refundinfo_R900.style.display = "none";

    if (comp.value=="R007"){
        //무통장 환불
        document.all.refundinfo_R007.style.display = "inline";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R400")){
        //실시간 이체 취소//ALL@ 결제 취소 //신용카드 결제 취소//휴대폰
        document.all.refundinfo_R100.style.display = "inline";
    }else if(comp.value=="R050"){
        //입점몰 결제 취소
        document.all.refundinfo_R050.style.display = "inline";
    }else if(comp.value=="R900"){
        //마일리지 환불
        document.all.refundinfo_R900.style.display = "inline";
    }

}


//CS 접수
function CsRegProc(frm){

    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return;
    }

    //선택 상품 체크
    if (!SaveCheckedItemList(frm)) {
        return;
    }

    if (IsReturnProcess){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }

    if (IsServiceDeliverProcess){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }

    if ((Fdivcd=="A009")||(Fdivcd=="A006")){
        if (!checkReturnProcessAvail(frm)){
            return;
        }
    }


    if ((IsCancelProcess)||(IsReturnProcess)){
        if (frm.subtotalprice!=undefined){
            if (frm.subtotalprice.value*1<frm.canceltotal.value*1){
                alert('실 취소 금액이 결제금액 보다 클 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            if (frm.canceltotal.value*1<0){
                alert('실 취소 금액이 0보다 작을 수 없습니다. - 쿠폰이나 마일리지 환급을 UnCheck 주세요.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            if (frm.nextsubtotal.value*1<0){
                alert('취소 후 결제 금액이 마이너스가 될 수 없습니다. - 쿠폰이나 마일리지 환급을 체크해 주세요.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            //returnmethod R400 인경우 당월 취소만 가능함
            if (frm.returnmethod){
            if (frm.returnmethod.value=="R400"){
                <% if (oordermaster.FResultCount>0) then %>
                <% if (datediff("m",oordermaster.FOneItem.FRegdate,now())>0) then %>
                alert('휴대폰 결제는 당월 취소만 가능합니다. 다른 환불방식을 선택해 주세요.');
                frm.returnmethod.focus();
                return;
                <% end if %>
                <% end if %>
            }
            }
        }
    }

    if ((frm.subtotalprice!=undefined)&&(frm.returnmethod!=undefined)){
        if (!CheckReturnMethod(frm)){
            return;
        }
    }

    //추가 정산 관련
    if (frm.add_upchejungsandeliverypay){
        if (!IsInteger(frm.add_upchejungsandeliverypay.value)){
            alert('숫자만 가능합니다.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='직접입력')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('추가 정산액이 있는경우 브랜드 아이디가 지정되어야 합니다. ');
                return;
            }

            //주문 내역에 아이디가 있는 경우만.

        }else{
            <% if (divcd="A700") then %>
            alert('추가 정산액을 입력하세요.');
            frm.add_upchejungsandeliverypay.focus();
            return;
            <% end if %>
        }
    }

    if (IsCancelProcess){
        if(confirm("취소 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if (IsReturnProcess){
        if (frm.ForceReturnByTen.checked){
            frm.requireupche.value = "N";
            frm.requiremakerid.value = "";
        }

        if (frm.requireupche.value=="Y"){
            if(confirm("업체 [" + frm.requiremakerid.value +"]로 반품/회수/교환 접수 하시겠습니까?")){
                frm.submit();
            }
        }else{
            if(confirm("[텐바이텐 물류센터]로 반품/회수/교환 접수 하시겠습니까?")){
                frm.submit();
            }
        }
    }else if (IsRefundProcess){
        if (frm.returnmethod.value.length<1){
            alert('환불 방식을 선택해 주세요.');
            frm.returnmethod.focus();
            return;
        }

        if (frm.returnmethod.value=="R000"){
            frm.refundrequire.value = "0";

        }


        if ((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R400")){
            alert('신용카드/실시간/올엣/휴대폰 환불 접수시 원 결제 금액 그대로 접수됩니다.');
            frm.refundrequire.value = frm.orgsubtotalprice.value;
        }

        if ((frm.refundrequire.value.length<1)||(!IsDigit(frm.refundrequire.value))){
            alert('환불 금액을 입력하세요.');
            frm.refundrequire.focus();
            return;
        }

        if(confirm("환불 접수 하시겠습니까?")){
            frm.submit();
        }
    }else if(confirm("접수 하시겠습니까?")){
        frm.submit();
    }
}

//업체처리완료=>접수 변경
function CsUpcheConfirm2RegProc(frm){
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

//수정
function CsRegEditProc(frm){
    if (frm.divcd.value.length<1){
        alert("접수 구분을 선택하세요.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("제목을 입력하세요.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("사유 구분을 입력하세요.");
        return;
    }

//반품 관련 수정일경우 체크함
    <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
    if (!SaveCheckedItemList(frm)) {
        return;
    }
    <% end if %>

/* 20090601 변경.*/
    //if ((frm.subtotalprice!=undefined)&&(frm.returnmethod!=undefined)){
    if ((frm.returnmethod!=undefined)){
        if (!CheckReturnForm(frm)){
            return;
        }
    }
/* */

    //추가 정산 관련
    if (frm.add_upchejungsandeliverypay){
        if (!IsDigit(frm.add_upchejungsandeliverypay.value)){
            alert('숫자만 가능합니다.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='직접입력')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('추가 정산 사유를 입력하세요.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('추가 정산액이 있는경우 브랜드 아이디가 지정되어야 합니다. ');
                return;
            }
        }
    }

    //returnmethod R400 인경우 당월 취소만 가능함
    if (frm.returnmethod){
    if (frm.returnmethod.value=="R400"){
        <% if (oordermaster.FResultCount>0) then %>
        <% if (datediff("m",oordermaster.FOneItem.FRegdate,now())>0) then %>
        alert('휴대폰 결제는 당월 취소만 가능합니다. 다른 환불방식을 선택해 주세요.');
        frm.returnmethod.focus();
        return;
        <% end if %>
        <% end if %>
    }
    }
    if (confirm('수정 하시겠습니까?')){
        frm.mode.value = "editcsas";
        frm.submit();
    }
}

function DispCheckedUpcheID(frm){
    var checkedUpcheid = "";
    var UpcheDuplicated = false;
    <% if (divcd="A004") or (divcd="A000") then %>
    var IsUpcheReturn = true;
    <% else %>
    var IsUpcheReturn = false;
    <% end if %>

    if (!frm.buf_requiremakerid) {
        return;
    }

    if (frm.orderdetailidx.length==undefined){
        if (IsUpcheReturn) {
            if (frm.isupchebeasong.value=="Y"){
                if (frm.orderdetailidx.checked){
                    checkedUpcheid = frm.makerid.value;
                }
            }
        }else{
            if (frm.orderdetailidx.checked){
                checkedUpcheid = frm.makerid.value;
            }
        }

    }else{
        for(var i=0;i<frm.orderdetailidx.length;i++){
            if (frm.orderdetailidx[i].checked){
                if (IsUpcheReturn){
                    if (frm.isupchebeasong[i].value=="Y"){
                        if (checkedUpcheid!="") {
                            if (checkedUpcheid != frm.makerid[i].value){
                                UpcheDuplicated = true;
                            }
                            checkedUpcheid = frm.makerid[i].value;
                        }
                        checkedUpcheid = frm.makerid[i].value;
                    }
                }else{
                    if (checkedUpcheid!="") {
                        if (checkedUpcheid != frm.makerid[i].value){
                            UpcheDuplicated = true;
                        }
                        checkedUpcheid = frm.makerid[i].value;
                    }
                    checkedUpcheid = frm.makerid[i].value;
                }
            }
        }
    }


    frm.buf_requiremakerid.value = "";

    if ((!UpcheDuplicated)&&(checkedUpcheid!="")){
        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = checkedUpcheid;
        }
    }
}


//회수관련 체크
function checkReturnProcessAvail(frm){
    var TenBeasongExists = false;
    var UpcheBeasongExists = false;
    var UpcheDuplicated = false;
    var checkedUpcheid = "";

    if (frm.orderdetailidx.length==undefined){
        //alert('1');
        //return true;

        // 서비스 발송시 브랜드를 지정하게..=> 상품을 선택하는것이 아니므로.
        if (frm.isupchebeasong.value=="Y"){
            UpcheBeasongExists = true;
            if (frm.orderdetailidx.checked){
                checkedUpcheid = frm.makerid.value;
            }
        }else{
            TenBeasongExists = true;
        }

    }else{
        for(var i=0;i<frm.orderdetailidx.length;i++){
            if (frm.orderdetailidx[i].checked){
                if (frm.isupchebeasong[i].value=="Y"){
                    UpcheBeasongExists = true;

                    if (checkedUpcheid!="") {
                        if (checkedUpcheid != frm.makerid[i].value){
                            UpcheDuplicated = true;
                        }

                        checkedUpcheid = frm.makerid[i].value;
                    }
                    checkedUpcheid = frm.makerid[i].value;
                }else{
                    TenBeasongExists   = true;
                }
            }
        }
    }

    if ((UpcheBeasongExists)&&(TenBeasongExists)){
        alert('텐바이텐 배송과 업체배송을 동시에 선택 하실 수 없습니다.');
        return false;
    }

    if (UpcheDuplicated){
        alert('업체배송 선택시 브랜드 별로 접수해 주세요.');
        return false;
    }

    if (checkedUpcheid!=""){
        frm.requireupche.value = "Y";
        frm.requiremakerid.value = checkedUpcheid;

        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = frm.requiremakerid.value;
        }
    }else{
        frm.requireupche.value = "N";
        frm.requiremakerid.value = "";

        if (frm.buf_requiremakerid){
            frm.buf_requiremakerid.value = "";
        }
    }
    return true;
}

//환불 관련 체크 form

function CheckReturnForm(frm){
    if (frm.returnmethod){
    if (frm.returnmethod.value=="R007"){
        var mooconfirm = false;
        if ((frm.rebankaccount!=undefined)&&(frm.rebankaccount.value.length<1)){
            //alert('환불 계좌를 입력해 주세요.');
            //frm.rebankaccount.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.rebankownername!=undefined)&&(frm.rebankownername.value.length<1)){
            //alert('예금주명을  입력해 주세요.');
            //frm.rebankownername.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.rebankname!=undefined)&&(frm.rebankname.value.length<1)){
            //alert('환불 은행을 선택해 주세요.');
            //frm.rebankname.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.refundrequire!=undefined)&&(frm.refundrequire.value.length<1)){
            alert('환불 금액을 재 계산하세요');
            return false;
        }

        if (mooconfirm){
            if (!confirm('환불 계좌가 없습니다. \n\n환불 계좌 없이 등록 하시겠습니까?')){
                frm.rebankaccount.focus();
                return false;
            }
        }

    }else if (frm.returnmethod.value=="R900"){
        if ((frm.refundbymile_userid!=undefined)&&(frm.refundbymile_userid.value.length<1)){
            alert('사용자 아이디가 없습니다. 다른 환불 방식을 선택하세요.');
            return false;
        }

        if ((frm.refundbymile_sum!=undefined)&&(frm.refundbymile_sum.value.length<1)){
            alert('환불 마일리지가 없습니다. 재계산 해주세요.');
            return false;
        }

    }

    if ((frm.returnmethod.value!="")&&(frm.returnmethod.value!="R000")&&(!IsDigit(frm.refundrequire.value))){
        alert('환불 금액은 양수(+) 만 가능합니다.');
        return false;
    }
    }
    return true;
}

//환불 관련 체크
function CheckReturnMethod(frm){
    var allselected = IsAllSelected(frm);

    var PayedNCancelEqual = (frm.subtotalprice.value*1==frm.canceltotal.value*1);

    if ((allselected)&&(!PayedNCancelEqual)&&(IsCancelProcess)){
        alert('전체 취소인경우 결제금액 전체를 환불해야합니다. - 원배송비 환급, 마일리지, 할인권 등을 체크해주세요.');
        return false;
    }

    //if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080"))){

    //if (((!allselected)||(!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080"))){
    if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R400"))){
        alert('전체 취소인 경우만 신용카드/실시간 이체/휴대폰 취소가 가능합니다. \n\n무통장 환불 또는 마일리지 환급을 선택해 주세요');
        frm.returnmethod.focus();
        return false;
    }

    if (frm.returnmethod.value.length<1){
        alert('환불 방식을 선택해 주세요.');
        frm.returnmethod.focus();
        return false;
    }

    if (!CheckReturnForm(frm)){
        return false;
    }


    <% if (oordermaster.FOneItem.FSiteName<>MAIN_SITENAME1 and oordermaster.FOneItem.FSiteName<>MAIN_SITENAME2) then %>
    //외부몰인 경우 외부몰 환불접수만 가능..
    if ((frm.returnmethod.value!="R050")&&(frm.returnmethod.value!="R000")){
        alert('외부몰인 경우 환불 없음 또는 외부몰 환불을 선택하세요. \n\n제휴 담당자를 통해 제휴몰에서 취소 환불 처리 합니다.');
        frm.returnmethod.focus();
        return
    }
    <% end if %>



    if (frm.refundrequire.value!=frm.canceltotal.value){
        if ((frm.returnmethod.value!="R007")&&(frm.returnmethod.value!="R900")&&(frm.returnmethod.value!="R000")){
            alert('환불 금액과 취소금액이 다를경우 무통장 또는 마일리지 환불만 가능합니다.');
            return false;
        }

        if (!confirm('환불 금액이 취소 금액과 다르게 수정될경우 보정치로 금액이 입력됩니다.\n\n진행 하시겠습니까?')){
            return false;
        }
    }

    return true;
}


//삭제
function CsRegCancelProc(frm){
    if (confirm('등록된 접수 내역을 삭제 하시겠습니까?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

//접수상태로 변경
function CsRegStateChg(frm){
    if (confirm('접수 상태로 변경 하시겠습니까?')){
        frm.mode.value = "state2jupsu";
        frm.submit();
    }
}


//완료처리
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    //환불요청 , 신용카드 취소요청
    if ((divcd=="A003")||(divcd=="A007")){
        if (frm.contents_finish.value.length<1){
            alert('처리 내용을 입력하세요.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '완료처리 진행 하시겠습니까?';

    if ((divcd=="A004")||(divcd=="A010")){
        confirmMsg = '완료처리 진행시 마이너스 주문 및 환불이 자동 접수됩니다. 진행 하시겠습니까?';
    }

    //20090601추가
    if (divcd="A003"){
        if (!CheckReturnForm(frm)){
            return;
        }

    }


    if (confirm(confirmMsg )){

        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "";
        frm.submit();
    }
}

function CsRegFinishProcNoRefund(frm){
    var divcd = frm.divcd.value;

    if (confirm('환불 및 마이너스 등록 없이 완료처리 진행 하시겠습니까?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

//상품 전체 체크 되었는지
function IsAllSelected(frm){
    var allselected = false;

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];
        if (e.name == "orderdetailidx") {
            if (e.checked == true) {
                    allselected = true;
            } else {
                    return false;
            }
        }
    }


    if (frm.regitemno.length==undefined){
        if (frm.regitemno.value!=frm.itemno.value){
            return false;
        }
    }else{
        for (var i = 0; i < frm.regitemno.length; i++) {
            if (frm.regitemno[i].value!=frm.itemno[i].value){
                return false;
            }
        }
    }


    return allselected;
}

function SaveCheckedItemList(frm) {
    var e;
    var ischecked = false;
    var checkitemExists = false;

    var orderdetailidx = "";
    var gubun01 = "";
    var gubun02 = "";
    var regitemno = "";
    var causecontent = "";

    frm.detailitemlist.value = "";

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            orderdetailidx = "";
            gubun01 = "";
            gubun02 = "";
            regitemno = "";
            causecontent = "";
        }

        if (e.name == "orderdetailidx") {
            if (e.checked == true) {

                ischecked = true;
                orderdetailidx = e.value;
                checkitemExists = true;
            } else {
                ischecked = false;
                orderdetailidx = "";
            }
        }

        if ((ischecked == true) && (e.name.indexOf("gubun01_") == 0)) {
            gubun01 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("gubun02_") == 0)) {
            gubun02 = e.value;
        }

        if ((ischecked == true) && (e.name.indexOf("regitemno") == 0)) {
            <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
            if ((e.value*1)<0){
                alert('수량을 입력하세요.');
                e.focus();
                e.select();
                return false;
            }
            <% else %>
            if ((e.value*1)==0){
                alert('수량을 입력하세요.');
                e.focus();
                e.select();
                return false;
            }
            <% end if %>

            regitemno = e.value;
        }

        // 사용안함.
        //if ((ischecked == true) && (e.name.indexOf("causecontent") == 0)) {
        //        causecontent = e.value;
        //}

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                frm.detailitemlist.value = frm.detailitemlist.value + "|" + orderdetailidx + "\t" + gubun01 + "\t" + gubun02 + "\t" + regitemno + "\t" + causecontent;
                ischecked = false;
                gubun01 = "";
                gubun02 = "";
                regitemno = "";
                causecontent = "";
            }
        }
    }

    //배송비 ----------------------------------------
    var upchedeliverPayStr = '';
    if (frm.Deliverdetailidx){
        if (frm.Deliverdetailidx.length>1){
            for (var i=0;i<frm.Deliverdetailidx.length;i++){
                if (frm.Deliverdetailidx[i].checked){
                    upchedeliverPayStr = frm.Deliverdetailidx[i].value + "\t" + frm.gubun01.value + "\t" + frm.gubun02.value + "\t" + "1" + "\t" + frm.Deliveritemcost[i].value;
                }
            }
        }else{
            if (frm.Deliverdetailidx.checked){
                upchedeliverPayStr = frm.Deliverdetailidx.value + "\t" + frm.gubun01.value + "\t" + frm.gubun02.value + "\t" + "1" + "\t" + frm.Deliveritemcost.value;
            }
        }
    }

    if ((upchedeliverPayStr.length>0)&&(frm.detailitemlist.value.length>0)){
        frm.detailitemlist.value = frm.detailitemlist.value + "|" + upchedeliverPayStr
    }
    //--------------------------------------------------

    //기타내역, 서비스발송 , 환불요청, 출고시유의사항, 업체 추가 정산 - 상세내역 체크 안함.
    if ((Fdivcd=="A009")||(Fdivcd=="A002")||(Fdivcd=="A003")||(Fdivcd=="A005")||(Fdivcd=="A006")||(Fdivcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('선택된 상세내역이 없습니다.');
            return false;
        }
    }

    return true;
}

//업체 추가 정산 삭제시
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";
}

function setTitle(frm,titlestr){
    frm.title.value=titlestr;
}

function getReCalcuBeasongPay(userlevel, itemsum){
    //Yellow : 5만원이상, 그린회원 : 4만원이상, 블루회원 : 3만원이상, vip회원 : 항상 무료 , 마니아 2만원이상

    //무료배송인경우 체크. 추가 해야함

    if (userlevel=="0"){
        if ((itemsum>=50000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="1"){
        if ((itemsum>=40000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="2"){
        if ((itemsum>=30000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="9"){
        if ((itemsum>=20000)||(itemsum<1)){
            return 0;
        }else{
            return 2000;
        }
    }else if (userlevel=="3"){
        return 0;
    }else{
        return 2000;
    }
}

function CalculateAndApplyItemCostSum(frm) {

    var e;
    var ischecked       = false;
    var regitemno       = 0;
    var itemno          = 0;
    var itemcost        = 0;
    var refunditemcostsum   = 0;
    var allatitemdiscount    = 0;
    var allatitemdiscountSum = 0;

    var percentBonusCouponDiscount =0;
    var percentBonusCouponDiscountSum =0;

    var orgitemcostsum     = 0;
    var orgbeasongpay       = 0;
    var refundadjustpay     = 0;

    var refundmileagesum    = 0;
    var refundcouponsum     = 0;
    var allatsubtractsum    = 0;

    var remainitemcostsum   = 0;
    var remainmileagesum    = 0;
    var remaincouponsum     = 0;
    var remainallatdiscount = 0;
    var refunddeliverypay   = 0;

    var recalcubeasongpay   = 0;
    var refundbeasongpay    = 0;

    //업체배송비 환급 추가
    var refundupchebeasongpay = GetCheckedUpcheBeasongPay(frm);

    for (var i = 0; i < frm.length; i++) {
        e = frm.elements[i];

        if (e.name == "dummystarter") {
            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
        }

        if (e.name == "orderdetailidx") {
            if (e.checked == true) {
                ischecked = true;
            }
        }

        if ((ischecked == true) && (e.name == "regitemno")) {
            if ((e.value * 0) == 0) {
                regitemno = e.value;
            } else {
                regitemno = 0;
            }
        }

        if ((ischecked == true) && (e.name == "itemno")) {
            if ((e.value * 0) == 0) {
                itemno = e.value;
            } else {
                itemno = 0;
            }
        }

        if ((ischecked == true) && (e.name == "itemcost")) {
            if ((e.value * 0) == 0) {
                itemcost = e.value;
            } else {
                itemcost = 0;
            }
        }

        if ((ischecked == true) && (e.name == "allatitemdiscount")) {
            if ((e.value * 0) == 0) {
                allatitemdiscount = e.value;
            } else {
                allatitemdiscount = 0;
            }
        }

        if ((ischecked == true) && (e.name == "percentBonusCouponDiscount")) {
            if ((e.value * 0) == 0) {
                percentBonusCouponDiscount = e.value;
            } else {
                percentBonusCouponDiscount = 0;
            }
        }

        if (e.name == "dummystopper") {
            if (ischecked == true) {
                refunditemcostsum = refunditemcostsum + (itemcost * regitemno * 1);
                allatitemdiscountSum = allatitemdiscountSum + (allatitemdiscount * regitemno * 1);
                percentBonusCouponDiscountSum = percentBonusCouponDiscountSum + (percentBonusCouponDiscount * regitemno * 1);
            }

            ischecked = false;
            regitemno = 0;
            itemno = 0;
            itemcost = 0;
        }
    }

    // 원래 상품 합계 금액
    if (frm.orgitemcostsum!=undefined){
        orgitemcostsum = frm.orgitemcostsum.value*1;
    }

    // 상품목록 하단 선택 합계
    if (frm.itemcanceltotal!=undefined){
         frm.itemcanceltotal.value = refunditemcostsum;
    }

// 등록 및 수정 시에만 재계산.
<% if (IsRegState) or (orefund.FResultCount>0)  then %>
    // 취소/반품 상품총액
    if (frm.refunditemcostsum!=undefined){
        frm.refunditemcostsum.value = refunditemcostsum;
    }

    // 선택안한(나머지) 상품총액
    if (frm.remainitemcostsum!=undefined){
        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;

        remainitemcostsum = frm.remainitemcostsum.value;
    }

    // 사용 마일리지 환원
    if (frm.milereturn!=undefined){
        if (frm.milereturn.checked){
            frm.refundmileagesum.value = frm.miletotalprice.value*-1;
        }else{
            frm.refundmileagesum.value = 0;
        }
        frm.remainmileagesum.value = (frm.miletotalprice.value*1 + frm.refundmileagesum.value*1)*-1;

        refundmileagesum = frm.refundmileagesum.value;
        remainmileagesum = frm.remainmileagesum.value;
    }

    // 사용 할인권 환원
    if (frm.couponreturn!=undefined){
        if (frm.couponreturn.checked){
            frm.refundcouponsum.value = frm.tencardspend.value*-1;
        }else{
            frm.refundcouponsum.value = 0;

            // % 할인권 사용시.
            if (percentBonusCouponDiscountSum!=0){
                frm.refundcouponsum.value = percentBonusCouponDiscountSum*-1;

                if (percentBonusCouponDiscountSum*-1==frm.tencardspend.value*-1){
                    frm.couponreturn.checked = true;
                }
            }

        }
        frm.remaincouponsum.value = (frm.tencardspend.value*1 + frm.refundcouponsum.value*1)*-1;

        refundcouponsum = frm.refundcouponsum.value;
        remaincouponsum = frm.remaincouponsum.value;
    }

    // 카드 할인 차감
    if (frm.allatsubtractsum){
//    if (frm.allatsubtract!=undefined){
//        if (frm.allatsubtract.checked){
//            frm.allatsubtractsum.value = allatitemdiscountSum*-1;
//        }else{
            frm.allatsubtractsum.value = 0;

            // 카드 할인 사용시 200906추가
            if (allatitemdiscountSum!=0){
                frm.allatsubtractsum.value = allatitemdiscountSum*-1;

                if (allatitemdiscountSum*-1==frm.allatdiscountprice.value*-1){
                    //frm.allatsubtract.checked = true;
                }
            }
//        }

        frm.remainallatdiscount.value = (frm.allatdiscountprice.value*1 + frm.allatsubtractsum.value*1)*-1;

        allatsubtractsum    = frm.allatsubtractsum.value ;
        remainallatdiscount = frm.remainallatdiscount.value ;
//    }
    }

    // 기타보정금액
    if (frm.refundadjustpay!=undefined){
        refundadjustpay = frm.refundadjustpay.value*1;
    }

    //원 배송비
    if (frm.orgbeasongpay!=undefined){
        orgbeasongpay  = frm.orgbeasongpay.value*1;
    }

    // 취소프로세스시 원 배송비 처리
    if (IsCancelProcess){

        //수정.. 업체배송 포함 인경우. : 원배송비 0인경우 있음.
        //recalcubeasongpay = getReCalcuBeasongPay('<%= oordermaster.FOneItem.FUserLevel %>',remainitemcostsum*1);

        //if (frm.recalcubeasongpay!=undefined){
        //    frm.recalcubeasongpay.value = recalcubeasongpay;
        //}

        //취소시 배송비 환급액
        //refundbeasongpay = orgbeasongpay - recalcubeasongpay;

        if (frm.ckbeasongpayAssign.checked){
            refundbeasongpay  = orgbeasongpay;
            recalcubeasongpay = 0;
        }else{
            refundbeasongpay  = refundupchebeasongpay;
            recalcubeasongpay = orgbeasongpay-refundupchebeasongpay;
        }
        refundbeasongpay  = orgbeasongpay - recalcubeasongpay;
        frm.recalcubeasongpay.value = recalcubeasongpay;
    }

    // 반품프로세스시 원 배송비 처리
    if (IsReturnProcess){

        if (frm.ckbeasongpayAssign.checked){
            refundbeasongpay  = orgbeasongpay;
            recalcubeasongpay = 0;
            frm.recalcubeasongpay.value = 0;
        }else{
            recalcubeasongpay = 0;
            refundbeasongpay  = refundupchebeasongpay;
            frm.recalcubeasongpay.value = refundupchebeasongpay;
        }
    }

    //취소 배송비 = refundbeasongpay - recalcubeasongpay
    if (frm.refundbeasongpay!=undefined){
        frm.refundbeasongpay.value = refundbeasongpay
    }

    // 회수 배송비 -
    if (frm.ckreturnpay!=undefined){
        if (frm.ckreturnpayHalf.checked){
            refunddeliverypay = CDEFAULTBEASONGPAY*-1;

        }else if (frm.ckreturnpay.checked){
            refunddeliverypay = CDEFAULTBEASONGPAY*2*-1;

        }else{
            refunddeliverypay = 0;
        }
    }


    if (frm.refunddeliverypay!=undefined){
        frm.refunddeliverypay.value = refunddeliverypay*1;
    }

    if (frm.buf_refunddeliverypay!=undefined){
        frm.buf_refunddeliverypay.value = refunddeliverypay*-1;
        //
        frm.buf_totupchejungsandeliverypay.value = frm.buf_refunddeliverypay.value*1 + frm.add_upchejungsandeliverypay.value*1;

    }

    //취소금액 합계
    if (frm.canceltotal!=undefined){
        frm.canceltotal.value  = refunditemcostsum + refundmileagesum*1 + refundcouponsum*1 + allatsubtractsum*1 + refundbeasongpay*1 + refundadjustpay*1 + refunddeliverypay*1;
    }

    //취소후 금액 합계
    if (frm.nextsubtotal!=undefined){
        frm.nextsubtotal.value = remainitemcostsum*1 + remainmileagesum*1 + remaincouponsum*1 + remainallatdiscount*1 + recalcubeasongpay*1 ;
    }

    if (parseInt(frm.ipkumdiv.value) >= 4) {
        if ((IsCancelProcess)||(IsReturnProcess)) {
            if (frm.refundrequire!=undefined){
                frm.refundrequire.value = frm.canceltotal.value*1;
            }
        }
    }

<% end if %>
}

function CheckDoubleCheck(frm,comp){
    if (comp.name=="ckreturnpay"){
        if (frm.ckreturnpay.checked){
            frm.ckreturnpayHalf.checked = false;
        }
    }else if (comp.name=="ckreturnpayHalf"){
        if (frm.ckreturnpayHalf.checked){
            frm.ckreturnpay.checked = false;
        }
    }
}

function ShowOLDCSList(){

}

//불량상품등록
function popBadItemReg(barcode,itemcount){
    var popwin = window.open('/common/do_bad_item_input.asp?mode=insert&itemcount=' + itemcount + '&itemid=' + barcode,'popBadItemReg','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//

function ChangeColor(a,b,c){
    //Nothing
}
function searchDetail(a){
    //Nothing
}

//추가정산배송비
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//추가정산배송비 사유
function Change_add_upchejungsancause(comp){
    if (comp.value=="직접입력") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// 배송추적 이력
function popDeliveryTrace(traceUrl, songjangNo)
{
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}
</script>
<body style="margin:10 10 10 10" bgcolor="#FFFFFF">
<form name="popForm" action="/cscenter/ordermaster/popDeliveryTrace.asp" target="_blank">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>
<% if (True) then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">
<form name="frmaction" method="post" action="pop_cs_process.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="">
<input type="hidden" name="id" value="<%= ocsaslist.FOneItem.Fid %>">
<input type="hidden" name="detailitemlist" value="">
<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="miletotalprice" value="<%= oordermaster.FOneItem.Fmiletotalprice %>">
<input type="hidden" name="tencardspend" value="<%= oordermaster.FOneItem.Ftencardspend %>">
<input type="hidden" name="allatdiscountprice" value="<%= oordermaster.FOneItem.Fallatdiscountprice %>">
<input type="hidden" name="requireupche" value="">
<input type="hidden" name="requiremakerid" value="">
<input type="hidden" name="orgsubtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td >
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>CS처리 요청 등록</b></td>
            <td width="140" align="right" <%= ChkIIF(ExistsRegedCSCount>1,"bgcolor='#33CC33'","") %> >
            <% if (ExistsRegedCSCount>1) then %>
                <a href="javascript:ShowOLDCSList();">기 접수된 CS 건 (<%= ExistsRegedCSCount-1 %>)</a>
            <% end if %>
            </td>
        </tr>
        </table>
    </td>
</tr>
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <% for i = 0 to (oOldcsaslist.FResultCount - 1) %>

            <% if CStr(oOldcsaslist.FItemList(i).Fid)<>id then %>
                <% if (oOldcsaslist.FItemList(i).Fdeleteyn = "Y") then %>
                <tr bgcolor="#EEEEEE" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% else %>
                <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% end if %>
                    <td height="20" nowrap><%= oOldcsaslist.FItemList(i).Fid %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= oOldcsaslist.FItemList(i).GetAsDivCDColor %>"><%= oOldcsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Forderserial %></a></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fmakerid %>"><%= Left(oOldcsaslist.FItemList(i).Fmakerid,32) %></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Fcustomername %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fuserid %>"><%= oOldcsaslist.FItemList(i).Fuserid %></acronym></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Ftitle %>"><%= oOldcsaslist.FItemList(i).Ftitle %></acronym></td>
                    <td nowrap><font color="<%= oOldcsaslist.FItemList(i).GetCurrstateColor %>"><%= oOldcsaslist.FItemList(i).GetCurrstateName %></font></td>
                    <td nowrap align="right"><%= FormatNumber(oOldcsaslist.FItemList(i).Frefundrequire,0) %></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Fregdate %>"><%= Left(oOldcsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Ffinishdate %>"><%= Left(oOldcsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
                    <td nowrap>
                    <% if oOldcsaslist.FItemList(i).Fdeleteyn="Y" then %>
                    <font color="red">삭제</font>
                    <% end if %>
                    </td>
                </tr>
            <% end if %>
        <% next %>
        </table>
    </td>
</tr>
<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">접수구분</td>
            <td bgcolor="#FFFFFF">
                <% if (IsRegState) then %>
		    	<% call drawSelectBoxCSCommCombo("divcd",divcd,"Z001","onChange='reloadMe(this);'") %>
		    	<% else %>
		    	<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
		    	<font style='line-height:100%; font-size:15px; color:blue; font-family:돋움; font-weight:bold'><%= ocsaslist.FOneItem.GetAsDivCDName %></font>
		    	&nbsp;
		    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:돋움; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>

		    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
		    	 <font style='line-height:100%; font-size:15px; color:#FF0000; font-family:돋움; font-weight:bold'>- 삭제된 내역</font>
		    	<% end if %>

		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">주문번호</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= orderserial %>
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %>]
                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수자</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자ID</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FUserID %>(<font color="<%= oordermaster.FOneItem.GetUserLevelColor %>"><%= oordermaster.FOneItem.GetUserLevelName %></font>)
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수일시</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                <%= now() %>
                <% else %>
                <%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">주문자정보</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FBuyname %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수제목</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                <input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= GetDefaultTitle(divcd, id, orderserial) %>" size="56" maxlength="56">
                <% else %>
                <input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">수령인정보</td>
            <td bgcolor="#FFFFFF">
                 <%= oordermaster.FOneItem.FReqName %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FReqHp %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">사유구분</td>
            <td bgcolor="#FFFFFF">
                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">


                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">

                <div id="causepop" style="position:absolute;"></div>

                <!-- Quick Menu -->

                <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">단순변심</a>]
                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">품절</a>]
                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]

                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">단순변심</a>]
                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
                [<a href="javascript:selectGubun('C006','CF01','배송관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오배송</a>]

                <% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">기타</a>]

                <% elseif (divcd="A001") then %>
                [<a href="javascript:selectGubun('C006','CF03','배송관련','구매상품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품누락</a>]

                <% elseif (divcd="A002") then %>
                [<a href="javascript:selectGubun('C006','CF04','배송관련','사은품누락','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(물류)사은품누락</a>]
                [<a href="javascript:selectGubun('C005','CE05','상품관련','이벤트오등록','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)이벤트오등록</a>]

                <% elseif (divcd="A000") then %>
                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품불량</a>]
                [<a href="javascript:selectGubun('C006','CF01','배송관련','오발송','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">오발송</a>]
                [<a href="javascript:selectGubun('C006','CF02','배송관련','상품파손','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">상품파손</a>]
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">결제정보</td>
            <td bgcolor="#FFFFFF">
            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>원</font>
            	<% else %>
            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>원
				<% end if %>
            	&nbsp;
                [<%= oordermaster.FOneItem.JumunMethodName %>]

                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
                (OK Cashbag사용 : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> 원)
                <% end if %>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">접수내용</td>
            <td bgcolor="#FFFFFF" rowspan="2"><textarea <% if IsFinishProcState then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea></td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">배송지정보</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
                <%= oordermaster.FOneItem.FReqZipAddr %><br>
                <%= oordermaster.FOneItem.FReqAddress %>
            </td>
        </tr>

        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">관련택배정보</td>
            <td bgcolor="#FFFFFF" valign="top">
            	<!-- 코딩 확인할것 -->
            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
			        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
			        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
			        <% dim ifindurl : ifindurl = fnGetSongjangURL(ocsaslist.FOneItem.Fsongjangdiv) %>
			        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">추적</a>
                	<% else %>
			            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">추적</a>
			        <% end if %>
			        <input type="button" class="button" value="수정" onClick="changeSongjang('<%= id %>');">
		        <% end if %>
            </td>

        </tr>

        <% if (IsFinishProcState) or (IsUpcheConfirmState) or (IsStateFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">처리내용</td>
            <td bgcolor="#FFFFFF">
            <% if (IsUpcheConfirmState) then %>
            <textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% else %>
            <textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% end if %>
            </td>
            <td bgcolor="<%= adminColor("pink") %>" align="center">처리관련<br>고객오픈<br>내용입력</td>
            <td bgcolor="#FFFFFF">
            	<table border="0" cellspacing="0" cellpadding="0" class="a" valign="top">
            	<tr>
				    <td>
				    	<input class="text" type="text" name="opentitle" value="<%= ocsaslist.FOneItem.Fopentitle %>" size="48" maxlength="60" readonly>
				    </td>
				</tr>
				<tr>
				    <td>
				    	<textarea class="textarea" name="opencontents" cols="48" rows="5" readonly><%= ocsaslist.FOneItem.Fopencontents %></textarea>
				    </td>
				</tr>
				</table>
			</td>
        </tr>
        <% end if %>
        <input type="hidden" name="orderserial" value="<%= orderserial %>" >
        <!--
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">관련주문</td>
            <td colspan="3"  bgcolor="#FFFFFF">
                <table width="100%" border="0" cellspacing="1" cellpadding="2"  bgcolor="<%= adminColor("tablebg") %>" class="a">
                <tr bgcolor="<%= adminColor("topbar") %>">
                    <td width="100">주문번호</td>
                    <td width="80">취소상태</td>
                    <td width="100">결제방법</td>
                    <td width="80">주문상태</td>
                    <td width="80">결제총액</td>
                    <td width="80">주문총액</td>
                    <td width="80">쿠폰</td>
                    <td width="80">마일리지</td>
                    <td width="80">기타카드할인</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td><input class="input_01" type="text" name="XXorderserial" value="<%= orderserial %>" size="13" maxlength="16" Readonly></td>
                    <td><font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font></td>
                    <td><%= oordermaster.FOneItem.JumunMethodName %></td>
                    <td><font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font></td>
                    <td align="right" <% if oordermaster.FOneItem.IsErrSubtotalPrice then response.write "bgcolor='red'" %> ><b><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %></b></td>
                    <td align="right"><%= FormatNumber(oordermaster.FOneItem.Ftotalsum,0) %></td>
                    <td align="right"><%= FormatNumber(oordermaster.FOneItem.Ftencardspend,0) %></td>
                    <td align="right"><%= FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) %></td>
                    <td align="right"><%= FormatNumber(oordermaster.FOneItem.Fspendmembership + oordermaster.FOneItem.Fallatdiscountprice,0) %></td>
                </tr>
                <tr bgcolor="<%= adminColor("topbar") %>">
                    <td >주문자ID</td>
                    <td >주문자명</td>
                    <td >주문자Hp</td>
                    <td >수령인</td>
                    <td >수령인Hp</td>
                    <td colspan="4">배송지주소</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td><%= oordermaster.FOneItem.FUserID %></td>
                    <td><%= oordermaster.FOneItem.FBuyname %></td>
                    <td><%= oordermaster.FOneItem.FBuyHp %></td>
                    <td><%= oordermaster.FOneItem.FReqName %></td>
                    <td><%= oordermaster.FOneItem.FReqHp %></td>
                    <td colspan="4">
                        [<%= oordermaster.FOneItem.FReqZipCode %>]
                        <%= oordermaster.FOneItem.FReqZipAddr %>
                        <%= oordermaster.FOneItem.FReqAddress %>
                    </td>
                </tr>
                </table>
            </td>
        </tr>
        -->
    <!-- 상품 상세 내역이 필요한 경우 -->
    <% if (IsItemDetailDisplay) then %>
        <% if (ocsOrderDetail.FResultCount>0) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">접수상품</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <!-- 상품 상세 시작 -->
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
                        <tr height="20" align="center" bgcolor="#F4F4F4">
                          <td width="30">선택</td>
                          <td width="50">이미지</td>
                          <td width="30">구분</td>
                          <td width="50">현상태</td>
                          <td width="50">상품코드</td>
                          <td width="90">브랜드ID</td>
                          <td>상품명<font color="blue">[옵션명]</font></td>
                          <td width="80">
                          <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
                          취소/원주문
                          <% else %>
                          접수/원주문
                          <% end if %>
                          </td>
                          <td width="60">판매가격</td>
                          <td width="130">사유구분</td>
                    	</tr>

            <% for i=0 to ocsOrderDetail.FResultCount-1 %>
                <% isAllchecked = true %>
                <% if (ocsOrderDetail.FItemList(i).Fitemid=0) then %>
                <%
                        baesongmethodstr = oordermaster.BeasongCD2Name(ocsOrderDetail.FItemList(i).Fitemoption)
                        ''원 배송비 = 배송비 Total
                        if (ocsOrderDetail.FItemList(i).FCancelyn<>"Y") then
                        orgbeasongpay = orgbeasongpay + ocsOrderDetail.FItemList(i).Fitemcost
                        end if
                %>
                        <% if (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
                        <tr align="center" bgcolor="#CCCCCC" class="gray">
                        <% else %>
                        <tr bgcolor="#FFFFFF" align="center" >
                        <% end if %>
                            <td>
                            <% if (True) or (ocsOrderDetail.FItemList(i).IsUpcheParticleDeliverPayCodeItem) then %>
                                <% if (IsRegState) then %>
                                <input type="checkbox" name="Deliverdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" <% if (Not ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %> disabled<% end if %> onClick="AnCheckClick(this); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);">
                                <% else %>
                                    <% if (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
                                    <input type="checkbox" name="Deliverdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" checked disabled >
                                    <% end if %>
                                <% end if %>
                            <input type="hidden" name="DeliverMakerid" value="<%= ocsOrderDetail.FItemList(i).FMakerid %>">
                            <input type="hidden" name="Deliveritemcost" value="<%= ocsOrderDetail.FItemList(i).Fitemcost %>">
                            <% end if %>
                            </td>
                            <td>배송비</td>
                            <td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelStateStr %></font></td>
                            <td></td>
                            <td><%= ocsOrderDetail.FItemList(i).FItemID %></td>
                            <td><%= ocsOrderDetail.FItemList(i).FMakerId %></td>
                            <td align="left">(<%= baesongmethodstr %>)</td>
                            <td ><%= ocsOrderDetail.FItemList(i).Fitemno %></td>
                            <td align="right"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></td>
                            <td></td>

                        </tr>
                <% else %>
                        <%
                            if (ocsOrderDetail.FItemList(i).FCancelyn<>"Y") then
                                orgitemcostsum = orgitemcostsum + ocsOrderDetail.FItemList(i).FItemNo*ocsOrderDetail.FItemList(i).Fitemcost
                            end if

                            regitemcostsum = regitemcostsum + ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState)*ocsOrderDetail.FItemList(i).Fitemcost
                            isDefaultCheckedItem = ocsOrderDetail.FItemList(i).IsDefaultCheckedItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd, ckAll)
                            isAllchecked = (isAllchecked And isDefaultCheckedItem)
                        %>
                        <% if (ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %>
                        <tr align="center" bgcolor="FFFFFF" <% if (isDefaultCheckedItem) then %>class="H"<% end if %>>
                        <% elseif (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
                        <tr align="center" bgcolor="#CCCCCC" class="gray">
                        <% else %>
                        <tr align="center" bgcolor="#EEEEEE" class="gray">
                        <% end if %>

                        <%
                            distinctid = ocsOrderDetail.FItemList(i).Forderdetailidx
                        %>
                            <td height="25">
                                <input type="hidden" name="dummystarter" value="">
                                <% if (IsRegState) then %>
                                <input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" <% if (isAllchecked) then %>checked<% end if %> <% if (Not ocsOrderDetail.FItemList(i).IsCheckAvailItem(oordermaster.FOneItem.FIpkumDiv,oordermaster.FOneItem.FCancelYn,divcd)) then %> disabled<% end if %> onClick="AnCheckClick(this); CheckSelect(this);">
                                <% else %>
                                    <% if (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
                                    <input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forderdetailidx %>" checked disabled >
                                    <% end if %>
                                <% end if %>
                            </td>
                            <td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ocsOrderDetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= ocsOrderDetail.FItemList(i).FSmallImage %>" width="50" border="0"></a></td>
                            <input type="hidden" name="gubun01_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01 %>">
                            <input type="hidden" name="gubun02_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02 %>">
                            <td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelStateStr %></font></td>
                            <td>
                                <font color="<%= ocsOrderDetail.FItemList(i).GetStateColor %>"><%= ocsOrderDetail.FItemList(i).GetStateName %></font>
                                <!--
                                <br>
                                (<%= ocsOrderDetail.FItemList(i).GetRegDetailStateName %>)
                                -->
                            </td>

                        	<td>
                        		<% if ocsOrderDetail.FItemList(i).Fisupchebeasong="Y" then %>
                        	    	<font color="red"><%= ocsOrderDetail.FItemList(i).Fitemid %><br>(업체)</font>
                            	<% else %>
                            		<%= ocsOrderDetail.FItemList(i).Fitemid %>
								<% end if %>
                            </td>

                            <td width="90"><acronym title="<%= ocsOrderDetail.FItemList(i).Fmakerid %>"><%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %></acronym></td>
                        	<td align="left">
                        	    <acronym title="<%= ocsOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsOrderDetail.FItemList(i).FItemName,16) %></acronym>
                            	<% if (ocsOrderDetail.FItemList(i).FItemoptionName <> "") then %>
                        	    <br>
                        	    <font color="blue">[<%= ocsOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
                            	<% end if %>
                            	<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
                        	</td>
                        	<td>
                        	    <% if (Not IsRegState) then %>
                        	        <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
                        	        <% ''반품상태/수정 모드 이면 수량 수정가능 %>
                        	        <input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, '<%= ocsOrderDetail.FItemList(i).FItemNo %>'); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);" >
                        	        <% else %>
                        	        <input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" style="text-align:center;background-color:#DDDDFF;" readonly >
                        	        <% end if %>
                        	    <% else %>
                        	    <input type="text" name="regitemno" value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, '<%= ocsOrderDetail.FItemList(i).FItemNo %>'); CheckUpcheDeliverPay(frmaction); CheckDeliverPay(frmaction); CalculateAndApplyItemCostSum(frmaction);" <% if Not ocsOrderDetail.FItemList(i).IsItemNoEditEnabled(divcd) then response.write "style='text-align:center;background-color:#DDDDFF;' readonly" %> >
                        	    <% end if %>
                        	    /
                        	    <input type="text" name="itemno" value="<%= ocsOrderDetail.FItemList(i).FItemNo %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
                        	</td>
                        	<input type="hidden" name="itemcost" value="<%= ocsOrderDetail.FItemList(i).Fitemcost %>">
                        	<!-- 국민카드 할인으로인해 변경함 -->
                        	<% if (oordermaster.FOneItem.FAccountDiv="80") or (ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice<>0) then %>
                        	<input type="hidden" name="allatitemdiscount" value="<%= ocsOrderDetail.FItemList(i).getAllAtDiscountedPrice %>">
                        	<% else %>
                        	<input type="hidden" name="allatitemdiscount" value="0">
                        	<% end if %>
                        	<input type="hidden" name="percentBonusCouponDiscount" value="<%= ocsOrderDetail.FItemList(i).getPercentBonusCouponDiscountedPrice %>">

                        	<% if (ocsOrderDetail.FItemList(i).FCancelyn="Y") then %>
                        	<td align="right"><font color="gray"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font></td>
                           	<% elseif (ocsOrderDetail.FItemList(i).FItemNo < 1) then %>
                           	<td align="right"><font color="red"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font></td>
                           	<% else %>
                           	<td align="right">
                           	    <font color="blue"><%= FormatNumber(ocsOrderDetail.FItemList(i).Fitemcost,0) %></font>
                           	    <% if ocsOrderDetail.FItemList(i).FdiscountAssingedCost<>0 and ocsOrderDetail.FItemList(i).FdiscountAssingedCost<>ocsOrderDetail.FItemList(i).Fitemcost then %>
                           	    <!-- %할인 or All@할인 : 반품시 사용값. -->
                           	    <br>(<%= FormatNumber(ocsOrderDetail.FItemList(i).FdiscountAssingedCost,0) %>)
                           	    <% end if %>
                           	</td>
                           	<% end if %>
                            <td align="center">
                                <input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
                                &gt;
                                <input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

                                <% if (IsStateFinished) and ((divcd="A010") or (divcd="A011")) and ((ocsOrderDetail.FItemList(i).Fgubun02="CE01") or (ocsOrderDetail.FItemList(i).Fgubun02="CF02")) then %>
                                <br><input type="button" class="button" value="불량등록" onClick="popBadItemReg('10<%= CHKIIF(ocsOrderDetail.FItemList(i).FItemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).FItemid),Format00(6,ocsOrderDetail.FItemList(i).FItemid)) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>');">
                                <% elseif (IsRegState) or (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
                                <a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >등록하기</div></a>
                                <% end if %>
                            </td>
                            <input type="hidden" name="isupchebeasong" value="<%= ocsOrderDetail.FItemList(i).Fisupchebeasong %>">
                            <input type="hidden" name="makerid" value="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
                            <input type="hidden" name="odlvtype" value="<%= ocsOrderDetail.FItemList(i).Fodlvtype %>">
                            <input type="hidden" name="dummystopper" value="">
                        </tr>
                <%
                end if
                %>
            <% next %>
            	<tr bgcolor="FFFFFF" height="20">
            	    <td colspan="7"></td>
            	    <td>상품합계금액</td>
            	    <td align="right"><input type="text" name="orgitemcostsum" value="<%= orgitemcostsum %>" size="7" readonly style="text-align:right;border: 1px solid #CCCCCC;" ></td>
            	    <td></td>
            	</tr>


            	<tr bgcolor="FFFFFF" height="20">
            	    <td colspan="7">
            	        &nbsp;
            	    </td>
            	    <td align="right" colspan="2">
            	        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
            	        <tr>
            	            <td>선택상품합계</td>
            	            <td align="right"><input type="text" name="itemcanceltotal" size="7" readonly style="text-align:right;border: 1px solid #333333;" ></td>
            	        </tr>
            	        </table>
            	    </td>
            	    <td>
            	    </td>
            	</tr>
            </table>
            <!-- 상품 상세 끝 -->
            </td>
           </tr>
        <% end if %>

    <% end if %>
        </table>
    </td>
</tr>

</table>

<!-- 환불 프로세스가 필요한 경우 -->
<% if (IsReFundInfoDisplay) or (IsCancelInfoDisplay) or (IsUpCheAddJungsanDisplay) then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="500" valign="top">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="5" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>취소관련 정보</b>
            </td>
        </tr>
        <% if (IsCancelInfoDisplay) then %>
            <% if (orefund.FResultCount>0) then %>
            <tr bgcolor="FFFFFF" align="center" height="23">
                <td></td>
                <td>선택</td>
                <td>원 내역</td>
                <td>취소/반품</td>
                <td>취소/반품 후</td>
            </tr>
            <% if (IsItemDetailDisplay) and (IsEditState) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('접수 금액 불일치-관리자 문의 요망');</script>
            <% end if %>
            <tr bgcolor="FFFFFF">
        		<td>상품총액</td>
        		<td width="80"></td>
        		<td align="right" width="70"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %></td>
        		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="<%= orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
        	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="<%= orefund.FOneItem.Forgitemcostsum-orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>주문시 배송비</td>
        		<td><div id="beasongpayAssign" ><input <% if (IsFinishProcState) then response.write "disabled" %> type="checkbox" name="ckbeasongpayAssign" <% if (ckAll<>"") or (orefund.FOneItem.Frefundbeasongpay>0) then response.write "checked" %> value="" onclick="CalculateAndApplyItemCostSum(frmaction);"><font color="red">환급</font></div></td>
        		<td align="right">
        		    <input type="hidden" name="orgbeasongpay" value="<%= orefund.FOneItem.Forgbeasongpay %>">
        		    <%= FormatNumber(orefund.FOneItem.Forgbeasongpay,0) %>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundbeasongpay" value="<%= orefund.FOneItem.Frefundbeasongpay %>" value="0" size="9" style="text-align:right;background-color:#DDDDFF" readonly><br>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="recalcubeasongpay" value="<%= orefund.FOneItem.Forgbeasongpay-orefund.FOneItem.Frefundbeasongpay %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>회수 배송비</td>
        		<td>
	        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)" <% if (orefund.FOneItem.Frefunddeliverypay=-4000) then response.write "checked" %> >
	        		-4000원 차감
	        		<!-- 추후 출고 배송비 차감으로 변경 -->
	        		<br>
	        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)"  <% if (orefund.FOneItem.Frefunddeliverypay=-2000) then response.write "checked" %> >
	        		-2000원 차감
        		</td>
        		<td></td>
        		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="9" style="text-align:right" style="text-align:right" ></td>
        	    <td></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>사용 마일리지 </td>
        		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="milereturn" <% if ((orefund.FOneItem.Forgmileagesum>0) and (orefund.FOneItem.Forgmileagesum+orefund.FOneItem.Frefundmileagesum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum *-1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundmileagesum" value="<%= orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text"" name="remainmileagesum" value="<%= orefund.FOneItem.Forgmileagesum*-1-orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>사용 할인권</td>
        		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="couponreturn" <% if ((orefund.FOneItem.Forgcouponsum>0) and (orefund.FOneItem.Forgcouponsum+orefund.FOneItem.Frefundcouponsum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundcouponsum" value="<%= orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remaincouponsum" value="<%= orefund.FOneItem.Forgcouponsum*-1 -orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>카드 할인금액</td>
        		<td><!-- input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="allatsubtract" <% if ((orefund.FOneItem.Fallatsubtractsum>0)  ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->할인차감</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="allatsubtractsum" value="<%= orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">

        		    <input class="text_ro" type="text" name="remainallatdiscount" value="<%= orefund.FOneItem.Forgallatdiscountsum*-1 - orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>기타보정금액</td>
        		<td></td>
        		<td align="right"></td>
        		<td align="right"><input class="text" type="text" name="refundadjustpay" value="<%= orefund.FoneItem.Frefundadjustpay %>" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
                <td align="right"></td>
        	</tr>
        	<tr bgcolor="FFFFFF">
                <td>총액/취소액</td>
                <td></td>
                <td align="right">
                    <input type="hidden" name="subtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >
                    <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>
                </td>
                <td align="right"><input class="text_ro" type="text" name="canceltotal" value="<%= orefund.FoneItem.Fcanceltotal %>" size="9" readonly style="text-align:right;background-color:#DDFFDD" ></td>
                <td align="right"><input class="text_ro" type="text" name="nextsubtotal" value="<%= oordermaster.FOneItem.Fsubtotalprice-orefund.FoneItem.Fcanceltotal %>" size="9" readonly style="text-align:right" ></td>
            </tr>
            <% else %>
            <tr bgcolor="FFFFFF">
        		<td>상품총액</td>
        		<td width="120"></td>
        		<td align="right" width="70"><%= FormatNumber(orgitemcostsum,0) %></td>
        		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="0" size="9" style="text-align:right" readonly></td>
        	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="0" size="9" style="text-align:right" readonly></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>주문시 배송비</td>
        		<td><div id="beasongpayAssign" ><input type="checkbox" name="ckbeasongpayAssign" <% if (ckAll<>"") then response.write "checked" %>  value="" onclick="CalculateAndApplyItemCostSum(frmaction);"><font color="red">배송비전체 환급</font></div></td>
        		<td align="right">
        		    <input type="hidden" name="orgbeasongpay" value="<%= orgbeasongpay %>">
        		    <%= FormatNumber(orgbeasongpay,0) %>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundbeasongpay" value="0" value="0" size="9" style="text-align:right" readonly><br>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="recalcubeasongpay" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>


        	<!-- 반품/ 회수 프로세스 -->
        	<% if (ocsaslist.FOneItem.IsReturnProcess) then %>
        	<tr bgcolor="FFFFFF">
        		<td>회수 배송비</td>
        		<td>
        			<input type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
        			-4000원 차감
            		<br>
            		<input type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
            		-2000원 차감
        		</td>
        		<td></td>
        		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="0" size="9" style="text-align:right" style="text-align:right" readonly></td>
        	    <td></td>
        	</tr>
        	<% end if %>

        	<% if (ocsaslist.FOneItem.IsCancelProcess) or (ocsaslist.FOneItem.IsReturnProcess) then %>
        	<tr bgcolor="FFFFFF">
        		<td>사용 마일리지</td>
        		<td><input type="checkbox" name="milereturn" <% if ((oordermaster.FOneItem.FMileTotalPrice>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FMileTotalPrice * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundmileagesum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remainmileagesum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>사용 할인권</td>
        		<td><input type="checkbox" name="couponreturn" <% if ((oordermaster.FOneItem.FTenCardSpend>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">환원</td>
        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FTenCardSpend * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundcouponsum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remaincouponsum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>카드 할인</td>
        		<td><!-- input type="checkbox" name="allatsubtract" <% if ((oordermaster.FOneItem.Fallatdiscountprice>0) and (ocsaslist.FOneItem.IsCancelProcess) ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->차감</td>
        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FAllatDiscountPrice * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="allatsubtractsum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remainallatdiscount" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>
    	    <% end if %>


    	<tr bgcolor="FFFFFF">
    		<td>기타보정금액</td>
    		<td></td>
    		<td align="right"></td>
    		<td align="right"><input class="text" type="text" name="refundadjustpay" value="0" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
            <td align="right"></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
            <td>총액/취소액</td>
            <td></td>
            <td align="right">
                <input type="hidden" name="subtotalprice" value="<%= oordermaster.FOneItem.Fsubtotalprice %>" >
                <%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>
            </td>
            <td align="right"><input class="text_ro" type="text" name="canceltotal" size="9" readonly style="text-align:right" readonly></td>
            <td align="right"><input class="text_ro" type="text" name="nextsubtotal" size="9" readonly style="text-align:right"  readonly></td>
        </tr>
    	<% end if %>
      <% end if %>
      </table>
    </td>
    <td bgcolor="#FFFFFF" valign="top" align="left">
        <% if (divcd<>"A700") then ''업체 기타정산  %>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	        <tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>환불관련 정보</b>
	            </td>
	        </tr>
	        <% if (IsReFundInfoDisplay) then %>
	        <tr bgcolor="#FFFFFF">
	            <td width="100">환불방식</td>
	            <td>
	                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,oordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
	                <% if (Not IsRegState) then %>
	                (<%= orefund.FOneItem.FreturnmethodName %>)
	                <% end if %>
	                <input name="RefundRecalcuButton" class="csbutton" type="button" value="재계산" onClick="CalculateAndApplyItemCostSum(frmaction);">
	            </td>
	        </tr>
	        <tr  bgcolor="FFFFFF" id="refundinfo_R007" <% if orefund.FOneItem.Freturnmethod="R007" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	            <td width="100">은행정보</td>
	            <td align="left">
	                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
		            	<tr bgcolor="FFFFFF">
		            		<td width="80">계좌번호</td>
		            		<td>
		            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
		            		    <input class="csbutton" type="button" value="이전내역" onClick="popPreReturnAcct('<%= oordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
		            		</td>
		            	</tr>
		            	<tr bgcolor="FFFFFF">
		            		<td>예금주명</td>
		            		<td><input class="text" type="text" size="20" name="rebankownername" value="<%= orefund.FOneItem.Frebankownername %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>></td>
		            	</tr>
		                <tr bgcolor="FFFFFF">
		            		<td>거래은행</td>
		            		<td><% DrawBankCombo "rebankname", orefund.FOneItem.Frebankname %></td>
		            	</tr>
	            	</table>
	            </td>

	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	    		<td width="100">PG사 ID</td>
	    		<td><input class="text_ro" type="text" name="paygateTid" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
	            <td colspan="2" align="left">외부몰 환불요청</td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
	    		<td width="100">아이디</td>
	    		<td><input class="text_ro" type="text" name="refundbymile_userid" value="<%= oordermaster.FOneItem.Fuserid %>" readonly></td>
	        </tr>
	        <tr bgcolor="FFFFFF">
	    		<td width="100">환불 예정액</td>
	    		<% if (orefund.FResultCount>0) then %>
	    		<td>
	    		    <input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" maxlength=7 >
	    		    (<%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>)
	    		</td>
	    		<% else %>
	    		<td><input class="text_ro" type="text" size="10" name="refundrequire" value="<%= orefund.FOneItem.Frefundrequire %>" <% if (Not ocsaslist.FOneItem.IsRefundProcess) then response.write "readonly" %> ></td>
	    		<% end if %>
	    	</tr>
	    	<% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
	        <tr bgcolor="FFFFFF">
	    	    <td colspan="2"><b>환불 파일 작성중이므로 수정 할 수 없습니다.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
	    	</tr>
	        <% end if %>

	    	<% if (orefund.FResultCount>0) then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td colspan="2"><b>(환불예정액 수정시 기타보정금액으로 차액이 입력됩니다.)</b></td>
	    	</tr>
	    	<% end if %>

	        	<% if (IsFinishProcState) then %>
	        	    <script language='javascript'>
	        	    frmaction.returnmethod.disabled=true;
	        	    frmaction.RefundRecalcuButton.disabled=true;
	        	    frmaction.rebankaccount.disabled=true;
	        	    frmaction.rebankname.disabled=true;
	        	    frmaction.rebankownername.disabled=true;
	        	    frmaction.refundrequire.disabled=true;
	        	    frmaction.paygateTid.disabled=true;
	        	    frmaction.refundbymile_userid.disabled=true;

	        	    if ((Fdivcd=="A003")&&(frmaction.returnmethod.value=="R900")){
	        	        alert('마일리지 환급은 완료처리시 자동 환급 됩니다.');
	        	    }

	        	    if ((Fdivcd=="A003")&&(frmaction.returnmethod.value=="R007")){
	        	        alert('무통장 환불 완료처리시 문자메세지를 발송해 주세요.');
	        	    }
	        	    </script>
	        	<% end if %>
	    	<% else %>
	        <tr bgcolor="FFFFFF" ><td align="center">환불 접수 불가 또는 결제 이전 상태 </td></tr>
	        <% end if %>
        </table>
        <% end if %>

        <p>

        <% if (IsUpCheAddJungsanDisplay) then %>
    	<!-- 업체 반품인경우 -->
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    		<tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>업체 추가 정산 내역</b>
	            </td>
	        </tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">브랜드ID</td>
	    	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    	    <% if (divcd="A700") then %>
	    	    <input type="button" class="button" value="브랜드ID검색" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    	    <% end if %>
	    	    </td>
	    	</tr>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">회수배송비</td>
	    	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">추가정산배송비</td>
	    	    <td ><input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">원
	    	    &nbsp;
	    	    <select class="select" name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
		    	    <option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>사유선택
		    	    <option value="추가배송비" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="추가배송비","selected","") %> >추가배송비
		    	    <option value="추가운임" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="추가운임","selected","") %>>추가운임
		    	    <option value="직접입력" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>직접입력
	    	    </select>

	    	    <span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'><input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가배송비" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"추가운임" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" ></span>
	    	    <a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">총정산배송비</td>
	    	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >원</td>
	    	</tr>
    	</table>

        	<% if (IsFinishProcState) then %>
            	    <script language='javascript'>
            	    frmaction.buf_refunddeliverypay.disabled=true;
        	        frmaction.add_upchejungsandeliverypay.disabled=true;
        	        frmaction.add_upchejungsancause.disabled=true;
        	        frmaction.buf_totupchejungsandeliverypay.disabled=true;
            	    </script>
            <% end if %>
    	<% end if %>

        <% if (divcd="A010") then %>
        <br>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
        <tr  bgcolor="FFFFFF" >
            <td>
            <input type="checkbox" name="ForceReturnByTen"><font color="red">업체배송 상품이라도 텐바이텐 물류센터로 회수할 경우 이곳을 체크.</font>
            </td>
        </tr>
        </table>
        <% else %>
        <input type="hidden" name="ForceReturnByTen">
        <% end if %>

    </td>
</tr>
</table>
<% end if %>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td colspan="4" align="center">
    <% if (IsRegState or IsFinishProcState) and _
        ((divcd="A000") or (divcd="A001") or _
        (divcd="A002") or (divcd="A003") or _
        (divcd="A004") or (divcd="A007") or _
        (divcd="A008") or (divcd="A010") or _
        (divcd="A011")) then %>

        <% if ((not (IsRegState)) and (datediff("d", ocsaslist.FOneItem.Fregdate, now()) > 21)) then %>
        <input type="checkbox" name="csmailsend" value="on" > CS 접수/처리 이메일 발송
        <font color=red>(필요한경우 체크하세요. 접수일과 처리일의 차이가 3주 초과)</font>
        <% else %>
        <input type="checkbox" name="csmailsend" value="on" <%= chkIIF(oordermaster.FOneItem.FSiteName="10x10","checked","") %> > CS 접수/처리 이메일 발송
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr>
    <td colspan="4" align="center">
    <% if (IsRegState) then %>
        <% if (IsJupsuProcessAvail) then %>
        <input class="csbutton" type="button" value=" 접 수 " onClick="CsRegProc(frmaction)">
        <% else %>
            <% if JupsuInValidMsg<>"" then %>
            <font color="red"><%= JupsuInValidMsg %></font>
            <script language='javascript'>alert('<%= JupsuInValidMsg %>');</script>
            <% end if %>
        <% end if %>
    <% elseif (Not IsStateFinished) then %>
        <% if (ocsaslist.FOneITem.FDeleteyn="N") then %>
            <% if (mode="finishreginfo") then %>


                <% if (divcd="A004") or (divcd="A010") then %>
                    <input class="csbutton" type="button" value=" 완료 처리 (마이너스/환불요청 등록)" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
                    <input class="csbutton" type="button" value=" [마이너스/환불요청 없는] 완료 처리 " onClick="CsRegFinishProcNoRefund(frmaction)" onFocus="blur()">
                <% else %>
                    <input class="csbutton" type="button" value=" 완료 처리 " onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
                <% end if %>
            <% else %>
                <% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
                환불파일 작성중이므로 수정 불가 합니다.
                <% else %>
                <input class="csbutton" type="button" value=" 접수 취소 " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" 접수내용 수정 " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                    <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input class="csbutton" type="button" value=" 접수상태로 변경 " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
                    <% end if %>
                <% end if %>
            <% end if %>


        <% else %>

        <% end if %>
    <% elseif (IsStateFinished) then %>
        <% if (divcd="A700") and (mode<>"finishreginfo") then %>
        <!--
            <input class="csbutton" type="button" value=" 접수 상태로 변경 " onClick="CsRegStateChg(frmaction)" onFocus="blur()">
          -->
        <% end if %>
    <% end if %>
    </td>
</tr>
</form>
</table>

<script language='javascript'>
function getOnload(){
<% if IsRegState then %>
    CalculateAndApplyItemCostSum(frmaction);

    ChangeReturnMethod(frmaction.returnmethod);
<% end if %>

<% if (IsFinishProcState) and ((divcd="A007") or (divcd="A003")) then %>
    alert('이곳에서 완료처리 하여도 \n\n\n신용카드 승인취소 및 무통장환불처리는 이루어 지지 않으니 유의하시기 바랍니다.!\n\n\n\n\n\n ');
<% end if %>
}
window.onload = getOnload;

<% if (ocsaslist.FOneITem.FDeleteyn="Y") then %>
alert('삭제된 내역입니다.');
<% end if %>
</script>

<% else %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="center">[ Error : Not Valid Param ]</td>
</tr>
</table>
<% end if %>


</body>
<%
set ocsaslist = Nothing
set ocsOrderDetail = Nothing
set oordermaster = Nothing
set orefund = Nothing
set oOldcsaslist = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
