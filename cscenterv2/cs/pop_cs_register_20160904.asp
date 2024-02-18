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
'' A008     �ֹ����
'' A020     ��ü��� - �ֹ���ҷ� ����
'' A021     �κ���� - �ֹ���ҷ� ����

'' A004     ��ǰ����
'' A006     �������ǻ���
'' A009     ��Ÿ����(Memo)
'' A010     ȸ����û  -  �ٹ����� ��۸� ����?
'' A011     �±�ȯȸ��

'' A700     ��ü��Ÿ����

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


''������ ������� �űԵ������ ������
if (ocsaslist.FResultCount<1) then
    set ocsaslist.FOneItem = new CCSASMasterItem
    ocsaslist.FOneItem.FId=0
    ocsaslist.FOneItem.Fdivcd = divcd
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    orderserial = ocsaslist.FOneItem.Forderserial
end if


''������� �������� ����
dim IsRegState
IsRegState = (ocsaslist.FOneItem.FId=0)



''�ֹ� ����Ÿ
dim oordermaster
set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' ���� 6���� ���� ���� �˻�
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

''���� ���¿����� ��ü �ֹ���� / ����, �Ϸ���¿����� ������ ������ ������
if (IsRegState) then
    ocsOrderDetail.GetOrderDetailByCsDetail
else
    ocsOrderDetail.GetCsDetailList
end if


''ȯ������
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


''���� ���� ����
dim IsEditState
IsEditState = (Not IsRegState) and ((mode="editreginfo") or (mode="editrefundinfo"))

''�Ϸ�ó�� ����
dim IsFinishProcState
IsFinishProcState = (Not IsRegState) and (mode="finishreginfo")

''�Ϸ��������
dim IsStateFinished
IsStateFinished = (ocsaslist.FOneItem.FCurrState="B007")

''��üó���Ϸ��������
dim IsUpcheConfirmState
IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")

''detail's distinct id
dim distinctid

''���� �Ұ��� �޼���
dim JupsuInValidMsg

if (Left(orderserial,1)<>"A") and (oordermaster.FResultCount<1) and (mode<>"editrefundinfo") then
    response.write "<br><br>!!! ���� �ֹ������̰ų� �ֹ� ������ �����ϴ�. - ������ ���� ���"
    dbget.close()	:	response.End
end if

''���� ���� ���� ''�ֹ������� ������� üũ.
dim IsJupsuProcessAvail

if (oordermaster.FResultCount>0) then
    IsJupsuProcessAvail = ocsaslist.FOneItem.IsAsRegAvail(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
else
    IsJupsuProcessAvail = false
end if


'' ��ۺ�, ��ۿɼ�
dim baesongmethodstr,orgbeasongpay

'' ���ֹ� ��ǰ�ݾ�
dim orgitemcostsum

'' ������ǰ �հ�ݾ�
dim regitemcostsum

dim isDefaultCheckedItem,isAllchecked

''�� ������ CS�� �ִ��� Ȯ��
dim oOldcsaslist
set oOldcsaslist = New CCSASList
oOldcsaslist.FRectNotCsID     = id
oOldcsaslist.FRectOrderserial = orderserial
oOldcsaslist.GetCSASMasterList

dim ExistsRegedCSCount
ExistsRegedCSCount = oOldcsaslist.FResultCount


''��� �������� Display����
dim IsCancelInfoDisplay
IsCancelInfoDisplay = ((IsRegState) or (orefund.FResultCount>0))
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A000")       '' �±�ȯ
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A001")       '' ����
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A002")       '' ����
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A009")       '' ��Ÿ�޸�
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A003")       '' ȯ������
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A005")       '' �ܺθ�ȯ������
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A006")       '' ���� ���ǻ���
IsCancelInfoDisplay = IsCancelInfoDisplay and (divcd<>"A700")       '' ��ü ��Ÿ����


''ȯ�� ����  ǥ�� :
dim IsReFundInfoDisplay
if (oordermaster.FResultCount>0) then
    IsReFundInfoDisplay = ocsaslist.FOneItem.IsRefundProcessRequire(oordermaster.FOneItem.Fipkumdiv,oordermaster.FOneItem.FCancelyn)
else
    IsReFundInfoDisplay = false
end if

IsReFundInfoDisplay = (IsReFundInfoDisplay and IsJupsuProcessAvail)
IsReFundInfoDisplay = IsReFundInfoDisplay or (divcd="A003") or (divcd="A005")
IsReFundInfoDisplay = IsReFundInfoDisplay or (orefund.FResultCount>0)

''��Ÿ���� ǥ�� :
dim IsUpCheAddJungsanDisplay
IsUpCheAddJungsanDisplay = (divcd="A004") or (divcd="A700") or (divcd="A000") ''��ǰ����, ��ü ��Ÿ����

''��ǰ ��� display ����
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
        PopCSSMSSend('<%= oordermaster.FOneItem.Freqhp %>','<%= orderserial %>','<%= oordermaster.FOneItem.Fuserid %>','�ٹ������Դϴ�. ���� ȯ���� �Ϸ�Ǿ����ϴ�. ��ſ� �Ϸ� �Ǽ��� �����մϴ�.^^*')
    }
}

FinishActType('<%= RequestCheckvar(request("finishtype"),10) %>');
<% end if %>
function divCsAsGubunSelect(value_gubun01,value_gubun02,name_gubun01,name_gubun02,name_gubun01name,name_gubun02name,name_frm,targetDiv){
<% if (IsFinishProcState) then %>
    alert('����â���� �������ּ���. - �Ϸ�ó���� �����Ұ�');
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
    alert("�����͸� �д� ���߿� ������ �߻��߽��ϴ�. ����� �ٽ� �õ��غ��ñ� �ٶ��ϴ�.[CODE:" + xmlHttp.status + "]");
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
    alert('����â���� �������ּ���. - �Ϸ�ó���� �����Ұ�');
    return;
<% end if %>
    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;

    eval(targetDiv).innerHTML = "";

    //����Ÿ���� ������ ��� ��ü ���õ� Detail�� ����
    if (targetDiv=="causepop"){
        for (var i=0;i<frm.elements.length;i++){
            var e = frm.elements[i];

            if ((e.type=="checkbox")&&(e.checked)&&(e.name=="orderdetailidx")){
                setDetailCause(e.value,value_gubun01,value_gubun02,value_gubun01name,value_gubun02name, name_frm);
            }
        }
    }

    //�Ϸ�ÿ��� �ݾ� ���� �Ұ�
    //��ü������� üũ
    CheckUpcheDeliverPay(frm);

    //��ۺ� üũ
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
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }

    <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
    if (obj.value <0){
        alert("0�� �̸� ����  �����Ҽ� �����ϴ�. ");
        obj.value = maxno;
    }
    <% else %>
    if (obj.value <1){
        alert("0�� ���Ϸ�  �����Ҽ� �����ϴ�. ��ǰ������ �������ּ���.");
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

    //��ü������� üũ
    CheckUpcheDeliverPay(frm);

    //��ۺ� üũ
    CheckDeliverPay(frm);

    CalculateAndApplyItemCostSum(frm);

    //���� �귣�� üũ
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
            //�ٹ�ۺ�
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
                //��ǰ ���μ����� �ܼ������̰ų�
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
                //��ǰ ���μ����� �ܼ������̰ų�
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
        //�ٹ�ۺ�
        if (upbeaMakerid.length<1){
            if (frm.orderdetailidx.length>1){
                for (var j=0;j<frm.orderdetailidx.length;j++){
                    if ((frm.odlvtype[j].value=="1")||(frm.odlvtype[j].value=="4")){
                        isCheckValExists = true;
                        // ��ǰ�� ���� �ȵǾ� �ְų�, üũ�Ǿ� �ְ� ��ϻ�ǰ���� ��һ�ǰ���� ���� ���� ���
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
            //��ǰ ���μ����� �ܼ������̰ų�
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
            //��ǰ ���μ����� �ܼ������̰ų�
            if ((IsReturnProcess)&&((value_gubun02=="CD01")||(value_gubun02==""))){
                frm.Deliverdetailidx.checked = false;
            }

            AnCheckClick(frm.Deliverdetailidx);
        }
    }

}

//��ۺ� üũ
function CheckDeliverPay(frm){


    var allselected = IsAllSelected(frm);
    var value_gubun02 = frm.gubun02.value;

    //���Process
    if (IsCancelProcess){
        if (allselected){
            //�� ��ۺ� ��ü ȯ�� :üũ ������.
            //frm.ckbeasongpayAssign.checked = true;

            frm.milereturn.checked = true;
            frm.couponreturn.checked = true;
            //frm.allatsubtract.checked = true;
        }else{
            //�� ��ۺ� ��ü ȯ�� :üũ ������.
            //frm.ckbeasongpayAssign.checked = false;

            frm.milereturn.checked = false;
            frm.couponreturn.checked = false;
            ////frm.allatsubtract.checked = false;
        }
    //��ǰProcess
    }else if (IsReturnProcess){
        //�� ��ۺ� ��ü ȯ�� :üũ ������.
        /*
        if ((allselected)&&(value_gubun02!="CD01")){
            frm.ckbeasongpayAssign.checked = true;
        }else{
            frm.ckbeasongpayAssign.checked = false;
        }
        */

        //ȸ����ۺ� ����
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
        //������ ȯ��
        document.all.refundinfo_R007.style.display = "inline";
    }else if((comp.value=="R020")||(comp.value=="R080")||(comp.value=="R100")||(comp.value=="R400")){
        //�ǽð� ��ü ���//ALL@ ���� ��� //�ſ�ī�� ���� ���//�޴���
        document.all.refundinfo_R100.style.display = "inline";
    }else if(comp.value=="R050"){
        //������ ���� ���
        document.all.refundinfo_R050.style.display = "inline";
    }else if(comp.value=="R900"){
        //���ϸ��� ȯ��
        document.all.refundinfo_R900.style.display = "inline";
    }

}


//CS ����
function CsRegProc(frm){

    if (frm.divcd.value.length<1){
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return;
    }

    //���� ��ǰ üũ
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
                alert('�� ��� �ݾ��� �����ݾ� ���� Ŭ �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            if (frm.canceltotal.value*1<0){
                alert('�� ��� �ݾ��� 0���� ���� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� UnCheck �ּ���.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            if (frm.nextsubtotal.value*1<0){
                alert('��� �� ���� �ݾ��� ���̳ʽ��� �� �� �����ϴ�. - �����̳� ���ϸ��� ȯ���� üũ�� �ּ���.');
                <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"iroo4") then %>
                return;
                <% end if %>
            }

            //returnmethod R400 �ΰ�� ��� ��Ҹ� ������
            if (frm.returnmethod){
            if (frm.returnmethod.value=="R400"){
                <% if (oordermaster.FResultCount>0) then %>
                <% if (datediff("m",oordermaster.FOneItem.FRegdate,now())>0) then %>
                alert('�޴��� ������ ��� ��Ҹ� �����մϴ�. �ٸ� ȯ�ҹ���� ������ �ּ���.');
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

    //�߰� ���� ����
    if (frm.add_upchejungsandeliverypay){
        if (!IsInteger(frm.add_upchejungsandeliverypay.value)){
            alert('���ڸ� �����մϴ�.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='�����Է�')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('�߰� ������� �ִ°�� �귣�� ���̵� �����Ǿ�� �մϴ�. ');
                return;
            }

            //�ֹ� ������ ���̵� �ִ� ��츸.

        }else{
            <% if (divcd="A700") then %>
            alert('�߰� ������� �Է��ϼ���.');
            frm.add_upchejungsandeliverypay.focus();
            return;
            <% end if %>
        }
    }

    if (IsCancelProcess){
        if(confirm("��� ���� �Ͻðڽ��ϱ�?")){
            frm.submit();
        }
    }else if (IsReturnProcess){
        if (frm.ForceReturnByTen.checked){
            frm.requireupche.value = "N";
            frm.requiremakerid.value = "";
        }

        if (frm.requireupche.value=="Y"){
            if(confirm("��ü [" + frm.requiremakerid.value +"]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                frm.submit();
            }
        }else{
            if(confirm("[�ٹ����� ��������]�� ��ǰ/ȸ��/��ȯ ���� �Ͻðڽ��ϱ�?")){
                frm.submit();
            }
        }
    }else if (IsRefundProcess){
        if (frm.returnmethod.value.length<1){
            alert('ȯ�� ����� ������ �ּ���.');
            frm.returnmethod.focus();
            return;
        }

        if (frm.returnmethod.value=="R000"){
            frm.refundrequire.value = "0";

        }


        if ((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R400")){
            alert('�ſ�ī��/�ǽð�/�ÿ�/�޴��� ȯ�� ������ �� ���� �ݾ� �״�� �����˴ϴ�.');
            frm.refundrequire.value = frm.orgsubtotalprice.value;
        }

        if ((frm.refundrequire.value.length<1)||(!IsDigit(frm.refundrequire.value))){
            alert('ȯ�� �ݾ��� �Է��ϼ���.');
            frm.refundrequire.focus();
            return;
        }

        if(confirm("ȯ�� ���� �Ͻðڽ��ϱ�?")){
            frm.submit();
        }
    }else if(confirm("���� �Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

//��üó���Ϸ�=>���� ����
function CsUpcheConfirm2RegProc(frm){
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "upcheconfirm2jupsu";
        frm.submit();
    }
}

//����
function CsRegEditProc(frm){
    if (frm.divcd.value.length<1){
        alert("���� ������ �����ϼ���.");
        frm.divcd.focus();
        return;
    }

    if (frm.title.value.length<1) {
        alert("������ �Է��ϼ���.");
        frm.title.focus();
        return;
    }

    if (frm.gubun01.value.length<1) {
        alert("���� ������ �Է��ϼ���.");
        return;
    }

//��ǰ ���� �����ϰ�� üũ��
    <% if (IsEditState) and (ocsaslist.FOneItem.IsReturnProcess) then %>
    if (!SaveCheckedItemList(frm)) {
        return;
    }
    <% end if %>

/* 20090601 ����.*/
    //if ((frm.subtotalprice!=undefined)&&(frm.returnmethod!=undefined)){
    if ((frm.returnmethod!=undefined)){
        if (!CheckReturnForm(frm)){
            return;
        }
    }
/* */

    //�߰� ���� ����
    if (frm.add_upchejungsandeliverypay){
        if (!IsDigit(frm.add_upchejungsandeliverypay.value)){
            alert('���ڸ� �����մϴ�.');
            frm.add_upchejungsandeliverypay.focus();
            return;
        }

        if (frm.add_upchejungsandeliverypay.value*1!=0){
            if (frm.add_upchejungsancause.value.length<1){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancause.focus();
                return;
            }else if ((frm.add_upchejungsancause.value=='�����Է�')&&(frm.add_upchejungsancauseText.value.length<1)){
                alert('�߰� ���� ������ �Է��ϼ���.');
                frm.add_upchejungsancauseText.focus();
                return;
            }

            if (frm.buf_requiremakerid.value.length<1){
                alert('�߰� ������� �ִ°�� �귣�� ���̵� �����Ǿ�� �մϴ�. ');
                return;
            }
        }
    }

    //returnmethod R400 �ΰ�� ��� ��Ҹ� ������
    if (frm.returnmethod){
    if (frm.returnmethod.value=="R400"){
        <% if (oordermaster.FResultCount>0) then %>
        <% if (datediff("m",oordermaster.FOneItem.FRegdate,now())>0) then %>
        alert('�޴��� ������ ��� ��Ҹ� �����մϴ�. �ٸ� ȯ�ҹ���� ������ �ּ���.');
        frm.returnmethod.focus();
        return;
        <% end if %>
        <% end if %>
    }
    }
    if (confirm('���� �Ͻðڽ��ϱ�?')){
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


//ȸ������ üũ
function checkReturnProcessAvail(frm){
    var TenBeasongExists = false;
    var UpcheBeasongExists = false;
    var UpcheDuplicated = false;
    var checkedUpcheid = "";

    if (frm.orderdetailidx.length==undefined){
        //alert('1');
        //return true;

        // ���� �߼۽� �귣�带 �����ϰ�..=> ��ǰ�� �����ϴ°��� �ƴϹǷ�.
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
        alert('�ٹ����� ��۰� ��ü����� ���ÿ� ���� �Ͻ� �� �����ϴ�.');
        return false;
    }

    if (UpcheDuplicated){
        alert('��ü��� ���ý� �귣�� ���� ������ �ּ���.');
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

//ȯ�� ���� üũ form

function CheckReturnForm(frm){
    if (frm.returnmethod){
    if (frm.returnmethod.value=="R007"){
        var mooconfirm = false;
        if ((frm.rebankaccount!=undefined)&&(frm.rebankaccount.value.length<1)){
            //alert('ȯ�� ���¸� �Է��� �ּ���.');
            //frm.rebankaccount.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.rebankownername!=undefined)&&(frm.rebankownername.value.length<1)){
            //alert('�����ָ���  �Է��� �ּ���.');
            //frm.rebankownername.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.rebankname!=undefined)&&(frm.rebankname.value.length<1)){
            //alert('ȯ�� ������ ������ �ּ���.');
            //frm.rebankname.focus();
            mooconfirm=true;
            //return false;
        }

        if ((frm.refundrequire!=undefined)&&(frm.refundrequire.value.length<1)){
            alert('ȯ�� �ݾ��� �� ����ϼ���');
            return false;
        }

        if (mooconfirm){
            if (!confirm('ȯ�� ���°� �����ϴ�. \n\nȯ�� ���� ���� ��� �Ͻðڽ��ϱ�?')){
                frm.rebankaccount.focus();
                return false;
            }
        }

    }else if (frm.returnmethod.value=="R900"){
        if ((frm.refundbymile_userid!=undefined)&&(frm.refundbymile_userid.value.length<1)){
            alert('����� ���̵� �����ϴ�. �ٸ� ȯ�� ����� �����ϼ���.');
            return false;
        }

        if ((frm.refundbymile_sum!=undefined)&&(frm.refundbymile_sum.value.length<1)){
            alert('ȯ�� ���ϸ����� �����ϴ�. ���� ���ּ���.');
            return false;
        }

    }

    if ((frm.returnmethod.value!="")&&(frm.returnmethod.value!="R000")&&(!IsDigit(frm.refundrequire.value))){
        alert('ȯ�� �ݾ��� ���(+) �� �����մϴ�.');
        return false;
    }
    }
    return true;
}

//ȯ�� ���� üũ
function CheckReturnMethod(frm){
    var allselected = IsAllSelected(frm);

    var PayedNCancelEqual = (frm.subtotalprice.value*1==frm.canceltotal.value*1);

    if ((allselected)&&(!PayedNCancelEqual)&&(IsCancelProcess)){
        alert('��ü ����ΰ�� �����ݾ� ��ü�� ȯ���ؾ��մϴ�. - ����ۺ� ȯ��, ���ϸ���, ���α� ���� üũ���ּ���.');
        return false;
    }

    //if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080"))){

    //if (((!allselected)||(!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080"))){
    if (((!PayedNCancelEqual))&&((frm.returnmethod.value=="R100")||(frm.returnmethod.value=="R020")||(frm.returnmethod.value=="R080")||(frm.returnmethod.value=="R400"))){
        alert('��ü ����� ��츸 �ſ�ī��/�ǽð� ��ü/�޴��� ��Ұ� �����մϴ�. \n\n������ ȯ�� �Ǵ� ���ϸ��� ȯ���� ������ �ּ���');
        frm.returnmethod.focus();
        return false;
    }

    if (frm.returnmethod.value.length<1){
        alert('ȯ�� ����� ������ �ּ���.');
        frm.returnmethod.focus();
        return false;
    }

    if (!CheckReturnForm(frm)){
        return false;
    }


    <% if (oordermaster.FOneItem.FSiteName<>MAIN_SITENAME1 and oordermaster.FOneItem.FSiteName<>MAIN_SITENAME2) then %>
    //�ܺθ��� ��� �ܺθ� ȯ�������� ����..
    if ((frm.returnmethod.value!="R050")&&(frm.returnmethod.value!="R000")){
        alert('�ܺθ��� ��� ȯ�� ���� �Ǵ� �ܺθ� ȯ���� �����ϼ���. \n\n���� ����ڸ� ���� ���޸����� ��� ȯ�� ó�� �մϴ�.');
        frm.returnmethod.focus();
        return
    }
    <% end if %>



    if (frm.refundrequire.value!=frm.canceltotal.value){
        if ((frm.returnmethod.value!="R007")&&(frm.returnmethod.value!="R900")&&(frm.returnmethod.value!="R000")){
            alert('ȯ�� �ݾװ� ��ұݾ��� �ٸ���� ������ �Ǵ� ���ϸ��� ȯ�Ҹ� �����մϴ�.');
            return false;
        }

        if (!confirm('ȯ�� �ݾ��� ��� �ݾװ� �ٸ��� �����ɰ�� ����ġ�� �ݾ��� �Էµ˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
            return false;
        }
    }

    return true;
}


//����
function CsRegCancelProc(frm){
    if (confirm('��ϵ� ���� ������ ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "deletecsas";
        frm.submit();
    }
}

//�������·� ����
function CsRegStateChg(frm){
    if (confirm('���� ���·� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "state2jupsu";
        frm.submit();
    }
}


//�Ϸ�ó��
function CsRegFinishProc(frm){
    var divcd = frm.divcd.value;

    //ȯ�ҿ�û , �ſ�ī�� ��ҿ�û
    if ((divcd=="A003")||(divcd=="A007")){
        if (frm.contents_finish.value.length<1){
            alert('ó�� ������ �Է��ϼ���.');
            frm.contents_finish.focus();
            return;
        }
    }

    var confirmMsg ;
    confirmMsg = '�Ϸ�ó�� ���� �Ͻðڽ��ϱ�?';

    if ((divcd=="A004")||(divcd=="A010")){
        confirmMsg = '�Ϸ�ó�� ����� ���̳ʽ� �ֹ� �� ȯ���� �ڵ� �����˴ϴ�. ���� �Ͻðڽ��ϱ�?';
    }

    //20090601�߰�
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

    if (confirm('ȯ�� �� ���̳ʽ� ��� ���� �Ϸ�ó�� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "finishcsas";
        frm.modeflag2.value = "norefund";
        frm.submit();
    }
}

//��ǰ ��ü üũ �Ǿ�����
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
                alert('������ �Է��ϼ���.');
                e.focus();
                e.select();
                return false;
            }
            <% else %>
            if ((e.value*1)==0){
                alert('������ �Է��ϼ���.');
                e.focus();
                e.select();
                return false;
            }
            <% end if %>

            regitemno = e.value;
        }

        // ������.
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

    //��ۺ� ----------------------------------------
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

    //��Ÿ����, ���񽺹߼� , ȯ�ҿ�û, �������ǻ���, ��ü �߰� ���� - �󼼳��� üũ ����.
    if ((Fdivcd=="A009")||(Fdivcd=="A002")||(Fdivcd=="A003")||(Fdivcd=="A005")||(Fdivcd=="A006")||(Fdivcd=="A700")){
        // no- check

    }else{
        if (!checkitemExists){
            alert('���õ� �󼼳����� �����ϴ�.');
            return false;
        }
    }

    return true;
}

//��ü �߰� ���� ������
function clearAddUpchejungsan(frm){
    frm.add_upchejungsandeliverypay.value = "0";
    frm.add_upchejungsancause.value = "";
    frm.add_upchejungsancauseText.value = "";
}

function setTitle(frm,titlestr){
    frm.title.value=titlestr;
}

function getReCalcuBeasongPay(userlevel, itemsum){
    //Yellow : 5�����̻�, �׸�ȸ�� : 4�����̻�, ���ȸ�� : 3�����̻�, vipȸ�� : �׻� ���� , ���Ͼ� 2�����̻�

    //�������ΰ�� üũ. �߰� �ؾ���

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

    //��ü��ۺ� ȯ�� �߰�
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

    // ���� ��ǰ �հ� �ݾ�
    if (frm.orgitemcostsum!=undefined){
        orgitemcostsum = frm.orgitemcostsum.value*1;
    }

    // ��ǰ��� �ϴ� ���� �հ�
    if (frm.itemcanceltotal!=undefined){
         frm.itemcanceltotal.value = refunditemcostsum;
    }

// ��� �� ���� �ÿ��� ����.
<% if (IsRegState) or (orefund.FResultCount>0)  then %>
    // ���/��ǰ ��ǰ�Ѿ�
    if (frm.refunditemcostsum!=undefined){
        frm.refunditemcostsum.value = refunditemcostsum;
    }

    // ���þ���(������) ��ǰ�Ѿ�
    if (frm.remainitemcostsum!=undefined){
        frm.remainitemcostsum.value = orgitemcostsum - refunditemcostsum;

        remainitemcostsum = frm.remainitemcostsum.value;
    }

    // ��� ���ϸ��� ȯ��
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

    // ��� ���α� ȯ��
    if (frm.couponreturn!=undefined){
        if (frm.couponreturn.checked){
            frm.refundcouponsum.value = frm.tencardspend.value*-1;
        }else{
            frm.refundcouponsum.value = 0;

            // % ���α� ����.
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

    // ī�� ���� ����
    if (frm.allatsubtractsum){
//    if (frm.allatsubtract!=undefined){
//        if (frm.allatsubtract.checked){
//            frm.allatsubtractsum.value = allatitemdiscountSum*-1;
//        }else{
            frm.allatsubtractsum.value = 0;

            // ī�� ���� ���� 200906�߰�
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

    // ��Ÿ�����ݾ�
    if (frm.refundadjustpay!=undefined){
        refundadjustpay = frm.refundadjustpay.value*1;
    }

    //�� ��ۺ�
    if (frm.orgbeasongpay!=undefined){
        orgbeasongpay  = frm.orgbeasongpay.value*1;
    }

    // ������μ����� �� ��ۺ� ó��
    if (IsCancelProcess){

        //����.. ��ü��� ���� �ΰ��. : ����ۺ� 0�ΰ�� ����.
        //recalcubeasongpay = getReCalcuBeasongPay('<%= oordermaster.FOneItem.FUserLevel %>',remainitemcostsum*1);

        //if (frm.recalcubeasongpay!=undefined){
        //    frm.recalcubeasongpay.value = recalcubeasongpay;
        //}

        //��ҽ� ��ۺ� ȯ�޾�
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

    // ��ǰ���μ����� �� ��ۺ� ó��
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

    //��� ��ۺ� = refundbeasongpay - recalcubeasongpay
    if (frm.refundbeasongpay!=undefined){
        frm.refundbeasongpay.value = refundbeasongpay
    }

    // ȸ�� ��ۺ� -
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

    //��ұݾ� �հ�
    if (frm.canceltotal!=undefined){
        frm.canceltotal.value  = refunditemcostsum + refundmileagesum*1 + refundcouponsum*1 + allatsubtractsum*1 + refundbeasongpay*1 + refundadjustpay*1 + refunddeliverypay*1;
    }

    //����� �ݾ� �հ�
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

//�ҷ���ǰ���
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

//�߰������ۺ�
function Change_add_upchejungsandeliverypay(comp){
    comp.form.buf_totupchejungsandeliverypay.value = comp.form.buf_refunddeliverypay.value*1 + comp.value*1;

    if (isNaN(comp.form.buf_totupchejungsandeliverypay.value)){
        comp.form.buf_totupchejungsandeliverypay.value = "0";
    }
}

//�߰������ۺ� ����
function Change_add_upchejungsancause(comp){
    if (comp.value=="�����Է�") {
        document.all.span_add_upchejungsancauseText.style.display = "inline";
    }else{
        document.all.span_add_upchejungsancauseText.style.display = "none";
    }
}

// ������� �̷�
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
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>CSó�� ��û ���</b></td>
            <td width="140" align="right" <%= ChkIIF(ExistsRegedCSCount>1,"bgcolor='#33CC33'","") %> >
            <% if (ExistsRegedCSCount>1) then %>
                <a href="javascript:ShowOLDCSList();">�� ������ CS �� (<%= ExistsRegedCSCount-1 %>)</a>
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
                    <font color="red">����</font>
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
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">��������</td>
            <td bgcolor="#FFFFFF">
                <% if (IsRegState) then %>
		    	<% call drawSelectBoxCSCommCombo("divcd",divcd,"Z001","onChange='reloadMe(this);'") %>
		    	<% else %>
		    	<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
		    	<font style='line-height:100%; font-size:15px; color:blue; font-family:����; font-weight:bold'><%= ocsaslist.FOneItem.GetAsDivCDName %></font>
		    	&nbsp;
		    	<font style='line-height:100%; font-size:15px; color:#CC3333; font-family:����; font-weight:bold'>[<%= ocsaslist.FOneItem.GetCurrstateName %>]</font>

		    	<% if ocsaslist.FOneITem.FDeleteyn<>"N" then %>
		    	 <font style='line-height:100%; font-size:15px; color:#FF0000; font-family:����; font-weight:bold'>- ������ ����</font>
		    	<% end if %>

		    	<% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" width="80" align="center">�ֹ���ȣ</td>
            <td bgcolor="#FFFFFF" width="200" >
                <%= orderserial %>
                [<font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %>]
                [<font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %></font>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                    <%= session("ssbctid") %>
                <% else %>
                    <%= ocsaslist.FOneItem.Fwriteuser %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ���ID</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FUserID %>(<font color="<%= oordermaster.FOneItem.GetUserLevelColor %>"><%= oordermaster.FOneItem.GetUserLevelName %></font>)
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����Ͻ�</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                <%= now() %>
                <% else %>
                <%= ocsaslist.FOneItem.Fregdate %>
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�ֹ�������</td>
            <td bgcolor="#FFFFFF">
                <%= oordermaster.FOneItem.FBuyname %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FBuyHp %>]
            </td>
        </tr>
        <tr height="20">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF" >
                <% if (IsRegState) then %>
                <input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= GetDefaultTitle(divcd, id, orderserial) %>" size="56" maxlength="56">
                <% else %>
                <input <% if IsFinishProcState then response.write "class='text_ro' ReadOnly" else response.write "class='text'" end if %> type="text" name="title" value="<%= ocsaslist.FOneItem.Ftitle %>" size="56" maxlength="56">
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">����������</td>
            <td bgcolor="#FFFFFF">
                 <%= oordermaster.FOneItem.FReqName %>
                 &nbsp;
                 [<%= oordermaster.FOneItem.FReqHp %>]
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
                <input type="hidden" name="gubun01" value="<%= ocsaslist.FOneItem.Fgubun01 %>">
                <input type="hidden" name="gubun02" value="<%= ocsaslist.FOneItem.Fgubun02 %>">


                <input class="text_ro" type="text" name="gubun01name" value="<%= ocsaslist.FOneItem.Fgubun01name %>" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="<%= ocsaslist.FOneItem.Fgubun02name %>" size="16" Readonly >
                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frmaction.gubun01.value, frmaction.gubun02.value, frmaction.gubun01.name, frmaction.gubun02.name, frmaction.gubun01name.name, frmaction.gubun02name.name,'frmaction','causepop');">

                <div id="causepop" style="position:absolute;"></div>

                <!-- Quick Menu -->

                <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">ǰ��</a>]
                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (ocsaslist.FOneItem.IsReturnProcess) then %>
                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�ܼ�����</a>]
                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
                [<a href="javascript:selectGubun('C006','CF01','��۰���','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">�����</a>]

                <% elseif (divcd="A009") or (divcd="A006") or (divcd="A700") or (divcd="A900") then %>
                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��Ÿ</a>]

                <% elseif (divcd="A001") then %>
                [<a href="javascript:selectGubun('C006','CF03','��۰���','���Ż�ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ����</a>]

                <% elseif (divcd="A002") then %>
                [<a href="javascript:selectGubun('C006','CF04','��۰���','����ǰ����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(����)����ǰ����</a>]
                [<a href="javascript:selectGubun('C005','CE05','��ǰ����','�̺�Ʈ�����','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">(MD)�̺�Ʈ�����</a>]

                <% elseif (divcd="A000") then %>
                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ҷ�</a>]
                [<a href="javascript:selectGubun('C006','CF01','��۰���','���߼�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">���߼�</a>]
                [<a href="javascript:selectGubun('C006','CF02','��۰���','��ǰ�ļ�','gubun01','gubun02','gubun01name','gubun02name','frmaction','causepop');">��ǰ�ļ�</a>]
                <% end if %>
            </td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">��������</td>
            <td bgcolor="#FFFFFF">
            	<% if oordermaster.FOneItem.IsErrSubtotalPrice then %>
            		<font color="red"><%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��</font>
            	<% else %>
            		<%= FormatNumber(oordermaster.FOneItem.Fsubtotalprice,0) %>��
				<% end if %>
            	&nbsp;
                [<%= oordermaster.FOneItem.JumunMethodName %>]

                <% if (oordermaster.FOneItem.Faccountdiv="110") then %>
                (OK Cashbag��� : <strong><%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %></strong> ��)
                <% end if %>
            </td>
        </tr>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" rowspan="2">��������</td>
            <td bgcolor="#FFFFFF" rowspan="2"><textarea <% if IsFinishProcState then response.write "class='textarea_ro' ReadOnly" else response.write "class='textarea'" end if %> name="contents_jupsu" cols="68" rows="6"><%= ocsaslist.FOneItem.Fcontents_jupsu %></textarea></td>
            <td bgcolor="<%= adminColor("topbar") %>" align="center">���������</td>
            <td bgcolor="#FFFFFF" valign="top">
            	[<%= oordermaster.FOneItem.FReqZipCode %>]<br>
                <%= oordermaster.FOneItem.FReqZipAddr %><br>
                <%= oordermaster.FOneItem.FReqAddress %>
            </td>
        </tr>

        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����ù�����</td>
            <td bgcolor="#FFFFFF" valign="top">
            	<!-- �ڵ� Ȯ���Ұ� -->
            	<% if ocsaslist.FOneItem.IsRequireSongjangNO then %>
			        <% Call drawSelectBoxDeliverCompany ("songjangdiv",ocsaslist.FOneItem.Fsongjangdiv) %>
			        <input type="text" class="text" name="songjangno" value="<%= ocsaslist.FOneItem.Fsongjangno %>" size="14" maxlength="16">
			        <% dim ifindurl : ifindurl = fnGetSongjangURL(ocsaslist.FOneItem.Fsongjangdiv) %>
			        <% if (ocsaslist.FOneItem.Fsongjangdiv="24") then %>
                		<a href="javascript:popDeliveryTrace('<%= ifindurl %>','<%= ocsaslist.FOneItem.Fsongjangno %>');">����</a>
                	<% else %>
			            <a href="<%= ifindurl + ocsaslist.FOneItem.Fsongjangno %>" target="_blank">����</a>
			        <% end if %>
			        <input type="button" class="button" value="����" onClick="changeSongjang('<%= id %>');">
		        <% end if %>
            </td>

        </tr>

        <% if (IsFinishProcState) or (IsUpcheConfirmState) or (IsStateFinished) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">ó������</td>
            <td bgcolor="#FFFFFF">
            <% if (IsUpcheConfirmState) then %>
            <textarea class='textarea_ro' readOnly name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% else %>
            <textarea class='textarea' name="contents_finish" cols="68" rows="7"><%= ocsaslist.FOneItem.Fcontents_finish %></textarea>
            <% end if %>
            </td>
            <td bgcolor="<%= adminColor("pink") %>" align="center">ó������<br>������<br>�����Է�</td>
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
            <td bgcolor="<%= adminColor("topbar") %>" align="center">�����ֹ�</td>
            <td colspan="3"  bgcolor="#FFFFFF">
                <table width="100%" border="0" cellspacing="1" cellpadding="2"  bgcolor="<%= adminColor("tablebg") %>" class="a">
                <tr bgcolor="<%= adminColor("topbar") %>">
                    <td width="100">�ֹ���ȣ</td>
                    <td width="80">��һ���</td>
                    <td width="100">�������</td>
                    <td width="80">�ֹ�����</td>
                    <td width="80">�����Ѿ�</td>
                    <td width="80">�ֹ��Ѿ�</td>
                    <td width="80">����</td>
                    <td width="80">���ϸ���</td>
                    <td width="80">��Ÿī������</td>
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
                    <td >�ֹ���ID</td>
                    <td >�ֹ��ڸ�</td>
                    <td >�ֹ���Hp</td>
                    <td >������</td>
                    <td >������Hp</td>
                    <td colspan="4">������ּ�</td>
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
    <!-- ��ǰ �� ������ �ʿ��� ��� -->
    <% if (IsItemDetailDisplay) then %>
        <% if (ocsOrderDetail.FResultCount>0) then %>
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center">������ǰ</td>
            <td colspan="3" bgcolor="#FFFFFF">
                <!-- ��ǰ �� ���� -->
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
                        <tr height="20" align="center" bgcolor="#F4F4F4">
                          <td width="30">����</td>
                          <td width="50">�̹���</td>
                          <td width="30">����</td>
                          <td width="50">������</td>
                          <td width="50">��ǰ�ڵ�</td>
                          <td width="90">�귣��ID</td>
                          <td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
                          <td width="80">
                          <% if (ocsaslist.FOneItem.IsCancelProcess) then %>
                          ���/���ֹ�
                          <% else %>
                          ����/���ֹ�
                          <% end if %>
                          </td>
                          <td width="60">�ǸŰ���</td>
                          <td width="130">��������</td>
                    	</tr>

            <% for i=0 to ocsOrderDetail.FResultCount-1 %>
                <% isAllchecked = true %>
                <% if (ocsOrderDetail.FItemList(i).Fitemid=0) then %>
                <%
                        baesongmethodstr = oordermaster.BeasongCD2Name(ocsOrderDetail.FItemList(i).Fitemoption)
                        ''�� ��ۺ� = ��ۺ� Total
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
                            <td>��ۺ�</td>
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
                        	    	<font color="red"><%= ocsOrderDetail.FItemList(i).Fitemid %><br>(��ü)</font>
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
                        	        <% ''��ǰ����/���� ��� �̸� ���� �������� %>
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
                        	<!-- ����ī�� ������������ ������ -->
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
                           	    <!-- %���� or All@���� : ��ǰ�� ��밪. -->
                           	    <br>(<%= FormatNumber(ocsOrderDetail.FItemList(i).FdiscountAssingedCost,0) %>)
                           	    <% end if %>
                           	</td>
                           	<% end if %>
                            <td align="center">
                                <input class="input_01" type="text" name="gubun01name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun01name %>" size="7" Readonly >
                                &gt;
                                <input class="input_01" type="text" name="gubun02name_<%= distinctid %>" value="<%= ocsOrderDetail.FItemList(i).Fgubun02name %>" size="7" Readonly >

                                <% if (IsStateFinished) and ((divcd="A010") or (divcd="A011")) and ((ocsOrderDetail.FItemList(i).Fgubun02="CE01") or (ocsOrderDetail.FItemList(i).Fgubun02="CF02")) then %>
                                <br><input type="button" class="button" value="�ҷ����" onClick="popBadItemReg('10<%= CHKIIF(ocsOrderDetail.FItemList(i).FItemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).FItemid),Format00(6,ocsOrderDetail.FItemList(i).FItemid)) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsRegState) %>');">
                                <% elseif (IsRegState) or (Not IsNULL(ocsOrderDetail.FItemList(i).Fid)) then %>
                                <a href="javascript:divCsAsGubunSelect(frmaction.gubun01_<%= distinctid %>.value, frmaction.gubun02_<%= distinctid %>.value, frmaction.gubun01_<%= distinctid %>.name, frmaction.gubun02_<%= distinctid %>.name, frmaction.gubun01name_<%= distinctid %>.name,frmaction.gubun02name_<%= distinctid %>.name,'frmaction','causepop_<%= distinctid %>')"><div id='causestring_<%= distinctid %>' >����ϱ�</div></a>
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
            	    <td>��ǰ�հ�ݾ�</td>
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
            	            <td>���û�ǰ�հ�</td>
            	            <td align="right"><input type="text" name="itemcanceltotal" size="7" readonly style="text-align:right;border: 1px solid #333333;" ></td>
            	        </tr>
            	        </table>
            	    </td>
            	    <td>
            	    </td>
            	</tr>
            </table>
            <!-- ��ǰ �� �� -->
            </td>
           </tr>
        <% end if %>

    <% end if %>
        </table>
    </td>
</tr>

</table>

<!-- ȯ�� ���μ����� �ʿ��� ��� -->
<% if (IsReFundInfoDisplay) or (IsCancelInfoDisplay) or (IsUpCheAddJungsanDisplay) then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0"  class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
    <td bgcolor="#FFFFFF" width="500" valign="top">
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
        <tr height="25">
            <td colspan="5" bgcolor="<%= adminColor("topbar") %>">
            	<img src="/images/icon_star.gif" align="absbottom">
            	&nbsp;<b>��Ұ��� ����</b>
            </td>
        </tr>
        <% if (IsCancelInfoDisplay) then %>
            <% if (orefund.FResultCount>0) then %>
            <tr bgcolor="FFFFFF" align="center" height="23">
                <td></td>
                <td>����</td>
                <td>�� ����</td>
                <td>���/��ǰ</td>
                <td>���/��ǰ ��</td>
            </tr>
            <% if (IsItemDetailDisplay) and (IsEditState) and (orefund.FOneItem.Frefunditemcostsum<>regitemcostsum) and (regitemcostsum<>0) then %>
            <script language='javascript'>alert('���� �ݾ� ����ġ-������ ���� ���');</script>
            <% end if %>
            <tr bgcolor="FFFFFF">
        		<td>��ǰ�Ѿ�</td>
        		<td width="80"></td>
        		<td align="right" width="70"><%= FormatNumber(orefund.FOneItem.Forgitemcostsum,0) %></td>
        		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="<%= orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
        	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="<%= orefund.FOneItem.Forgitemcostsum-orefund.FOneItem.Frefunditemcostsum %>" size="9" style="text-align:right" readonly></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>�ֹ��� ��ۺ�</td>
        		<td><div id="beasongpayAssign" ><input <% if (IsFinishProcState) then response.write "disabled" %> type="checkbox" name="ckbeasongpayAssign" <% if (ckAll<>"") or (orefund.FOneItem.Frefundbeasongpay>0) then response.write "checked" %> value="" onclick="CalculateAndApplyItemCostSum(frmaction);"><font color="red">ȯ��</font></div></td>
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
        		<td>ȸ�� ��ۺ�</td>
        		<td>
	        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)" <% if (orefund.FOneItem.Frefunddeliverypay=-4000) then response.write "checked" %> >
	        		-4000�� ����
	        		<!-- ���� ��� ��ۺ� �������� ���� -->
	        		<br>
	        		<input <% if (IsFinishProcState) then response.write "disabled" %>  type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)"  <% if (orefund.FOneItem.Frefunddeliverypay=-2000) then response.write "checked" %> >
	        		-2000�� ����
        		</td>
        		<td></td>
        		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay %>" size="9" style="text-align:right" style="text-align:right" ></td>
        	    <td></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>��� ���ϸ��� </td>
        		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="milereturn" <% if ((orefund.FOneItem.Forgmileagesum>0) and (orefund.FOneItem.Forgmileagesum+orefund.FOneItem.Frefundmileagesum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgmileagesum *-1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundmileagesum" value="<%= orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text"" name="remainmileagesum" value="<%= orefund.FOneItem.Forgmileagesum*-1-orefund.FOneItem.Frefundmileagesum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>��� ���α�</td>
        		<td><input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="couponreturn" <% if ((orefund.FOneItem.Forgcouponsum>0) and (orefund.FOneItem.Forgcouponsum+orefund.FOneItem.Frefundcouponsum=0)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Forgcouponsum * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundcouponsum" value="<%= orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remaincouponsum" value="<%= orefund.FOneItem.Forgcouponsum*-1 -orefund.FOneItem.Frefundcouponsum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>ī�� ���αݾ�</td>
        		<td><!-- input type="checkbox" <% if (IsFinishProcState) then response.write "disabled" %> name="allatsubtract" <% if ((orefund.FOneItem.Fallatsubtractsum>0)  ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->��������</td>
        		<td align="right"><%= FormatNumber(orefund.FOneItem.Fallatsubtractsum * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="allatsubtractsum" value="<%= orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">

        		    <input class="text_ro" type="text" name="remainallatdiscount" value="<%= orefund.FOneItem.Forgallatdiscountsum*-1 - orefund.FOneItem.Fallatsubtractsum %>" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>��Ÿ�����ݾ�</td>
        		<td></td>
        		<td align="right"></td>
        		<td align="right"><input class="text" type="text" name="refundadjustpay" value="<%= orefund.FoneItem.Frefundadjustpay %>" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
                <td align="right"></td>
        	</tr>
        	<tr bgcolor="FFFFFF">
                <td>�Ѿ�/��Ҿ�</td>
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
        		<td>��ǰ�Ѿ�</td>
        		<td width="120"></td>
        		<td align="right" width="70"><%= FormatNumber(orgitemcostsum,0) %></td>
        		<td align="right" width="80"><input class="text_ro" type="text" name="refunditemcostsum" value="0" size="9" style="text-align:right" readonly></td>
        	    <td align="right" width="80"><input class="text_ro" type="text" name="remainitemcostsum" value="0" size="9" style="text-align:right" readonly></td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>�ֹ��� ��ۺ�</td>
        		<td><div id="beasongpayAssign" ><input type="checkbox" name="ckbeasongpayAssign" <% if (ckAll<>"") then response.write "checked" %>  value="" onclick="CalculateAndApplyItemCostSum(frmaction);"><font color="red">��ۺ���ü ȯ��</font></div></td>
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


        	<!-- ��ǰ/ ȸ�� ���μ��� -->
        	<% if (ocsaslist.FOneItem.IsReturnProcess) then %>
        	<tr bgcolor="FFFFFF">
        		<td>ȸ�� ��ۺ�</td>
        		<td>
        			<input type="checkbox" name="ckreturnpay" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
        			-4000�� ����
            		<br>
            		<input type="checkbox" name="ckreturnpayHalf" onClick="CheckDoubleCheck(frmaction,this);CalculateAndApplyItemCostSum(frmaction)">
            		-2000�� ����
        		</td>
        		<td></td>
        		<td align="right"><input class="text_ro" type="text" name="refunddeliverypay" value="0" size="9" style="text-align:right" style="text-align:right" readonly></td>
        	    <td></td>
        	</tr>
        	<% end if %>

        	<% if (ocsaslist.FOneItem.IsCancelProcess) or (ocsaslist.FOneItem.IsReturnProcess) then %>
        	<tr bgcolor="FFFFFF">
        		<td>��� ���ϸ���</td>
        		<td><input type="checkbox" name="milereturn" <% if ((oordermaster.FOneItem.FMileTotalPrice>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FMileTotalPrice * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundmileagesum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remainmileagesum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>��� ���α�</td>
        		<td><input type="checkbox" name="couponreturn" <% if ((oordermaster.FOneItem.FTenCardSpend>0) and (ocsaslist.FOneItem.IsCancelProcess) and (isAllchecked)) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)">ȯ��</td>
        		<td align="right"><%= FormatNumber(oordermaster.FOneItem.FTenCardSpend * -1,0) %></td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="refundcouponsum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        		<td align="right">
        		    <input class="text_ro" type="text" name="remaincouponsum" value="0" size="9" style="text-align:right" readonly>
        		</td>
        	</tr>

        	<tr bgcolor="FFFFFF">
        		<td>ī�� ����</td>
        		<td><!-- input type="checkbox" name="allatsubtract" <% if ((oordermaster.FOneItem.Fallatdiscountprice>0) and (ocsaslist.FOneItem.IsCancelProcess) ) then response.write "checked" %> onClick="CalculateAndApplyItemCostSum(frmaction)" -->����</td>
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
    		<td>��Ÿ�����ݾ�</td>
    		<td></td>
    		<td align="right"></td>
    		<td align="right"><input class="text" type="text" name="refundadjustpay" value="0" size="9" style="text-align:right" onBlur="CalculateAndApplyItemCostSum(frmaction);"></td>
            <td align="right"></td>
    	</tr>
    	<tr bgcolor="FFFFFF">
            <td>�Ѿ�/��Ҿ�</td>
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
        <% if (divcd<>"A700") then ''��ü ��Ÿ����  %>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	        <tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>ȯ�Ұ��� ����</b>
	            </td>
	        </tr>
	        <% if (IsReFundInfoDisplay) then %>
	        <tr bgcolor="#FFFFFF">
	            <td width="100">ȯ�ҹ��</td>
	            <td>
	                <% call drawSelectBoxCancelTypeBox("returnmethod",orefund.FOneItem.Freturnmethod,oordermaster.FOneItem.Faccountdiv,divcd,"onChange='ChangeReturnMethod(this);'") %>
	                <% if (Not IsRegState) then %>
	                (<%= orefund.FOneItem.FreturnmethodName %>)
	                <% end if %>
	                <input name="RefundRecalcuButton" class="csbutton" type="button" value="����" onClick="CalculateAndApplyItemCostSum(frmaction);">
	            </td>
	        </tr>
	        <tr  bgcolor="FFFFFF" id="refundinfo_R007" <% if orefund.FOneItem.Freturnmethod="R007" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	            <td width="100">��������</td>
	            <td align="left">
	                <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="BABABA">
		            	<tr bgcolor="FFFFFF">
		            		<td width="80">���¹�ȣ</td>
		            		<td>
		            		    <input class="text" type="text" size="20" name="rebankaccount" value="<%= orefund.FOneItem.Frebankaccount %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %> >
		            		    <input class="csbutton" type="button" value="��������" onClick="popPreReturnAcct('<%= oordermaster.FOneItem.Fuserid %>','frmaction','rebankaccount','rebankownername','rebankname');" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>>
		            		</td>
		            	</tr>
		            	<tr bgcolor="FFFFFF">
		            		<td>�����ָ�</td>
		            		<td><input class="text" type="text" size="20" name="rebankownername" value="<%= orefund.FOneItem.Frebankownername %>" <%= CHKIIF(IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""),"","disabled") %>></td>
		            	</tr>
		                <tr bgcolor="FFFFFF">
		            		<td>�ŷ�����</td>
		            		<td><% DrawBankCombo "rebankname", orefund.FOneItem.Frebankname %></td>
		            	</tr>
	            	</table>
	            </td>

	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R100" <% if orefund.FOneItem.Freturnmethod="R100" then response.write "style='display:block'" else response.write "style='display:none'" %>>
	    		<td width="100">PG�� ID</td>
	    		<td><input class="text_ro" type="text" name="paygateTid" size="30" value="<%= oordermaster.FOneItem.Fpaygatetid %>" readonly></td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R050" style="display:none">
	            <td colspan="2" align="left">�ܺθ� ȯ�ҿ�û</td>
	        </tr>
	        <tr bgcolor="FFFFFF" id="refundinfo_R900" style="display:none">
	    		<td width="100">���̵�</td>
	    		<td><input class="text_ro" type="text" name="refundbymile_userid" value="<%= oordermaster.FOneItem.Fuserid %>" readonly></td>
	        </tr>
	        <tr bgcolor="FFFFFF">
	    		<td width="100">ȯ�� ������</td>
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
	    	    <td colspan="2"><b>ȯ�� ���� �ۼ����̹Ƿ� ���� �� �� �����ϴ�.</b> [<%= orefund.FOneItem.Fupfiledate %>]</td>
	    	</tr>
	        <% end if %>

	    	<% if (orefund.FResultCount>0) then %>
	    	<tr bgcolor="FFFFFF">
	    	    <td colspan="2"><b>(ȯ�ҿ����� ������ ��Ÿ�����ݾ����� ������ �Էµ˴ϴ�.)</b></td>
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
	        	        alert('���ϸ��� ȯ���� �Ϸ�ó���� �ڵ� ȯ�� �˴ϴ�.');
	        	    }

	        	    if ((Fdivcd=="A003")&&(frmaction.returnmethod.value=="R007")){
	        	        alert('������ ȯ�� �Ϸ�ó���� ���ڸ޼����� �߼��� �ּ���.');
	        	    }
	        	    </script>
	        	<% end if %>
	    	<% else %>
	        <tr bgcolor="FFFFFF" ><td align="center">ȯ�� ���� �Ұ� �Ǵ� ���� ���� ���� </td></tr>
	        <% end if %>
        </table>
        <% end if %>

        <p>

        <% if (IsUpCheAddJungsanDisplay) then %>
    	<!-- ��ü ��ǰ�ΰ�� -->
    	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    		<tr height="25">
	            <td colspan="2" bgcolor="<%= adminColor("topbar") %>">
	            	<img src="/images/icon_star.gif" align="absbottom">
	            	&nbsp;<b>��ü �߰� ���� ����</b>
	            </td>
	        </tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�귣��ID</td>
	    	    <td ><input type="text" class="text_ro" name="buf_requiremakerid" value="<%= ocsaslist.FOneItem.Fmakerid %>" size="20" ReadOnly >
	    	    <% if (divcd="A700") then %>
	    	    <input type="button" class="button" value="�귣��ID�˻�" onclick="jsSearchBrandID(this.form.name,'buf_requiremakerid');" >
	    	    <% end if %>
	    	    </td>
	    	</tr>

	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">ȸ����ۺ�</td>
	    	    <td ><input type="text" class="text_ro" name="buf_refunddeliverypay" value="<%= orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�߰������ۺ�</td>
	    	    <td ><input type="text" class="text" name="add_upchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay %>" size="10" onKeyUp="Change_add_upchejungsandeliverypay(this);">��
	    	    &nbsp;
	    	    <select class="select" name="add_upchejungsancause" class="text" onChange='Change_add_upchejungsancause(this);'>
		    	    <option value="" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="","selected","") %>>��������
		    	    <option value="�߰���ۺ�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰���ۺ�","selected","") %> >�߰���ۺ�
		    	    <option value="�߰�����" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause="�߰�����","selected","") %>>�߰�����
		    	    <option value="�����Է�" <%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","selected","") %>>�����Է�
	    	    </select>

	    	    <span name="span_add_upchejungsancauseText" id="span_add_upchejungsancauseText" style='display:<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"","inline","none") %>'><input type="text" name="add_upchejungsancauseText" value="<%= ChkIIF(ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰���ۺ�" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"�߰�����" and ocsaslist.FOneItem.Fadd_upchejungsancause<>"",ocsaslist.FOneItem.Fadd_upchejungsancause,"") %>" size="10" maxlength="16" ></span>
	    	    <a href="javascript:clearAddUpchejungsan(frmaction);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	    	    </td>
	    	</tr>
	    	<tr bgcolor="FFFFFF">
	    	    <td width="100">�������ۺ�</td>
	    	    <td ><input type="text" class="text_ro" name="buf_totupchejungsandeliverypay" value="<%= ocsaslist.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1 %>" size="10" ReadOnly >��</td>
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
            <input type="checkbox" name="ForceReturnByTen"><font color="red">��ü��� ��ǰ�̶� �ٹ����� �������ͷ� ȸ���� ��� �̰��� üũ.</font>
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
        <input type="checkbox" name="csmailsend" value="on" > CS ����/ó�� �̸��� �߼�
        <font color=red>(�ʿ��Ѱ�� üũ�ϼ���. �����ϰ� ó������ ���̰� 3�� �ʰ�)</font>
        <% else %>
        <input type="checkbox" name="csmailsend" value="on" <%= chkIIF(oordermaster.FOneItem.FSiteName="10x10","checked","") %> > CS ����/ó�� �̸��� �߼�
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr>
    <td colspan="4" align="center">
    <% if (IsRegState) then %>
        <% if (IsJupsuProcessAvail) then %>
        <input class="csbutton" type="button" value=" �� �� " onClick="CsRegProc(frmaction)">
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
                    <input class="csbutton" type="button" value=" �Ϸ� ó�� (���̳ʽ�/ȯ�ҿ�û ���)" onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
                    <input class="csbutton" type="button" value=" [���̳ʽ�/ȯ�ҿ�û ����] �Ϸ� ó�� " onClick="CsRegFinishProcNoRefund(frmaction)" onFocus="blur()">
                <% else %>
                    <input class="csbutton" type="button" value=" �Ϸ� ó�� " onClick="CsRegFinishProc(frmaction)" onFocus="blur()">
                <% end if %>
            <% else %>
                <% IF (Not (IsNULL(orefund.FOneItem.Fupfiledate) or (orefund.FOneItem.Fupfiledate=""))) then %>
                ȯ������ �ۼ����̹Ƿ� ���� �Ұ� �մϴ�.
                <% else %>
                <input class="csbutton" type="button" value=" ���� ��� " onClick="CsRegCancelProc(frmaction)" onFocus="blur()">
                <input class="csbutton" type="button" value=" �������� ���� " onClick="CsRegEditProc(frmaction)" onFocus="blur()">
                    <% if (IsUpcheConfirmState) then %>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input class="csbutton" type="button" value=" �������·� ���� " onClick="CsUpcheConfirm2RegProc(frmaction)" onFocus="blur()">
                    <% end if %>
                <% end if %>
            <% end if %>


        <% else %>

        <% end if %>
    <% elseif (IsStateFinished) then %>
        <% if (divcd="A700") and (mode<>"finishreginfo") then %>
        <!--
            <input class="csbutton" type="button" value=" ���� ���·� ���� " onClick="CsRegStateChg(frmaction)" onFocus="blur()">
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
    alert('�̰����� �Ϸ�ó�� �Ͽ��� \n\n\n�ſ�ī�� ������� �� ������ȯ��ó���� �̷�� ���� ������ �����Ͻñ� �ٶ��ϴ�.!\n\n\n\n\n\n ');
<% end if %>
}
window.onload = getOnload;

<% if (ocsaslist.FOneITem.FDeleteyn="Y") then %>
alert('������ �����Դϴ�.');
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
