<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim id
id = requestCheckvar(request("id"),10)

dim makerid, groupid
makerid = session("ssBctId")
groupid = getPartnerId2GroupID(makerid)

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
''ojungsan.FRectdesigner = session("ssBctID")
ojungsan.FRectGroupID = groupid
if (groupid<>"") then
ojungsan.JungsanMasterList
end if

if ojungsan.FResultCount<1 then
	response.write "<script>alert('���������� �����ϴ�.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

''������ ���� �Ұ�
if ojungsan.FItemList(0).IsCommissionTax then
    response.write "<script>alert('������ ��꼭 ���� �Ұ�');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

makerid = ojungsan.FItemList(0).Fdesignerid

dim opartner, ogroup
dim stypename

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid ''session("ssBctId")
opartner.FPageSize = 1
opartner.GetOnePartnerNUser



set ogroup = new CPartnerGroup
ogroup.FRectGroupid = ojungsan.FItemList(0).Fgroupid
ogroup.GetOneGroupInfo

if ogroup.FResultCount<1 then
    response.write "<script>alert('�׷� �ڵ尡 �������� �ʾҰų�, ���������� �����ϴ�.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

if (ojungsan.FItemList(0).IsElecTaxCase) then
	stypename = "���ݰ�꼭"
elseif (ojungsan.FItemList(0).IsElecFreeTaxCase) then
	stypename = "��꼭"
else
	response.write "<script>alert('���ݰ�꼭 Ȥ�� ��꼭�� ���� �����մϴ�. - �̹� ���� �Ͽ��ų� ������ ������ �����ϴ�.');</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
end if

dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3
jungsan_hpall = Trim(ogroup.FOneItem.Fjungsan_hp)
jungsan_hpall = split(jungsan_hpall,"-")

if UBound(jungsan_hpall)>=0 then
	jungsan_hp1 = jungsan_hpall(0)
end if

if UBound(jungsan_hpall)>=1 then
	jungsan_hp2 = jungsan_hpall(1)
end if

if UBound(jungsan_hpall)>=2 then
	jungsan_hp3 = jungsan_hpall(2)
end if



dim Bill365URL : Bill365URL = "http://www.bill36524.com"  '' :8090: test, 80: real
dim swfName    : swfName = "DzEBankFlexAPI" ''"dZAmfApp"

IF application("Svr_Info")="Dev" THEN
    ''�Ƴ�..
    ''Bill365URL = "http://www.bill36524.com:8090"
    ''swfName = "DzEBankFlexAPI_test"
end if

%>



<script language="VBScript">
<!--
Sub <%= swfName %>_FSCommand(ByVal command, ByVal args)
    call <%= swfName %>_DoFSCommand(command, args)
end sub
//-->
</script>

<script src="AC_OETags.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
	AC_FL_RunContent(
    	"src", "<%= swfName %>",
    	"width", "0",
    	"height", "0",
    	"align", "middle",
    	"id", "<%= swfName %>",
    	"quality", "high",
    	"bgcolor", "#869ca7",
    	"name", "<%= swfName %>",
    	"allowScriptAccess","always",
    	"type", "application/x-shockwave-flash",
    	"pluginspage", "http://www.adobe.com/go/getflashplayer"
    );
</script>



<script language='javascript'>
var pLogIdx = 0;
var fxStarted = false;

function getMatchStr(stre,pt){

    var pat = "[<]"+pt+"[>](.*?)[<]\/"+pt+"[>]";

    var re = new RegExp(pat,"g");

    var resultArray = re.exec(stre);

    if (resultArray==null){
        return "";
    }else{
        return (resultArray[1])
    }

}

//���. :: �ٸ�����. dZAmfApp.swf
function <%= swfName %>_DoFSCommand(cmd, args) {
    var result, result_msg, no_tax, no_iss ;

    var compNo = '<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>';

    //alert(args);
    switch (cmd) {
        case "Login" :
            result = getMatchStr(args,"RESULT");
            result_msg = getMatchStr(args,"RESULT_MSG");
            if (result!="00000"){
                alert(result_msg);

            }else{
                var no_id = getMatchStr(args,"NO_ID");
                //����ڹ�ȣ üũ
                if (no_id!=compNo){
                    alert('bill36524 ����Ʈ�� ���Ե� ����ڹ�ȣ�� �ٹ����ٿ� ��ϵ� ����ڹ�ȣ�� ��ġ���� �ʽ��ϴ�.\n\nbill36524�� ��ϵ� ����ڹ�ȣ:' + no_id + '\n�ٹ����ٿ� ��ϵ� ����ڹ�ȣ:' + compNo);
                    return;
                }

                preSaveLog();
            }
            break;
        case "SendTaxAccount" :
            result  = getMatchStr(args,"RESULT");
            result_msg  = getMatchStr(args,"RESULT_MSG");
            no_tax = getMatchStr(args,"NO_TAX");
            no_iss = getMatchStr(args,"NO_ISS"); //����û���ι�ȣ

            saveTaxEvalResult(result,no_tax,result_msg,no_iss);
            if (result!="00000"){
                if (result=="10000"){
                    alert(result + result_msg + "\n\nbill36524.com���� \n�����ȯ�漳�� => ������ ��Ͽ��� ������ ����� ����Ͻñ� �ٶ��ϴ�.");
                }else{
                    alert(result + result_msg);
                }
            }else{
                //popTax
                FxShowTaxAccount(no_tax,compNo);
            }
            break;
        default :
            alert(cmd);
            alert(args);
            break;
    }
}


function thisMovie(movieName){
    if(navigator.appName.indexOf("Microsoft") != -1){
        return window[movieName];
    }else {
        return document[movieName];
    }
}

function AddNew(key, value)
{
 var obj = new Object();
 obj.key = key;
 obj.value = value;
 return obj;
}


//01.�α���
function FxLogin(iid,ipwd){

    if (fxStarted) return;
    fxStarted = true;

    pLogIdx = 0;

    var obj = AddNew("ID", iid);
    var obj1 = AddNew("PASSWD", ipwd);
    var obj2 = AddNew("USER_IP", "<%= request.ServerVariables("REMOTE_ADDR") %>");

    var arr = new Array(obj, obj1, obj2);
    try {
        thisMovie("<%= swfName %>").Login(arr);
    } catch (e) {
        alert('�÷��� ���� �ε� ���� - ���� ���(070-7515-5403 ������)');
	}
    document.all.txtMsg.innerHTML = "bill36524.com �� �α������Դϴ�. ��� ��ٷ��ּ���..";
    //alert('startedlogin');
}




//01.�α��� ���
function FxLoginResult(retObj){
    //alert(retObj);
    var result = retObj.RESULT;
    var company_no = "<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>";

    document.all.txtMsg.innerHTML = "";
    if (result=="00000"){
        //����ڹ�ȣ üũ
        if (retObj.NO_ID!=company_no){
            hideLogin();
            alert('bill36524 ����Ʈ�� ���Ե� ����ڹ�ȣ�� �ٹ����ٿ� ��ϵ� ����ڹ�ȣ�� ��ġ���� �ʽ��ϴ�.\n\nbill36524�� ��ϵ� ����ڹ�ȣ:' + retObj.NO_ID + '\n�ٹ����ٿ� ��ϵ� ����ڹ�ȣ:'+company_no);
            return;
        }

        preSaveLog();
    }else{
        hideLogin();
        alert(retObj.RESULT_MSG);
    }

}

//������ ����
function preSaveLog(){
    var frm = document.frm;
    <% if (jungsan_hp1="") or (jungsan_hp2="") or (jungsan_hp3="") or (Len(jungsan_hp1)>3) or (Len(jungsan_hp2)>4) or (Len(jungsan_hp3)>4) then %>

        alert('���� ����� �ڵ��� ��ȣ�� �ùٸ��� �ʽ��ϴ�. \n��ü������������ �������� �ڵ����� 000-000-0000 ��� ���·� ������ ����ϼ���.');
        hideLogin();
        return;

    <% end if %>


    frm.action="dotaxreg.asp";
	frm.target = "ipreSave";
	frm.submit();

}

//������ ����
function saveTaxEvalResult(result,no_tax,result_msg,no_iss){
    var frm = taxSaveFrm;
    frm.action="saveTaxResult.asp";
    frm.idx.value = pLogIdx;
    frm.result.value = result;
    frm.no_tax.value = no_tax;
    frm.no_iss.value = no_iss;
    frm.result_msg.value = result_msg;

	frm.target = "ipreSave";
	frm.submit();

	fxStarted = false;
}


function billTaxEvalFlexApi(pidx){
    pLogIdx = pidx;
    <%
    dim FG_VAT
    if (ojungsan.FItemList(0).Ftaxtype="03") then
        FG_VAT = "2"
    elseif (ojungsan.FItemList(0).Ftaxtype="02") then
        FG_VAT = "3"
    else
        FG_VAT = "1"
    end if
    %>
    var obj1 = AddNew("FG_BILL","1");   //û��1 ����2
    var obj2 = AddNew("YN_TURN","Y");   //Y������ N������  :: ������� �����û , ������� ���ο�û
    var obj3 = AddNew("FG_IO","1");     //1���� 2����
    var obj4 = AddNew("FG_PC","1");     //1��� 2����
    var obj5 = AddNew("FG_FINAL","1");  //0���� 1 �߼� 2���� 3�ݷ� 4������ҿ�û
    var obj6 = AddNew("YN_CSMT","N");
    var obj7 = AddNew("FG_VAT","<%= FG_VAT %>");    // 1����,2����,3�鼼
    var obj8 = AddNew("AM","<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");
    var obj9 = AddNew("AM_VAT","<%= ojungsan.FItemList(0).GetTotalTaxvat %>");
    var obj10 = AddNew("AMT","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");

    var obj11 = AddNew("AMT_CASH","0");
    var obj12 = AddNew("AMT_CHECK","0");
    var obj13 = AddNew("AMT_NOTE","0");
    var obj14 = AddNew("YMD_WRITE","<%= Replace(ojungsan.FItemList(0).GetPreFixSegumil,"-","") %>");

    var obj15 = AddNew("SELL_NO_BIZ","<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>");
    var obj16 = AddNew("SELL_NM_CORP","<%= ogroup.FOneItem.FCompany_name %>");
    var obj17 = AddNew("SELL_NM_CEO","<%= ogroup.FOneItem.Fceoname %>");
    var obj18 = AddNew("SELL_BIZ_STATUS","<%= ogroup.FOneItem.Fcompany_uptae %>");
    var obj19 = AddNew("SELL_BIZ_TYPE","<%= ogroup.FOneItem.Fcompany_upjong %>");

    var obj20 = AddNew("SELL_ADDR1","<%= ogroup.FOneItem.Fcompany_address %>");
    var obj21 = AddNew("SELL_ADDR2","<%= ogroup.FOneItem.Fcompany_address2 %>");
    var obj22 = AddNew("SELL_DAM_DEPT","");
    var obj23 = AddNew("SELL_DAM_NM","<%= ogroup.FOneItem.Fjungsan_name %>");
    var obj24 = AddNew("SELL_DAM_EMAIL","<%= ogroup.FOneItem.Fjungsan_email %>");

    var obj25 = AddNew("SELL_DAM_MOBIL1","<%= jungsan_hp1 %>");
    var obj26 = AddNew("SELL_DAM_MOBIL2","<%= jungsan_hp2 %>");
    var obj27 = AddNew("SELL_DAM_MOBIL3","<%= jungsan_hp3 %>");

    var obj28 = AddNew("SELL_DAM_TEL1","<%= jungsan_hp1 %>");
    var obj29 = AddNew("SELL_DAM_TEL2","<%= jungsan_hp2 %>");
    var obj30 = AddNew("SELL_DAM_TEL3","<%= jungsan_hp3 %>");


    var obj31 = AddNew("BUY_NO_BIZ","2118700620");
    var obj32 = AddNew("BUY_NM_CEO","������");
    var obj33 = AddNew("BUY_NM_CORP","(��)�ٹ�����");

    var obj34 = AddNew("BUY_DAM_NM","������");
    var obj35 = AddNew("BUY_DAM_EMAIL","accounts@10x10.co.kr");

    var obj36 = AddNew("BUY_DAM_MOBIL1","02");
    var obj37 = AddNew("BUY_DAM_MOBIL2","554");
    var obj38 = AddNew("BUY_DAM_MOBIL3","2033");

    var obj39 = AddNew("BUY_DAM_TEL1","02");
    var obj40 = AddNew("BUY_DAM_TEL2","554");
    var obj41 = AddNew("BUY_DAM_TEL3","2033");

    var obj42 = AddNew("BUY_ADDR1","����� ���α� ������");
    var obj43 = AddNew("BUY_ADDR2","1-45 ��������2��");
    var obj44 = AddNew("BUY_BIZ_STATUS","���Ҹſ�");
    var obj45 = AddNew("BUY_BIZ_TYPE","���ڻ�ŷ���");

    var obj46 = AddNew("BUY_DAM_DEPT","�¶���");

    var obj47 = AddNew("AMT_AR","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");   //�ܻ�̼���
    //var obj48 = AddNew("CD_SVC","<%= ojungsan.FItemList(0).GetTotalSuplycash %>");   //CD_SVC ??
    var obj48 = AddNew("NO_SERIAL",pidx);   //�Ϸù�ȣ

    var obj49 = AddNew("DC_RMK2", "[10x10 scm �������� : ID" + pidx + "]");

    //201002����.
    var obj50 = AddNew("YN_FX","N"); // ���� ���ݰ�꼭 ����  Y:���� ���� ��꼭, N: ���� ���� <== �ʼ� �Է� �Դϴ�
    var obj51 = AddNew("NO_SENDER_PK","DZ_TEN_ON_<%= ojungsan.FRectId %>_<%= ojungsan.FItemList(0).Fdifferencekey %>_<%= ojungsan.FItemList(0).GetTotalSuplycash %>");
    
    //2016/04/18 �߰�
    var obj52 = AddNew("YN_ISS","0");  //FG_VAT �� 3(�鼼) �ϰ�� YN_ISS : NULL �ϰ�� �������� YN_ISS : 0 �ϰ�� ����û ���ۿ�û
    
    <% if (TRUE) or (FG_VAT="3") then %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51, obj52);
    <% else %>
    var arr = new Array(obj1 ,obj2 ,obj3 ,obj4 ,obj5 ,obj6 ,obj7 ,obj8 ,obj9 ,obj10,obj11,obj12,obj13,obj14,obj15,obj16,obj17,obj18,obj19,obj20,obj21,obj22,obj23,obj24,obj25,obj26,obj27,obj28,obj29,obj30,obj31,obj32,obj33,obj34,obj35,obj36,obj37,obj38,obj39,obj40,obj41,obj42,obj43,obj44,obj45, obj46, obj47, obj48, obj49, obj50, obj51);
    <% end if %>

    var objline1 = AddNew("ITEM_STD", "<%= Right(Replace(ojungsan.FItemList(0).Fyyyymm,"-",""),4) %>");
    var objline2 = AddNew("NM_ITEM", "<%= ojungsan.FItemList(0).getBillItemName %>");
    var objline3 = AddNew("NO_ITEM", "1");
    var objline4 = AddNew("AM", "<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");
    var objline5 = AddNew("AM_VAT", "<%= ojungsan.FItemList(0).GetTotalTaxvat %>");
    var objline6 = AddNew("AMT", "<%= ojungsan.FItemList(0).GetTotalSuplycash %>");
    var objline7 = AddNew("DD_WRITE", "<%= Mid(ojungsan.FItemList(0).GetPreFixSegumil,9,2) %>");
    var objline8 = AddNew("MM_WRITE", "<%= Mid(ojungsan.FItemList(0).GetPreFixSegumil,6,2) %>");
    //var objline9 = AddNew("QTY", "1");      //����
    //var objline10 = AddNew("UM", "<%= ojungsan.FItemList(0).GetTotalTaxSuply %>");      //�ܰ�

    var arrline1 = new Array(objline1, objline2,objline3, objline4, objline5, objline6, objline7, objline8);

    var arrlineArr = new Array(arrline1);

    thisMovie("<%= swfName %>").SendTaxMuch(1);

    thisMovie("<%= swfName %>").SendTaxAccount("", arr, arrlineArr);
    //thisMovie("<%= swfName %>").SendTaxAccount("", arr, arrlineArr, null, "");
    document.all.txtMsg.innerHTML = "��꼭 �������Դϴ�. ��� ��ٷ��ּ���..";
}

function closeMe(){
    opener.location.reload();
    window.close();
}

//02.���ݰ�꼭 ���� ���
function FxSendTaxAccountResult(retObj){
    var result = retObj.RESULT;
    var result_msg  = retObj.RESULT_MSG;
    var tb_tax = retObj.OBJ_TBTAX;
    if (tb_tax!=null){
        var no_tax = tb_tax.NO_TAX;
        var no_iss = tb_tax.NO_ISS; //����û���ι�ȣ
    }else{
        var no_tax = "";
    }

    saveTaxEvalResult(result,no_tax,result_msg,no_iss);
    document.all.txtMsg.innerHTML = "";
    hideLogin();


    if (result!="00000"){
        if (result=="10000"){
            if (result_msg=="API ����� ���ݰ�꼭") {
                alert("���� : " + result_msg + "");
            }else{
                alert("���� : " + result_msg + "\n\nbill36524.com �α��� �Ͻ��� \n�����ȯ�漳�� => ������ ��Ͽ��� ������ ����� ����Ͻñ� �ٶ��ϴ�.");
            }
        }else{
            alert(result_msg);
        }
        location.reload();  //��ε� ���ϸ� ���������� �����߻�(�ߺ������� ���µ�)
    }


    /*
    else{
        //popTax :: �������

        //FxShowTaxAccount(no_tax,compNo);

        alert("��꼭�� ���� �Ǿ����ϴ�. \n�ٹ����ٿ��� ������ (����)��°����մϴ�.");
        opener.location.reload();
        window.close();
    }
    */
}

//popupMove ����
var bdown = false;
var x, y;
var sElem;

function mdown(evt)
{
	evt = (evt) ? evt : ((window.event) ? window.event : "");
	sElem = evt.target ? evt.target : evt.srcElement;
	if (evt.stopPropagation)
	{
		evt.stopPropagation();
		evt.preventDefault();
	}
	evt.returnValue  = false;
	evt.cancelBubble = true;

	if(sElem.className == "drag")
	{
		bdown = true;
		x = evt.clientX;
		y = evt.clientY;
	}
}

function mup()
{
	bdown = false;
}

document.onmousemove = function moveimg(event)
{
	event = (event) ? event : ((window.event) ? window.event : "");
	if(bdown)
	{
		var distX = event.clientX - x;
		var distY = event.clientY - y;
		var targetImg = document.getElementById('POPBillLogin');
		targetImg.style.left = (parseInt(targetImg.style.left) + distX) + 'px';
		targetImg.style.top = (parseInt(targetImg.style.top) + distY) + 'px';
		x = event.clientX;
		y = event.clientY;
		return false;
	}
}


function hideLogin(){
    document.all["POPBillLogin"].style.visibility='hidden';
    document.frm.evalButton.disabled=false;
}

function showLogin(){

    var frm = document.billfrm;
    frm.billid.value = '';
    frm.billpass.value = '';
    hideDoing();
    document.all["POPBillLogin"].style.visibility='visible';

    document.frm.evalButton.disabled=true;
    fxStarted = false;
}

function showDoing(){
    var frm = document.billfrm;
    document.all.ievalBtn.style.display='none';
    document.all.idoingMsg.style.display='inline';
    document.all.popcloseId.style.display='none';
    frm.billid.disabled = true;
    frm.billpass.disabled = true;
}

function hideDoing(){
    var frm = document.billfrm;
    document.all.ievalBtn.style.display='inline';
    document.all.idoingMsg.style.display='none';
    document.all.popcloseId.style.display='inline';
    frm.billid.disabled = false;
    frm.billpass.disabled = false;
}

function billTaxEval(frm){

    if (frm.billid.value.length<1){
        alert('Bill36524 ���̵� �Է��ϼ���.');
        frm.billid.focus();
        return;
    }

    if (frm.billpass.value.length<1){
        alert('Bill36524 �н����带 �Է��ϼ���.');
        frm.billpass.focus();
        return;
    }

    showDoing();
    FxLogin(frm.billid.value,frm.billpass.value);
    // ������ �Ϸ�ȴ��� ����.. hideLogin();
}

//05. ���ݰ�꼭 Ȯ��

function FxShowTaxAccount(no_tax, no_biz_no){
    var url = "<%= Bill365URL %>/popupBillTax.jsp?";
    url += "NO_TAX=" + no_tax;
    url += "&NO_BIZ_NO=" + no_biz_no;


    var popwin = window.open(url, "taxwin", "height=700,width=660, menubar=no, location=no, resizeable=no, status=no, scrollbars=no, top=200, left=300");
    popwin.focus();
}

// FLESH ���ο��� ��Ÿ ���� �߻��� ���� ����
/*
//�ѹ� ��������� ����ؼ� ������ : ��꼭 �������̶�� �޼��� 201002����..
function FxErrorResult(retObj) {

    alert("ERR:" + retObj + "\n������ ���� ���.");

    if (pLogIdx!=0){
        var frm = taxSaveFrm;
        frm.action="saveTaxResult.asp";
        frm.idx.value = pLogIdx;
        frm.result.value = "999";
        frm.no_tax.value = "";
        frm.result_msg.value = retObj;

    	frm.target = "ipreSave";
    	frm.submit();

    	fxStarted = false;
	}
	hideLogin();
}
*/

//���ݰ�꼭 ����� ���� ó��:��Ʈ������ �� ó������ ���� ���� �߻�
function DzErrorEvent(faultEvent){
    var errinfo = "";

    errinfo = "faultEvent.message:" + faultEvent.message + "\n";
    errinfo += "faultEvent.errorID:" + faultEvent.errorID + "\n";
    errinfo += "faultEvent.faultCode:" + faultEvent.faultCode + "\n";
    errinfo += "faultEvent.faultDetail:" + faultEvent.faultDetail + "\n";
    errinfo += "faultEvent.faultString:" + faultEvent.faultString + "\n";

    //form1.fxlog.value = errinfo;
    alert("ERR:" + errinfo + "\n������ ���� ���");

    hideLogin();
}


function ActTaxReg(frm){
//alert('�˼��մϴ�. \n bill36524.com ����Ʈ ������ ��Ȱ���� �ʾ� ��� ��꼭 ������ �����մϴ�.');
//return;
	if (frm.biz_no.value.length!=10){
		alert('����� ��� ��ȣ�� �ùٸ��� �ʰų� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.corp_nm.value.length<1){
		alert('����� ���� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.ceo_nm.value.length<1){
		alert('��ǥ�� ���� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.biz_status.value.length<1){
		alert('���°� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.biz_type.value.length<1){
		alert('������ ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.addr.value.length<1){
		alert('����� �ּҰ� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.dam_nm.value.length<1){
		alert('����� ������ ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.email.value.length<1){
		alert('����� �̸����� ��ϵǾ� ���� �ʽ��ϴ�. - ��ü���� ������ ����ϼ���.');
		return;
	}

	if (frm.write_date.value.length<1){
		alert('��꼭 ������ �Է� �� ����ϼ���.');
		return;
	}

    if (!thisMovie("<%= swfName %>")){
        alert('swf ������ �ε� ���� �ʾҽ��ϴ�.');
        return;
    }

    if (frm.billSite[1].checked){
        if (confirm('�˾�â���� bill36524.com ���̵�� �н����带 �Է��Ͻ��� �����Ͻø� �˴ϴ�. ��� �Ͻðڽ��ϱ�?')){
            showLogin();

        }
        return;
    }

    if (confirm('<%= stypename %> �� ���� �Ͻðڽ��ϱ�?')){
	    frm.action="dotaxreg.asp";
	    frm.target = "";
		frm.submit();
	}
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" width="16" height="16" align="absbottom">
        	<strong>���� <%= stypename %> ����</strong>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="dotaxreg.asp">
	<input type=hidden name=jungsanid value="<%= ojungsan.FItemList(0).FId %>">
	<input type=hidden name=jungsanname value="<%= ojungsan.FItemList(0).Ftitle %>">
	<input type=hidden name=jungsangubun value="ON">
	<input type=hidden name=makerid value="<%= makerid %>">

	<input type=hidden name=biz_no value="<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>" >
	<input type=hidden name=corp_nm value="<%= ogroup.FOneItem.FCompany_name %>">
	<input type=hidden name=ceo_nm value="<%= ogroup.FOneItem.Fceoname %>">
	<input type=hidden name=biz_status value="<%= ogroup.FOneItem.Fcompany_uptae %>">
	<input type=hidden name=biz_type value="<%= ogroup.FOneItem.Fcompany_upjong %>">


	<input type=hidden name=addr value="<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>">
	<input type=hidden name=dam_nm value="<%= ogroup.FOneItem.Fjungsan_name %>">
	<input type=hidden name=email value="<%= ogroup.FOneItem.Fjungsan_email %>">
	<input type=hidden name=hp_no1 value="<%= jungsan_hp1 %>">
	<input type=hidden name=hp_no2 value="<%= jungsan_hp2 %>">
	<input type=hidden name=hp_no3 value="<%= jungsan_hp3 %>">

	<input type=hidden name=sb_type value="02"> <!-- ���� 01 ���� 02 -->
	<input type=hidden name=tax_type value="<%= ojungsan.FItemList(0).Ftaxtype %>">
	<input type=hidden name=bill_type value="18"> <!-- ���� 01 û�� 18 -->
	<input type=hidden name=pc_gbn value="C"> <!-- ���� P ��� C -->

	<input type=hidden name=item_count value="1">
	<input type=hidden name=item_nm value="<%= ojungsan.FItemList(0).getBillItemName %>">
	<input type=hidden name=item_qty value="1">
	<input type=hidden name=item_price value="<%= ojungsan.FItemList(0).GetTotalSuplycash %>">
	<input type=hidden name=item_amt value="<%= ojungsan.FItemList(0).GetTotalTaxSuply %>">
	<input type=hidden name=item_vat value="<%= ojungsan.FItemList(0).GetTotalTaxvat %>">
	<input type=hidden name=item_remark value="">

	<input type=hidden name=credit_amt value="<%= ojungsan.FItemList(0).GetTotalTaxSuply + ojungsan.FItemList(0).GetTotalTaxvat %>">

	<input type=hidden name=cur_u_user_no value="261744"> <!-- DEV 1000394, REAL 244730, ON 261744 -->
	<input type=hidden name=cur_dam_nm value="�ֺ���">
	<input type=hidden name=cur_email value="accounts@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="000">
	<input type=hidden name=cur_hp_no2 value="000">
	<input type=hidden name=cur_hp_no3 value="0000">


    <!--
    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
		* 2005�� 3���� �����(������ 3�� 31��)���ʹ� ���� <%= stypename %> ������ ����ϼž� �մϴ�.
		</td>
	</tr>
	-->
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td height="20" colspan="2">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>���� <%= stypename %> ������</strong>
	   		&nbsp;&nbsp;&nbsp;&nbsp;
	   		<a href="http://www.bill36524.com" target="_blank"><font color="blue">>>bill36524 ȸ�������ϱ�</font></a>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			<img src="/images/icon_num01.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>������ �׿���Ʈ���� ������ �ϼ̰�, bill36524�� �̰��� �ϽŰ��</b></font><br>
				&nbsp;&nbsp;1.���ݰ�꼭 ����â���� [bill36524]�� �����Ͻð� �����Ͻø� �˴ϴ�.<br>
				&nbsp;&nbsp;2.bill36524�� �Ǽ������� �Ǿ� �־�� �ϸ�, �������� �־�� �մϴ�.<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<!--
			<font color="red"><b>�׿���Ʈ�� �̿����̰�, ���� bill36524�� �̰��� ���Ͻ� ���</b></font><br>
				&nbsp;&nbsp;1.���ݰ�꼭 ����â���� [�׿���Ʈ]�� �����Ͻð� �����Ͻø� �˴ϴ�.<br>
				&nbsp;&nbsp;2.�̹��޿� ���ؼ� �ѽ������� �ٹ����ٿ��� ��������Ḧ �����մϴ�.<br>
			<img src="/images/icon_num03.gif" width="16" height="16" align="absbottom">
			-->
			<font color="red"><b>�ű�������ü�� ���</b></font><br>
				<!-- &nbsp;&nbsp;1.���� �׿���Ʈ(www.neoport.net)�� �����ȸ������ ȸ�������� �Ұ����մϴ�.<br>
				&nbsp;&nbsp;2.(�׿���Ʈ ���񽺴� 12������ ���񽺰� �ߴܵǾ����ϴ�.)<br>
				-->
				&nbsp;&nbsp;1.bill36524�� ȸ�� ���� �Ŀ� �� 1���� ���� ����â���� [bill36524]�� �����Ͻð� �����Ͻø� �˴ϴ�.

			<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>������� ��ó���</b></font><br>
			[������ ������ ���� �����ϴ�.] �Ǵ� [����û�������� �������] => bill36524�α��� �� ���� ���θ޴��� [�����ȯ�漳��] Ŭ�� ��������� �ǿ� ������ ������ ���<br>
			[����Ʈ�� �����մϴ�.] =>  bill36524�α��� �� ����Ʈ ���� �� ���<br>
			&nbsp;&nbsp;<a href="/designer/jungsan/popTaxHelp.asp" target="_taxHelp"><font color="#0000FF"><strong>[���ݰ�꼭 ������ �ȳ� �ڼ��� ���� ��]</strong></font></a>

		</td>
	</tr>
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td colspan="2" height="20" valign="middle">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>��ϵ� ��������� Ȯ��</strong>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF" width="30%">����ڸ�</td>
		<td><%= ogroup.FOneItem.FCompany_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��ǥ�ڸ�</td>
		<td><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����ڹ�ȣ</td>
		<td><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��������</td>
		<td><%= ogroup.FOneItem.Fjungsan_gubun %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����������</td>
		<td><%= ogroup.FOneItem.Fcompany_address %>&nbsp;<%= ogroup.FOneItem.Fcompany_address2 %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����</td>
		<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">����</td>
		<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">��꼭������</td>
		<% if False and (ojungsan.FItemList(0).Fdifferencekey>0) then %>
		<td><input type=text name=write_date value="" size="10" maxlength=10 readonly ><a href="javascript:calendarOpen(frm.write_date);"><img src="/images/calicon.gif" border=0 align=absmiddle></a></td>
		<% else %>
		<td><input type=text name=write_date value="<%= ojungsan.FItemList(0).GetPreFixSegumil %>" size="10" maxlength=10 readonly style="border:0"></td>
		<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">����ݾ�</td>
		<td><b><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></b> (���ް� : <%= FormatNumber(ojungsan.FItemList(0).GetTotalTaxSuply,0) %> �ΰ���: <%= FormatNumber(ojungsan.FItemList(0).GetTotalTaxvat,0) %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;<b>* �ſ� 10�� ���� ����� : �������</b><br>
			&nbsp;&nbsp;<b>* �ſ� 11�� ���� ����� : �̿�����(�Ա�ó���� �̿�(15��)�˴ϴ�.)</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">�������ڸ�</td>
		<td><%= ogroup.FOneItem.Fjungsan_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">��������E-mail</td>
		<td><%= ogroup.FOneItem.Fjungsan_email %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">�������� �ڵ�����ȣ</td>
		<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;* ��ü������ Ȯ���Ͻð�, ���Էµ� ������ ���� ��ü������������ ������ �����Ͻñ� �ٶ��ϴ�.<br>
			&nbsp;&nbsp;* ���������� ������ �Է��Ͻø�, ���ݰ�꼭�� �����Ȳ�� E-mail�� ���ڼ��񽺷� �˷��帳�ϴ�.
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" align="center">
	    <% if (ojungsan.FItemList(0).Favailneo=1) then %>
	    <input type="radio" name="billSite" value="N" checked ><strong>�׿���Ʈ </strong>
	    <input type="radio" name="billSite" value="B" ><font color=red><strong>bill36524.com (2010�����)</strong></font>
	    <% else %>
	    <input type="radio" name="billSite" value="N" disabled ><font color=gray><strong>�׿���Ʈ (���Ұ�<!--�׿���Ʈ����ȸ��-->)</strong></font>
	    <input type="radio" name="billSite" value="B" checked><font color=red><strong>bill36524.com (2010�����)</strong></font>
	    <% end if %>
	    </td>
	</tr>

</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">

    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type=button name="evalButton" value="���� <%= stypename %> ����" onClick="ActTaxReg(frm)">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- ǥ �ϴܹ� ��-->
<div id='POPBillLogin' style='position:absolute; left:100px; top:240px; width:140; height:100; z-index:2; visibility: hidden'>
<table width="420" height="260" border="0" cellpadding="0" cellspacing="2" bgcolor="#000000" class="a">
  <form name="billfrm">
  <tr >
    <td height="20" onMouseDown="mdown(event);" onMouseUp="mup();"  class="drag" bgcolor="#333399">
    &nbsp;<font color="#ffffff"><strong>bill36524 ��꼭����</strong></font>
    </td>
  </tr>
  <tr>
    <td height="210" colspan="2" valign="top" bgcolor="#FFFFFF" align="center">
        <table border=0 width="100%" class="a">
        <tr>
            <td>
            <table border=0 width="90%" class="a">
                <tr>
                    <td>1. http://www.bill36524.com �� ȸ���������ϼ���.</td>
                </tr>
                <tr>
                    <td>2. ��������� �� �������Ʈ�� �����ϼ���.</td>
                </tr>
                <tr>
                    <td>3. �Ʒ� ���������� http://www.bill36524.com �� ���̵�� �н����带 �Է��Ͻ��� ��꼭���� ��ư�� Ŭ���ϼ���.</td>
                </tr>
                <tr>
                    <td>4. ��꼭 ����� �ð��� ���� �ҿ�� �� ������ ��ٷ��ֽñ� �ٶ��ϴ�.(���� 1��)</td>
                </tr>
            </table>
            </td>
        </tr>
        <tr height="120">
            <td align="center">
                <table border="0" cellspacing="2" cellpadding="2" width="330" height="100" class="a" bgcolor="#CCCCCC" >
                    <tr bgcolor="#FFFFFF">
                        <td width="130">bill36524 ���̵�</td>
                        <td><input type="text" name="billid" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) billfrm.billpass.focus();"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                        <td width="130">bill36524 �н�����</td>
                        <td><input type="password" name="billpass" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) billTaxEval(billfrm);"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF" id="ievalBtn">
                        <td colspan="2" align="center"><input type="button" value="��꼭����" onclick="billTaxEval(billfrm);"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF" id="idoingMsg" style="display:none">
                        <td colspan="2" align="center"><img src="http://fiximage.10x10.co.kr/web2007/receipt/loading.gif" width="269" height="14">
                        <br><div id="txtMsg" name="txtMsg"><!-- ó�����Դϴ�.��ø� ��ٷ��ּ���... --></div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        </table>
    </td>
  </tr>
  <tr id="popcloseId" ><td bgcolor="#FFFFFF" align="right"><a href="javascript:hideLogin();">close</a></td></tr>
  </form>
</table>
</div>

<form name="taxSaveFrm" method="post">
<input type="hidden" name="idx" value="">
<input type="hidden" name="result" value="">
<input type="hidden" name="no_tax" value="">
<input type="hidden" name="no_iss" value="">
<input type="hidden" name="billsiteCode" value="B"> <!-- ����B, ��ĳ��W -->
<input type="hidden" name="result_msg" value="">
<input type="hidden" name="jungsangubun" value="ON">
<input type="hidden" name="write_date" value="<%= ojungsan.FItemList(0).GetPreFixSegumil %>">
<input type="hidden" name="jungsanid" value="<%= ojungsan.FItemList(0).FId %>">
</form>
<iframe name="ipreSave" id="ipreSave" width="0" height="0"></iframe>
<%
set ojungsan = Nothing
set opartner = Nothing
set ogroup = Nothing
%>

<script language=javascript>
function SvcErrMsg(){
    //alert('��� ���� �������Դϴ�.');
    var alertStr;
    //alertStr = "�̹��޺��� ���ڼ��ݰ�꼭 ��������� ��������� �ֽ��ϴ�.";
    //alertStr += "\n\nbill36524�� �α��� �Ͻ� ��, ���� ���θ޴��� [�����ȯ�漳��]����";
    //alertStr += "\n4��° �׸� �ִ� ������ ����� ���ּž� ���������� ���ݰ�꼭�� ��������˴ϴ�.";
    //alertStr += "\n������ ����� �ȵǾ� ���� ���, �ٹ�����SCM���� ���������� ���� �ʽ��ϴ�.\n\n";

    //alertStr = "!!! ���� bill36524.com ������ ��Ȱ���� �ʽ��ϴ�.";
    //alertStr += "\n����� �ð��� �����ɸ���� ��� �� �õ��Ͻñ� �ٶ��ϴ�."
    //alertStr = "";

    if (alertStr!="") {
        //alert(alertStr);
    }

}
window.onload = SvcErrMsg;
</script>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
