<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
dim IsGroupidValid : IsGroupidValid =false
dim ErrInfoStr
dim groupid,makerid,i
dim ecAUser,ecBUser
	makerid         = requestCheckvar(request("makerid"),32)
	groupid         = requestCheckvar(request("groupid"),10)


if (makerid<>"") then
    groupid = getPartnerId2GroupID(makerid)

    if (groupid="") then
        ErrInfoStr = "�׷��ڵ尡 �������� �ʾҽ��ϴ�.("&makerid&")"
    end if
end if



dim ogroupInfo
SET ogroupInfo = new CPartnerGroup
ogroupInfo.FRectGroupid = groupid
if (groupid<>"") then
    ogroupInfo.GetOneGroupInfo

    if (ogroupInfo.FResultCount<1) then
        ErrInfoStr = "�ش� �׷� ������ �����ϴ�.("&groupid&")"
    end if
end if

''�⺻ ��༭ ����Ʈ
dim oDftContractList
set oDftContractList = new CPartnerContract
oDftContractList.FPageSize=20
oDftContractList.FCurrPage = 1
oDftContractList.FRectGroupID = groupid
oDftContractList.FRectContractTypeGbn="D" ''�⺻��༭
if (groupid<>"") then
    oDftContractList.GetNewContractList
end if


''�μ����Ǽ� ����ƮON
dim oAddContractList
set oAddContractList = new CPartnerContract
oAddContractList.FPageSize=20
oAddContractList.FCurrPage = 1
oAddContractList.FRectMakerid = makerid
oAddContractList.FRectGroupID = groupid
oAddContractList.FRectContractTypeGbn="A" ''�μ����Ǽ�
if (groupid<>"") then
    oAddContractList.GetCurrAddContractListONBrand
end if

''�μ����Ǽ� ����ƮOFF
dim oAddContractListOF
set oAddContractListOF = new CPartnerContract
oAddContractListOF.FPageSize=30
oAddContractListOF.FCurrPage = 1
oAddContractListOF.FRectMakerid = makerid
oAddContractListOF.FRectGroupID = groupid
oAddContractListOF.FRectContractTypeGbn="A" ''�μ����Ǽ�
if (groupid<>"") then
    oAddContractListOF.GetCurrAddContractListOFBrand
end if

dim ContractID,mode, ContractType, sqlStr
dim isReqOpenContractExists : isReqOpenContractExists=false

dim isOldBrand : isOldBrand = fnCgeckIsOldBrand(makerid,2)
dim isOnContractExists, isOfContractExists
dim isOnHoldContract, isOFHoldContract

Call fnCheckHoldContract(makerid, isOnHoldContract, isOFHoldContract)

dim ideFaultCtrDate, def_enddate, nmonth

if (Now()<"2014-01-01") then
    ideFaultCtrDate = "2014-01-01"
else
    ideFaultCtrDate = Left(Now(),10)  ''Left(Buf,4)+"�� "+Mid(Buf,6,2)+"�� "+Mid(Buf,9,2)+"��" //��༭ ���븸 ġȯ
end if

nmonth = mid(ideFaultCtrDate,6,2)

if (nMonth<=3) then
    def_enddate = year(date())&"-06-30"
elseif (nMonth>3 and nMonth<=6) then
    def_enddate = year(date())&"-09-30"
elseif (nMonth>6 and nMonth<=9) then
    def_enddate = year(date())&"-12-31"
elseif (nMonth>9 and nMonth<=12) then
    def_enddate = year(dateadd("yyyy",1,date())) &"-03-31"
end if

''-------------------------------------------------------------------------------------------------
'dim opartner
'set opartner = new CPartnerUser
'opartner.FRectDesignerID = makerid
'
'if (makerid<>"") then
'    opartner.GetOnePartnerNUser
'end if
'
'
'
'''������ ��༭ or �������� ��༭
'dim ocontract, ocontractDetail
'set ocontract = new CPartnerContract
'ocontract.FRectContractID = ContractID
'ocontract.FRectMakerID = makerid
'
'if (ContractID<>"") then
'    ocontract.GetOneContract
'elseif (mode="") then
'    'ocontract.GetLastOneContract
'end if
'
'if ocontract.FResultCount>0 then
'    ContractID = ocontract.FOneItem.FContractID
'end if
'
'set ocontractDetail = new CPartnerContract
'ocontractDetail.FRectContractID = ContractID
'if (ContractID<>"") then
'    ocontractDetail.GetContractDetailList
'end if
'
''' ���õ�(��������) ����� �ִ°��
'dim CONTRACTING_EXISTS
'CONTRACTING_EXISTS = ocontract.FresultCount>0
'
''' ������ ����� ���°�� : ��༭ ProtoType �⺻ Setting
'if (Not CONTRACTING_EXISTS) and (opartner.FResultCount>0) and (ContractType="") then
'    if opartner.FOneItem.Fmaeipdiv="U" then
'        ContractType="5"
'    elseif opartner.FOneItem.Fmaeipdiv="W" then
'        ContractType="1"
'    elseif opartner.FOneItem.Fmaeipdiv="M" then
'        ContractType="2"
'    end if
'else
'    if ocontract.FResultCount>0 then
'        ContractType = ocontract.FOneItem.FContractType
'    end if
'end if
'
'dim ocontractProtoType
'set ocontractProtoType = new CPartnerContract
'ocontractProtoType.FRectContractType = ContractType
'
'if (Not CONTRACTING_EXISTS) and (ContractType<>"") then
'    ocontractProtoType.getOneContractProtoType
'end if
'
'dim ocontractProtoTypeDetail
'set ocontractProtoTypeDetail = new CPartnerContract
'ocontractProtoTypeDetail.FRectContractType = ContractType
'
'if (Not CONTRACTING_EXISTS) and (ContractType<>"") then
'    ocontractProtoTypeDetail.getContractDetailProtoType
'end if
'
'dim sqlStr,marginRows
'sqlStr = "select mwdiv, (100-buycash/sellcash*100) ,count(itemid) as cnt"
'sqlStr = sqlStr & " from [db_item].[dbo].tbl_item"
'sqlStr = sqlStr & " where itemid<>0"
'sqlStr = sqlStr & " and makerid='" & makerid & "'"
'sqlStr = sqlStr & " and sellcash<>0"
'sqlStr = sqlStr & " and sellyn='Y'"
'sqlStr = sqlStr & " and isusing='Y'"
'sqlStr = sqlStr & " group by mwdiv, (100-buycash/sellcash*100)"
'if makerid<>"" then
'    rsget.Open sqlStr,dbget,1
'    if  not rsget.EOF  then
'        marginRows = rsget.getRows()
'    end if
'    rsget.close
'end if
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="contract.js"></script>
<script language='javascript'>

function ChangeBrand(comp){
    var frm = document.frmReSearch;
    frm.makerid.value = comp.value;
    frm.ContractType.value = "";
    frm.submit();
}

function ChangeContractID(v){
    var frm = document.frmReSearch;
    frm.ContractType.value = "";
    frm.ContractID.value = v;
    frm.submit();
}

function ChangeContractType(comp){
    var frm = document.frmReSearch;
    frm.ContractType.value = comp.value;
    frm.submit();
}

function NewContractReg(){
    var frm = document.frmReSearch;
    frm.ContractType.value = "";
    frm.mode.value = "RegContract";
    frm.submit();
}

function SaveContract(frm){

    if (frm.makerid.value.length<1){
        alert('�귣�� ���̵� �����ϼ���.');
        frm.makerid.focus();
        return;
    }

    if (frm.contractType.value.length<1){
        alert('��༭ ������ �����ϼ���.');
        frm.contractType.focus();
        return;
    }

    //�ӽ�
    if (frm.contractType.value=="2"){
        alert('���� ���԰�༭�� �������� �ʽ��ϴ�.');
        frm.contractType.focus();
        return;
    }

    for (var i=0;i<frm.elements.length;i++){

        if (frm.elements[i].type=="text"){
            if (frm.elements[i].value.length<1){
                alert('�ʼ� �Է� �����Դϴ�.');
                frm.elements[i].focus();
                return;
            }
        }
    }

    if (confirm('��༭�� ����Ͻðڽ��ϱ�?')){
		frm.action = 'ctrReg_Process.asp';
        frm.submit();
    }
}

function preViewContract(ContractID){
    var popwin = window.open('preViewContract.asp?ContractID=' + ContractID,'preViewContract','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function DocDownloadContract(ContractID){
    var popwin = window.open('DocDownloadContract.asp?ContractID=' + ContractID,'DocDownloadContract','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goNextState(CurrState,NextState,confirmMsg){
    if (!confirm(confirmMsg)) return;

    var frm = document.frmReg;
    frm.mode.value = "stateChange";
    frm.CurrState.value = CurrState;
    frm.NextState.value = NextState;

    frm.submit();
}

//---------------------------------------------------------------------------------
function rejectContract(onoff){
    var frm = document.frmAct;
    frm.addsellplace.value = onoff;

    if (confirm('��� ���� �귣��� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value="rjtCtr";
        frm.submit();
    }
}

function rejectExpireContract(onoff){
    var frm = document.frmAct;
    frm.addsellplace.value = onoff;

    if (confirm('��� ���� ���� �Ͻðڽ��ϱ�?')){
        frm.mode.value="rjtCtrDel";
        frm.submit();
    }
}

function addDefaultContract(groupid){
    //������ ��༭ confirm
    var frm = document.frmCtrAdd;
    if (frm.defaultCtrKey.value!=""){
        if (frm.defaultCtrState.value=="7"){ //���Ϸ�
            if (!confirm('��� �Ϸ�� �⺻ ��༭�� �ֽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                return;
            }
        }else{
          //  alert('�������� �⺻ ��༭�� �ֽ��ϴ�. ���� �Ǵ� ���Ϸ��� �߰� ��� �����մϴ�.');
          //  return;
        }
    }
    if (document.frmCtrAdd.addftkey){
        $("#divaddDftCtr").empty();
    }else{
        $.ajax({
    		url: "/admin/member/contract/ajaxContract.asp?mode=addDft&groupid="+groupid,
    		cache: false,
    		async: false,
    		success: function(message) {

           		// ���� �ֱ�
           		$("#divaddDftCtr").empty().html(message);
    		}
    	});
    };

    $("#TRaddDftCtr").toggle();
}

function addAdditionContract(groupid,makerid,sellplace,mwdiv,addmargin,scmmwdiv,scmmargin,istate,isellitemcnt,ijungsansum,ecAUser, ecBUser){
    <% if isOnHoldContract then %>
        alert('��� ���� �귣���Դϴ�. ��� ���� ������ ��밡���մϴ�.');
        return;
    <% end if %>

    var frm=document.frmCtrAdd;

    if (istate!=""){
        if (istate=="7"){
            if (!confirm('��� �Ϸ�� ��༭�� �ֽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                return;
            }
        }else{
            alert('�������� ��༭�� �ֽ��ϴ�. ���� �Ǵ� ���Ϸ��� �߰� ��� �����մϴ�.');
            return;
        }
    }

    <% if (isOldBrand) then %>
    if ((isellitemcnt<1)&&(ijungsansum<1)){
        if (!confirm('�Ǹ� ��ǰ �� �ֱ� ������� �����ϴ�. ����Ͻðڽ��ϱ�?')){
            return;
        }
    }
    <% end if %>

    //�¶��� ��ǥ�������� ����.

    if ((scmmargin==0)&&(mwdiv==frm.onDfMaeipdiv.value)){
        addmargin = frm.onDfMargin.value;
    }


    var marginBox = "";
    marginBox += "<input type='hidden' name='addsellplace' value='"+sellplace+"'>"
    marginBox += "<input type='hidden' name='scmmwdiv' value='"+scmmwdiv+"'>"
    marginBox += "<input type='hidden' name='scmmargin' value='"+scmmargin+"'>"
    marginBox += "<input type='hidden' name='addmwdiv' value='"+mwdiv+"'>"
    marginBox += "<input type='text' name='addmargin' value='"+addmargin+"' class='text' size='6' style='text-align:center'>%"

    $("#addON_"+mwdiv).empty().html(marginBox);
    $("#addON_"+mwdiv).toggle();

    if (!$("#addON_"+mwdiv).is(":visible")){
        $("#addON_"+mwdiv).empty();
    }

    var ctrDateBox ="";
    ctrDateBox += "<table width='100%' border='0' cellspacing='1' cellpadding='4' class='a' bgcolor='#BABABA'>"
 		ctrDateBox += "	<tr >"
 		ctrDateBox += "		<td bgcolor='#DDDDFF' width='20%' colspan='2' align='center'>�����</td><td bgcolor='#FFFFFF'><input type='text' class='text' name='addON_ctrDate' value='<%=ideFaultCtrDate%>' size='10' maxlength='10'></td>"
 		ctrDateBox += "		<td bgcolor='#DDDDFF'  width='20%' align='center'>���������</td><td  bgcolor='#FFFFFF'><input type='text' class='text' name='addON_endDate' value='<%=def_enddate%>' size='10' maxlength='10'></td>	"
 		ctrDateBox += "	</tr>"
 		ctrDateBox += "	<tr> "
 		ctrDateBox += "</table>"

   // ctrDateBox += "�����:<input type='text' class='text' name='addON_ctrDate' value='' size='10' maxlength='10'>"
   // ctrDateBox += "&nbsp;&nbsp;���������:<input type='text' class='text' name='addON_endDate' value='' size='10' maxlength='10'>"
    var addExists=$("#addON_M").is(":visible")||$("#addON_W").is(":visible")||$("#addON_U").is(":visible");


    if (addExists){
        $("#addON_ctrData").empty().html(ctrDateBox);
        $("#addON_ctrData").show();
    }else{
        $("#addON_ctrData").empty().hide();
    }
/*
    var dlvBox =""
    if ((mwdiv=="U")&&(dlvtype=="9")){ //���ǹ���ΰ�츸
        dlvBox += " &nbsp;&nbsp;�����å:<select class='select' name='addON_dlvtype'><option value=''>�⺻��å<option value='9' "+((dlvtype=="9")?"selected":"")+">��ü���ǹ��<option value='7' "+((dlvtype=="7")?"selected":"")+">��ü���ҹ��</select>"
        dlvBox += " <input type='text' class='text' name='addON_dlvlimit' value='"+dlvmilit+"' size='7' style='text-align:right'>�̸�"
        dlvBox += " <input type='text' class='text' name='addON_dlvpay' value='"+dlvpay+"' size='5' maxlength='5' style='text-align:right'>��"
        if (addExists){
            $("#addON_ctrDlv").empty().html(dlvBox);
            $("#addON_ctrDlv").show();
        }else{
            $("#addON_ctrDlv").empty().hide();
        }
    }
*/
}

function addAdditionContractOFF(groupid,makerid,sellplace,mwdiv,addmargin,scmmwdiv,scmmargin,mjmwdiv,mjmargin,istate, ijungsansum,ecAUser, ecBUser){
    <% if isOfHoldContract then %>
        alert('��� ���� �귣���Դϴ�. ��� ���� ������ ��밡���մϴ�.');
        return;
    <% end if %>

    if (istate!=""){
        if (istate=="7"){
            if (!confirm('��� �Ϸ�� ��༭�� �ֽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                return;
            }
        }else{
            alert('�������� ��༭�� �ֽ��ϴ�. ���� �Ǵ� ���Ϸ��� �߰� ��� �����մϴ�.');
            return;
        }
    }

    //���������� ��ǥ ��� ������ ���� ����Ǿ�� ���� ==> ��༭ �ۼ��� ��ǥ���� ����(��ǥ������ ������츸).
    if (mjmwdiv==""){
        alert('��ǥ������ �����Ǿ� ���� �ʽ��ϴ�.\n\n���� ��ǥ������ ������ ���� ���� �����մϴ�.');
        return;
    }

    <% if (isOldBrand) then %>
    if ((ijungsansum<1)){
        //if (!confirm('�ֱ� ������� �����ϴ�. ����Ͻðڽ��ϱ�?')){
        //    return;
        //}
    }
    <% end if %>

    if (scmmwdiv==""){
        if (!confirm('SCM ���� ��౸���� �����ϴ�. ��� �Ͻðڽ��ϱ�?\n\n��༭ ���½� �ű԰�� ������ SCM ��౸��/������ �����˴ϴ�.')){
            return;
        }
    }

    var marginBox = "";
    marginBox += "<input type='hidden' name='addsellplace' value='"+sellplace+"'>"
    marginBox += "<input type='hidden' name='scmmwdiv' value='"+scmmwdiv+"'>"
    marginBox += "<input type='hidden' name='scmmargin' value='"+scmmargin+"'>"
    marginBox += "<br><select name='addmwdiv'><option value='B012' "+((mwdiv=="B012")?"selected":"")+">��ü��Ź<option value='B031' "+((mwdiv=="B031")?"selected":"")+">������<option value='B013' "+((mwdiv=="B013")?"selected":"")+">�����Ź</select>"
    marginBox += "<input type='text' name='addmargin' value='"+addmargin+"' class='text' size='5' style='text-align:center'>%"

    $("#add"+sellplace).empty().html(marginBox);
    $("#add"+sellplace).toggle();

    if (!$("#add"+sellplace).is(":visible")){
        $("#add"+sellplace).empty();
    }

     var ctrDateBox ="";
    ctrDateBox += "<table width='100%' border='0' cellspacing='1' cellpadding='4' class='a' bgcolor='#BABABA'>"
 		ctrDateBox += "	<tr >"
 		ctrDateBox += "		<td bgcolor='#DDDDFF' width='20%' colspan='2' align='center'>�����</td><td bgcolor='#FFFFFF'><input type='text' class='text' name='addOF_ctrDate' value='<%=ideFaultCtrDate%>' size='10' maxlength='10'></td>"
 		ctrDateBox += "		<td bgcolor='#DDDDFF'  width='20%' align='center'>���������</td><td  bgcolor='#FFFFFF'><input type='text' class='text' name='addOF_endDate' value='<%=def_enddate%>' size='10' maxlength='10'></td>	"
 		ctrDateBox += "	</tr>"
 		ctrDateBox += "</table>"


    var addExists=true;//$("#addON_M").is(":visible")||$("#addON_W").is(":visible")||$("#addON_U").is(":visible");

    if (addExists){
        $("#addOF_ctrData").empty().html(ctrDateBox);
        $("#addOF_ctrData").show();
    }else{
        $("#addOF_ctrData").empty().hide();
    }


}


function regContract(itype){
    var frm=document.frmCtrAdd;

    //�⺻��༭ ���� check

    if ((frm.defaultCtrKey.value=="")&&(!frm.addftkey)){
        alert('�⺻ ��༭�� ���� ���� �ʽ��ϴ�.\n\n�⺻��༭ �ű� �߰� �� �ۼ��� �� �ֽ��ϴ�.');
        return;
    }


    //�μ����Ǽ� ���� üũ
    if (frm.addmwdiv){
        if (frm.addmwdiv.length){

            for (var i=0;i<frm.addmwdiv.length;i++){
                if (frm.addmwdiv[i].value.length<1){
                    alert('����(����)������ �����ϼ���.');
                    return;
                }
            }
        }else{
            if (frm.addmwdiv.value.length<1){
                alert('����(����)������ �����ϼ���.');
                return;
            }
        }
    }

    if (frm.addmargin){
        if (frm.addmargin.length){
            for (var i=0;i<frm.addmargin.length;i++){
                if ((frm.addmargin[i].value.length<1)||(frm.addmargin[i].value*1<1)||(frm.addmargin[i].value*1>=100)){
                    alert('������ ��Ȯ�� �Է� �ϼ���(1~99)');
                    frm.addmargin[i].focus();
                    return;
                }

                //��ǥ������ �������� check
                if ((frm.onDfMaeipdiv.value==frm.addmwdiv[i].value)&&(frm.onDfMargin.value!=frm.addmargin[i].value)){
                    if (!confirm('�¶��� �⺻ ���� '+frm.onDfMargin.value+'�� ������ ������ �ٸ��ϴ�.\n\n����Ͻðڽ��ϱ�?\n\n(��༭ ���½� SCM���� ������ �ڵ����� ������Ʈ �˴ϴ�.)')){
                        frm.addmargin[i].focus();
                        return;
                    }
                }
            }
        }else{
            if ((frm.addmargin.value.length<1)||(frm.addmargin.value*1<1)||(frm.addmargin.value*1>=100)){
                alert('������ ��Ȯ�� �Է� �ϼ���(1~99)');
                frm.addmargin.focus();
                return;
            }

            //��ǥ������ �������� check
            if ((frm.onDfMaeipdiv.value==frm.addmwdiv.value)&&(frm.onDfMargin.value!=frm.addmargin.value)){
                if (!confirm('�¶��� �⺻ ���� '+frm.onDfMargin.value+'�� ������ ������ �ٸ��ϴ�.\n\n����Ͻðڽ��ϱ�?\n\n(��༭ ���½� SCM���� ������ �ڵ����� ������Ʈ �˴ϴ�.)')){
                    frm.addmargin.focus();
                    return;
                }
            }
        }
    }

    if ((!frm.addftkey)&&(!frm.addmwdiv)){
        alert('����� ��༭�� �����ϴ�. - �⺻��༭ �Ǵ� �μ����Ǽ� ��� �� ����Ͻñ� �ٶ��ϴ�.');
        return;
    }




    //ON OF ���� ��Ͻ� ��༭ ��¥�� �����ؾ���
    if ((frm.addON_ctrDate)&&(frm.addOF_ctrDate)){
        if (frm.addON_ctrDate.value!=frm.addOF_ctrDate.value){
            alert('��/���� �μӳ��� ���� ���� ������� ���� �ؾ� �մϴ�.');
            return;
        }
    }

   // if (frm.defaultCtrKey.value!=""){

if(itype==2 ) {
	   var winH = $(document).height()/2-500;
		var winW = $(document).width();

  		$("#ecLyr").css('top', winH-$("#ecLyr").height());
		$("#ecLyr").css('left', winW/2-$("#ecLyr").width()/2);
		$("#ecDiv").show();

  }else{
        if (confirm('��༭�� �ű� ��� �Ͻðڽ��ϱ�?')){
        	frm.signtype.value = itype;
        	frm.action="ctrReg_Process.asp"

            frm.submit();
        }
      }
  //  }else{
   //     if (confirm('��༭�� �ű� ��� �Ͻðڽ��ϱ�?')){
  //      		frm.signtype.value = itype;
  //          frm.submit();
  //      }
   // }
}

 	function jsEcCancel(){
 		$("#ecDiv").hide();
 	}

 	function jsEcSubmit(){
 		var frm = document.frmCtrAdd;

//    		if(!document.frmEc.LgUID.value){
//    			alert("���ڰ����� ���̵� �Է����ּ���");
//    			return;
//    		}
//    		if(!document.frmEc.LgUPW.value){
//    			alert("���ڰ����� ��й�ȣ�� �Է����ּ���");
//    			return;
//    		}
    		if(!document.frmEc.ecAUser.value){
    			alert("�ٹ����� ���ڰ�����ڸ�(LG u+ ����Ʈ��)�� �Է����ּ���");
    			return;
    		}
//    		if(!document.frmEc.ecBUser.value){
//    			alert("���»� ���ڰ�����ڸ�(LG u+ ����Ʈ��)�� �Է����ּ���");
//    			return;
//    		}

 		 if (confirm('��༭�� �ű� ��� �Ͻðڽ��ϱ�?')){
        	frm.signtype.value = 2;
       // 	frm.LgUID.value = document.frmEc.LgUID.value;
 	//		frm.LgUPW.value = document.frmEc.LgUPW.value;
 			frm.ecAUser.value = document.frmEc.ecAUser.value;
 			frm.ecBUser.value = document.frmEc.ecBUser.value;
        	frm.action="ctrReg_Process.asp";

            frm.submit();
        }

 	}

function popShopUpcheInfo(imakerid){
    var popwin = window.open('/admin/lib/popshopupcheinfo.asp?shopid=streetshop000&designer='+imakerid,'popShopUpcheInfo','width=800,height=900,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function chgBrand(comp){
    var imakerid=comp.value;

    if (comp.value!=""){
        document.frm.makerid.value=imakerid;
        document.frm.submit();
    }
}

function jsSetEcState(){
	$("#btnSubmit").prop("disabled", true);
	document.frmEcState.submit();
}

function jsSetUser(){
	/*
	if(!document.frmCtrAdd.ecBU.value){
 		alert("����ڸ��� �Է����ּ���");
 		return;
	}
	*/
	document.frmEcUser.ecBUser.value = document.frmCtrAdd.ecBU.value;
	document.frmEcUser.submit();
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<form name="frm" method="get" action="">
<input type="hidden" name="groupid" value="<%=groupid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">

    �귣��ID : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
    &nbsp;&nbsp;
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<form name="frmEcState" method="post"  action="ctrReg_ProcessTest.asp">
<input type="hidden" name="mode" value="ecstate">
<input type="hidden" name="groupid" value="<%=groupid%>">
</form>
<form name="frmEcUser" method="post" action="ctrReg_Process.asp">
	<input type="hidden" name="mode" value="ecuser">
	<input type="hidden" name="groupid" value="<%=groupid%>">
	<input type="hidden" name="ecBUser" value="">
</form>
<% if ogroupInfo.FResultCount>0 then %>
<form name="frmCtrAdd" method="post" >
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<input type="hidden" name="mode" value="reg">
<input type="hidden" name="groupid" value="<%=groupid%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="signtype" value="">
<input type="hidden" name="LgUID" value="">
<input type="hidden" name="LgUPW" value="">
<input type="hidden" name="ecAUser" value="">
<input type="hidden" name="ecBUser" value="">
<tr bgcolor="#FFFFFF">
    <td colspan="4"><b>* ��ü�⺻����</b></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="10%" bgcolor="<%= adminColor("gray") %>">��ü��</td>
    <td width="40%" ><%=ogroupInfo.FOneItem.Fcompany_name %></td>
    <td width="10%" bgcolor="<%= adminColor("gray") %>">�׷��ڵ�</td>
    <td width="40%" ><a href="javascript:PopUpcheInfoEdit('<%=ogroupInfo.FOneItem.Fgroupid %>');"><%=ogroupInfo.FOneItem.Fgroupid %></a>
    &nbsp;
    <% CALL DrawSameGroupBrand(ogroupInfo.FOneItem.Fgroupid,makerid,"linkmakerid","onChange='chgBrand(this)'") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("gray") %>">��ǥ��</td>
    <td ><%=ogroupInfo.FOneItem.Fceoname %></td>
    <td bgcolor="<%= adminColor("gray") %>">����ڹ�ȣ</td>
    <td ><%=ogroupInfo.FOneItem.Fcompany_no %><input type="hidden" name="bcompno" value="<%=ogroupInfo.FOneItem.Fcompany_no %>"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("gray") %>">������ּ�</td>
    <td colspan="3"><%=ogroupInfo.FOneItem.Fcompany_address %>&nbsp;<%=ogroupInfo.FOneItem.Fcompany_address2 %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("gray") %>">������</td>
    <td colspan="3" >
    �¶��� : <%= ogroupInfo.FOneItem.Fjungsan_date %>
    &nbsp;/&nbsp;�������� :
    <% if ogroupInfo.FOneItem.Fjungsan_date<>ogroupInfo.FOneItem.Fjungsan_date_off then %>
    <font color="red"><%= ogroupInfo.FOneItem.Fjungsan_date_off %></font>
    <% else %>
    <%= ogroupInfo.FOneItem.Fjungsan_date_off %>
    <% end if %>
    </td>

</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("gray") %>">���ڰ������</td>
    <td colspan="3"><input type="text" name="ecBU" value="<%=fnGetEcBUser(ogroupInfo.FOneItem.Fgroupid)%>"> <input type="button" class="button" value="����" onClick="jsSetUser();"></td>
  </tr>
</table>
<div style="text-align:right;padding:5px;">
 <span style="left-margin:10px;"><input type="button" id="btnSubmit" class="button" value="���ڰ�༭ ����Update" onClick="jsSetEcState('<%=groupid%>');"></span>
    <span id="reqCtrOpen" style="display:none"><input type="button" class="button" value="��༭ ����" onClick="popOpenContract('<%=groupid%>');"></span>
    </div>
<p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="8"><b>* �⺻��༭</b></td>
    <td align="right"><a onClick="addDefaultContract('<%=groupid%>');" style="cursor:pointer">�ű� <img src="/images/icon_plus.gif" align="absmiddle"></a></td>
</tr>

<tr bgcolor="#FFFFFF" id="TRaddDftCtr" style="display:none">
    <td colspan="9">
    <div id="divaddDftCtr"><div>
    </td>
</tr>

<tr bgcolor="<%= adminColor("gray") %>" align="center">
	<td>��༭Ÿ��</td>
    <td>��༭��ȣ</td>
    <td>��༭��</td>
    <td>�����</td>
    <td>���������</td>
    <td>���������</td>
    <td>�������</td>
    <td>�����</td>
    <td>�߼���</td>
</tr>
<% if (oDftContractList.FResultCount<1) then %>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center"> �⺻��༭�� �����ϴ�.
    <input type="hidden" name="defaultCtrKey" value="">
    </td>
</tr>
<% else %>
<% for i=0 to oDftContractList.FresultCount-1 %>
<% if (i=0) then %>
<input type="hidden" name="defaultCtrKey" value="<%=oDftContractList.FItemList(i).FctrKey %>">
<input type="hidden" name="defaultCtrState" value="<%=oDftContractList.FItemList(i).FctrState %>">
<% end if %>
<%
if oDftContractList.FItemList(i).FCtrState=0 then
    isReqOpenContractExists=true
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%if oDftContractList.FItemList(i).FecCtrSeq <> "" or not isNull(oDftContractList.FItemList(i).FecCtrSeq ) or oDftContractList.FItemList(i).FecCtrSeq <> "0"  then%>����(<%=oDftContractList.FItemList(i).FecCtrSeq %>)<%else%>����<%end if%></td>
    <td><a href="javascript:modiContract('<%=oDftContractList.FItemList(i).FctrKey %>');"><%=oDftContractList.FItemList(i).FctrNo %></a></td>
    <td><%=oDftContractList.FItemList(i).FcontractName %></td>
    <td><%=oDftContractList.FItemList(i).FcontractDate %></td>
    <td><%=oDftContractList.FItemList(i).FendDate %></td>
    <td><%=oDftContractList.FItemList(i).FcontractJungsanDate %></td>
    <td><%=oDftContractList.FItemList(i).GetContractStateName %></td>
    <td><span title="<%=oDftContractList.FItemList(i).FRegDate %>"><%=oDftContractList.FItemList(i).FRegUserName %></span></td>
    <td>
        <!--
        <img src="/images/iexplorer.gif" style="cursor:pointer" onClick="dnWebAdm('<%=oDftContractList.FItemList(i).FctrKey %>');">
        &nbsp;
        <img src="/images/pdficon.gif" style="cursor:pointer" onClick="dnPdfAdm('<%=oDftContractList.FItemList(i).FctrKey %>');">
        -->
    </td>
</tr>
<%
 next
		ecAUser = oDftContractList.FItemList(0).FecAUser '���� �ֱ� ���� ��༭�� �����
		ecBUser = oDftContractList.FItemList(0).FecBUser

		if ecAUser = "" or isNull(ecAuser) then 			ecAUser = FecAUser
%>
<% end if %>
</table>
<p><br>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="15"><b>* �μ����Ǽ� - �¶���</b></td>
     <% if (makerid<>"") then %>
    <td align="right" colspan="2">
    <% if  (isOnHoldContract) then %>
    <input type="button" class="button" value="��ຸ������" onClick="rejectExpireContract('ON');">
    <% else %>
    <div id="reqCtrXpireOn" style="display:none"><input type="button" class="button_auth" value="��ຸ�����" onClick="rejectContract('ON');"></div>
    <% end if %>
    </td>
    <% end if %>
     <td align="right"><a href="javascript:PopBrandInfoEdit('<%=makerid%>');">SCM����</a></td>
 </tr>
 <tr  bgcolor="#FFFFFF">
 	<td colspan="18">
 		<div id="addON_ctrData" style="display:none"></div>
  </td>
</tr>
<% if (makerid="") then %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" height="30" colspan="18">�μ� ��༭�� �ۼ��Ͻ÷���, ���� �귣�带 �����ϼ���</td>
</tr>
<% else %>
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td colspan="7">�������</td>
    <td colspan="4">SCM��������</td>
    <td colspan="4">��ǰ����</td>
    <td colspan="2">3��������</td>
    <td rowspan="2">���</td>
</tr>

<tr bgcolor="<%= adminColor("gray") %>" align="center">
	<td>���Ÿ��</td>
    <td>��༭��ȣ</td>
    <td>��༭��</td>
    <td>�����</td>
    <td>���������</td>
    <td>��������/����</td>
    <td>�������</td>

    <td>�Ǹ�ó</td>
    <td>��౸��</td>
    <td>������/����</td>
    <td>��ۺ���å</td>

    <td>����</td>
    <td>����</td>
    <td>�Ǹż�</td>
    <td>����</td>

    <td>�����</td>
    <td>�ݾ�</td>
</tr>
<% for i=0 to oAddContractList.FresultCount-1 %>
<%
if oAddContractList.FItemList(i).FCtrState=0 then
    isReqOpenContractExists=true
end if

if Not isNULL(oAddContractList.FItemList(i).FctrKey) then
    isOnContractExists = true
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%if oAddContractList.FItemList(i).FecCtrSeq <> "" and not isNull(oAddContractList.FItemList(i).FecCtrSeq) and oAddContractList.FItemList(i).FecCtrSeq <> "0" then%>����(<%=oAddContractList.FItemList(i).FecCtrSeq %>)<%else%>����<%end if%></td>
    <td><a href="javascript:modiContract('<%=oAddContractList.FItemList(i).FctrKey %>');"><%=oAddContractList.FItemList(i).FctrNo %></a></td>
    <td><%=oAddContractList.FItemList(i).FcontractName %></td>
    <td><%=oAddContractList.FItemList(i).FcontractDate %></td>
    <td><%=oAddContractList.FItemList(i).FendDate %></td>
    <td><%=oAddContractList.FItemList(i).getContractMwDivStr %> <%=oAddContractList.FItemList(i).getContractMarginStr %>
    <% if oAddContractList.FItemList(i).FSeq<>0 then %>
    <span id="add<%=oAddContractList.FItemList(i).Fsellplace%>_<%=oAddContractList.FItemList(i).Fmaeipdiv%>" style="display:none"></span>
    <% else %>
    <input type="hidden" name="onDfMaeipdiv" value="<%=oAddContractList.FItemList(i).FMaeipdiv%>">
    <input type="hidden" name="onDfMargin" value="<%=oAddContractList.FItemList(i).FSCMDefaultmargine%>">
    <% end if %>
    </td>
    <td><%=fnContractStateName(oAddContractList.FItemList(i).FCtrState) %></td>
    <td><%=oAddContractList.FItemList(i).getSellplaceName %></td>
    <td><%=fnMaeipdivName(oAddContractList.FItemList(i).FMaeipdiv) %></td>
    <td><%=oAddContractList.FItemList(i).getSCMDefaultmargineStr %></td>
    <td><%=oAddContractList.FItemList(i).getSCMDefaultDlvStr %></td>

    <td <%=CHKIIF(oAddContractList.FItemList(i).FuseitemCnt<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractList.FItemList(i).FuseitemCnt,0) %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).Fuseitemmargin<1,"bgcolor='#EEBBBB'","")%> ><%=CLNG(oAddContractList.FItemList(i).Fuseitemmargin*100)/100 %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).FsellitemCnt<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractList.FItemList(i).FsellitemCnt,0) %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).Fsellitemmargin<1,"bgcolor='#EEBBBB'","")%> ><%=CLNG(oAddContractList.FItemList(i).Fsellitemmargin*100)/100 %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).FjungsanCnt<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractList.FItemList(i).FjungsanCnt,0) %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).FjungsanSum<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractList.FItemList(i).FjungsanSum,0) %></td>
    <td>
        <% if oAddContractList.FItemList(i).FSeq=0 then %>

        <% else %>
        <a onClick="addAdditionContract('<%=groupid%>','<%=makerid%>','<%=oAddContractList.FItemList(i).Fsellplace%>','<%=oAddContractList.FItemList(i).Fmaeipdiv%>','<%=oAddContractList.FItemList(i).getAddDefaultMargin%>','<%=oAddContractList.FItemList(i).Fmaeipdiv%>','<%=oAddContractList.FItemList(i).Fscmdefaultmargine%>','<%=oAddContractList.FItemList(i).FCtrState%>','<%=oAddContractList.FItemList(i).FsellitemCnt%>','<%=oAddContractList.FItemList(i).FjungsanSum%>','<%=ecAUser%>','<%=ecBUser%>');" style="cursor:pointer">�ű� <img src="/images/icon_plus.gif" align="absmiddle"></a>
        <% end if %>
    </td>
</tr>
<% if oAddContractList.FItemList(i).FSeq=0 then %>
<tr bgcolor="#FFFFFF" align="center"><td colspan="18"></td></tr>
<% end if %>
<% next %>
<% end if %>
</table>

<p><br>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="13"><b>* �μ����Ǽ� - ��������</b>
    &nbsp;&nbsp;
    <!--<span id="addOF_ctrDate" style="display:none"></span>-->
    </td>
    <% if (makerid<>"") then %>
    <td align="right" colspan="2">
    <% if (isOfHoldContract) then %>
    <input type="button" class="button" value="��ຸ������" onClick="rejectExpireContract('OF');">
    <% else %>
    <div id="reqCtrXpireOf" style="display:none"><input type="button" class="button_auth" value="��ຸ�����" onClick="rejectContract('OF');"></div>
    <% end if %>
    </td>
    &nbsp;
    <td align="right"><a href="javascript:popShopUpcheInfo('<%=makerid%>');">SCM����</a></td>
    <% end if %>
</tr>
<% if (makerid="") then %>
<tr bgcolor="#FFFFFF" align="center">
    <td align="center" height="30" colspan="14">�μ� ��༭�� �ۼ��Ͻ÷���, ���� �귣�带 �����ϼ���</td>
</tr>
<% else %>
 <tr  bgcolor="#FFFFFF">
 	<td colspan="18">
 		<div id="addOF_ctrData" style="display:none"></div>
  </td>
</tr>
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td colspan="7">�������</td>
    <td colspan="3">SCM��������</td>
    <td colspan="3">��ǥ����</td>
    <td colspan="2">3��������</td>
    <td rowspan="2">���</td>
</tr>
<tr bgcolor="<%= adminColor("gray") %>" align="center">
	<td>���Ÿ��</td>
    <td>��༭��ȣ</td>
    <td>��༭��</td>
    <td>�����</td>
    <td>���������</td>
    <td>��������/����</td>
    <td>�������</td>

    <td>�Ǹ�ó</td>
    <td>��౸��</td>
    <td>������/����</td>

    <td>��ǥ����</td>
    <td>��౸��</td>
    <td>������/����</td>

    <td>�����</td>
    <td>�ݾ�</td>
</tr>
<% for i=0 to oAddContractListOF.FresultCount-1 %>
<%
if oAddContractListOF.FItemList(i).FCtrState=0 then
    isReqOpenContractExists=true
end if

if Not isNULL(oAddContractListOF.FItemList(i).FctrKey) then
    isOfContractExists = true
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%if oAddContractListOF.FItemList(i).FecCtrSeq <> "" and not isNull(oAddContractListOF.FItemList(i).FecCtrSeq) and oAddContractListOF.FItemList(i).FecCtrSeq <> "0" then%>����(<%=oAddContractListOF.FItemList(i).FecCtrSeq %>)<%else%>����<%end if%></td>
    <td><a href="javascript:modiContract('<%=oAddContractListOF.FItemList(i).FctrKey %>');"><%=oAddContractListOF.FItemList(i).FctrNo %></a></td>
    <td><%=oAddContractListOF.FItemList(i).FcontractName %></td>
    <td><%=oAddContractListOF.FItemList(i).FcontractDate %></td>
    <td><%=oAddContractListOF.FItemList(i).FendDate %></td>
    <td><%=oAddContractListOF.FItemList(i).getContractMwDivStr %> <%=oAddContractListOF.FItemList(i).getContractMarginStr %>
    <span id="add<%=oAddContractListOF.FItemList(i).Fsellplace%>" style="display:none"></span>
    </td>
    <td><%=fnContractStateName(oAddContractListOF.FItemList(i).FCtrState) %></td>
    <td><%=oAddContractListOF.FItemList(i).getSellplaceName %></td>
    <td><%=fnMaeipdivName(oAddContractListOF.FItemList(i).FMaeipdiv) %></td>
    <td><%=oAddContractListOF.FItemList(i).getSCMDefaultmargineStr %></td>

    <td><%=oAddContractListOF.FItemList(i).FMjshopname %></td>
    <td><%=fnMaeipdivName(oAddContractListOF.FItemList(i).FMjmaeipdiv) %></td>
    <td><%=oAddContractListOF.FItemList(i).FMjdefaultmargin %></td>
    <td <%=CHKIIF(oAddContractListOF.FItemList(i).FjungsanCnt<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractListOF.FItemList(i).FjungsanCnt,0) %></td>
    <td <%=CHKIIF(oAddContractListOF.FItemList(i).FjungsanSum<1,"bgcolor='#EEBBBB'","")%> ><%=FormatNumber(oAddContractListOF.FItemList(i).FjungsanSum,0) %></td>
    <td>
        <a onClick="addAdditionContractOFF('<%=groupid%>','<%=makerid%>','<%=oAddContractListOF.FItemList(i).Fsellplace%>','<%=CHKIIF(not isNULL(oAddContractListOF.FItemList(i).Fmaeipdiv),oAddContractListOF.FItemList(i).Fmaeipdiv,oAddContractListOF.FItemList(i).FMjmaeipdiv)%>','<%=CHKIIF(not isNULL(oAddContractListOF.FItemList(i).getAddDefaultMargin),oAddContractListOF.FItemList(i).getAddDefaultMargin,oAddContractListOF.FItemList(i).FMjdefaultmargin)%>','<%=oAddContractListOF.FItemList(i).Fmaeipdiv%>','<%=oAddContractListOF.FItemList(i).Fscmdefaultmargine%>','<%=oAddContractListOF.FItemList(i).FMjmaeipdiv%>','<%=oAddContractListOF.FItemList(i).FMjdefaultmargin%>','<%=oAddContractListOF.FItemList(i).FCtrState%>','<%=oAddContractListOF.FItemList(i).FjungsanSum%>','<%=ecAUser%>','<%=ecBUser%>');" style="cursor:pointer">�ű� <img src="/images/icon_plus.gif" align="absmiddle"></a>
    </td>
</tr>
<% if oAddContractListOF.FItemList(i).FSeq=0 then %>
<tr bgcolor="#FFFFFF" align="center"><td colspan="18"></td></tr>
<% end if %>
<% next %>
<% end if %>
</form>
</table>


<p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td align="center">
    <input type="button" class="button" value="�ű� ��� ��� [������]" onClick="regContract(1)">
    <input type="button" class="button" value="�ű� ��� ��� [���ڼ���]" onClick="regContract(2)">
    </td>
</tr>
</table>

<p><p><p><p><p>
<br><br><br><br>

<% if (FALSE) then %>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="4"><b>* ���HISTORY</b></td>
</tr>
<% if (oDftContractList.FResultCount<1) then %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="4" align="center">�ֱ� ��� ������ �����ϴ�.</td>
</tr>
<% else %>
<% for i=0 to oDftContractList.FResultCount - 1 %>

<% next %>
<% end if %>
</table>
<% end if %>
<p><p>
<% end if %>
<style type="text/css">
#ecDiv {display:none; width:100%; height:100%; position:fixed; left:0; top:0; z-index:900000;}
#ecDiv .ecIn {display:; width:600px;height:100px;position:absolute; left:50%; top:50%; margin:0px 0 0 0px; background:#efefef; padding:50px; z-index:999999;}
#mask {display:; width:100%; height:100%;position:absolute; left:0; top:0; z-index:9000; background:url(http://webadmin.10x10.co.kr/images/mask_bg.png) left top repeat;}
 </style>

<div  id="ecDiv" >
	<div id="ecLyr" class="ecIn">
		<form name="frmEc" method="post" >
			<div>
		<table cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA" width="600">
			<!--<%if C_ADMIN_AUTH then%>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" width="20%" align="center"  rowspan="2" >���ڰ�����</td>
			<td bgcolor="#DDDDFF" width="20%" align="center"  >���̵�</td>
			<td><input type="text" name="LgUID" value="<%=FecID%>" size="10" class="text"></td>
		</tr>

		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" width="20%" align="center"  >��й�ȣ</td>
			<td><input type="password" name="LgUPW" value="" size="10" class="text"  AUTOCOMPLETE="off"></td>
		</tr>
		<%end if%>-->
		<tr bgcolor="#FFFFFF" >
	  	<td bgcolor="#DDDDFF" width="20%" align="center" rowspan="2" >���ڰ������<br/><span style="font-size:8pt;color:blue;">(LGU+ ����Ʈ  �����)</span></td>
	  	<td bgcolor="#DDDDFF" width="20%" align="center"  >�ٹ�����</td>
	  	<td ><input type="text" class="text" name="ecAUser" id="ecAuser" value="<%=FecAUser%>" size="10"> </td>
	  </tr>
	  <tr bgcolor="#FFFFFF" >
	  	<td bgcolor="#DDDDFF" width="10%" align="center"  >���»�</td>
	  	<td ><input type="text" class="text" name="ecBUser" value="" size="10"> (��ҹ��� ����!)</td>
	  </tr>
	</table>
	</div>
	<div style="width:100%;text-align:center;padding:5px;">
	<input type="button" class="button" value="���" onClick="jsEcCancel();" />
	<input type="button" class="button" value="���ڰ����" onClick="jsEcSubmit();" />
	</div>
	</form>
	</div>
	<div id="mask"></div>
  </div>
<form name="frmReSearch" method="get" action="">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="ContractType" value="<%= ContractType %>">
<input type="hidden" name="ContractID" value="">
</form>

<form name="frmAct"  method="post" action="ctrReg_Process.asp">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="mode" value="rjtCtr">
<input type="hidden" name="addsellplace" value="">
</form>

<script language='javascript'>
<% if isReqOpenContractExists then%>
$("#reqCtrOpen").show();
<% end if %>

<% if (makerid<>"" and isOldBrand and Not isOnContractExists) then %>
$("#reqCtrXpireOn").show();
<% end if %>

<% if (makerid<>"" and isOldBrand and Not isOfContractExists) then %>
$("#reqCtrXpireOf").show();
<% end if %>

</script>
<%
SET ogroupInfo = Nothing
SET oDftContractList = Nothing
SET oAddContractList = Nothing
SET oAddContractListOF = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
