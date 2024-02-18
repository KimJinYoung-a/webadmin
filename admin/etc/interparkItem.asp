<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim itemid, itemname, eventid, mode
dim itemidArr, eventidArr, makeridArr
dim page, makerid, ExtNotReg, MatchCate
dim delitemid, extitemid, showminusmagin, showminusmagin15, onlysoldout, onlynotusing, expensive10x10, interyes10x10no, onreginotmapping,diffPrc, isMadeHand
dim availreg, failCntExists
dim bestOrd, bestOrdMall, sellyn, sailyn
dim reqExpire, extsellyn, infoDivYn

page    = request("page")
itemid  = request("itemid")

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp) 
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

itemname= request("itemname")
eventid = request("eventid")
mode    = request("mode")
itemidArr = Trim(request("itemidArr"))
eventidArr= Trim(request("eventidArr"))
makeridArr = Trim(request("makeridArr"))
makerid= request("makerid")
ExtNotReg = request("ExtNotReg")
MatchCate = request("MatchCate")
delitemid = requestCheckvar(request("delitemid"),9)
extitemid = requestCheckvar(request("extitemid"),10)
showminusmagin = request("showminusmagin")
showminusmagin15 = request("showminusmagin15")
onlysoldout = request("onlysoldout")
onlynotusing= request("onlynotusing")
expensive10x10 = request("expensive10x10")
interyes10x10no = request("interyes10x10no")
onreginotmapping = request("onreginotmapping")
diffPrc		= request("diffPrc")
availreg    = request("availreg")
failCntExists = request("failCntExists")
bestOrd     = request("bestOrd")
bestOrdMall = request("bestOrdMall")
sellyn      = request("sellyn")
sailyn      = request("sailyn")
reqExpire   = request("reqExpire")
extsellyn   = request("extsellyn")
infoDivYn   = request("infoDivYn")
isMadeHand	= request("isMadeHand")

if page="" then page=1
if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)
if Right(eventidArr,1)="," then eventidArr=Left(eventidArr,Len(eventidArr)-1)


dim sqlStr, resultRow
if (mode="regByItemIDarr") then

elseif (mode="regByEventIDarr") then
 
elseif (mode="recentBestSeller") then

elseif (mode="regByMakerid") then

elseif (mode="delitem") then
    sqlStr = "delete from [db_item].[dbo].tbl_interpark_reg_item" + VbCrlf
    sqlStr = sqlStr + " where itemid=" & delitemid

    dbget.Execute sqlStr, resultRow
    response.write "<script >alert('" + CStr(resultRow) + "�� �����Ǿ����ϴ�.')</script>"
    dbget.close()	:	response.End
end if

  
dim oInterParkitem
set oInterParkitem = new CExtSiteItem
oInterParkitem.FPageSize		= 50
oInterParkitem.FCurrPage       = page
oInterParkitem.FRectCate_large = request("cdl")
oInterParkitem.FRectCate_mid = request("cdm")
oInterParkitem.FRectCate_small = request("cds")
oInterParkitem.FRectItemID     = itemid
oInterParkitem.FRectItemName   = itemname
oInterParkitem.FRectEventid    = eventid
oInterParkitem.FRectMakerid    = makerid
oInterParkitem.FRectExtNotReg  = ExtNotReg
oInterParkitem.FRectMatchCate  = MatchCate
oInterParkitem.FRectExtItemID  = extitemid
oInterParkitem.FRectMinusMigin = showminusmagin
oInterParkitem.FRectMinusMigin15 = showminusmagin15
oInterParkitem.FRectIsSoldOut  = onlysoldout
oInterParkitem.FRectSellYn  = sellyn
oInterParkitem.FRectSailYn  = sailyn
oInterParkitem.FRectExtSellYn  = extsellyn
oInterParkitem.FRectUseYn  = CHKIIF(onlynotusing="on","N","")
oInterParkitem.FRectExpensive10x10 = expensive10x10
oInterParkitem.FRectInteryes10x10no = interyes10x10no
oInterParkitem.FRectOnreginotmapping = onreginotmapping
oInterParkitem.FRectdiffPrc = diffPrc
oInterParkitem.FRectAvailReg = availreg
oInterParkitem.FRectFailCntExists = failCntExists
oInterParkitem.FRectFailCntOverExcept = ""
oInterParkitem.FRectInfoDivYn = infoDivYn
oInterParkitem.FRectisMadeHand		= isMadeHand

IF (bestOrd="on") then
    oInterParkitem.FRectOrdType = "B"
ELSEIF (bestOrdMall="on") then
    oInterParkitem.FRectOrdType = "BM"
end if

if (reqExpire="on") then
    oInterParkitem.GetInterParkExpireItemList
else
    oInterParkitem.GetInterParkRegedItemList
end if


'rw "ExtNotReg="&ExtNotReg
'rw "MatchCate="&MatchCate
'rw "extitemid="&extitemid
'rw "showminusmagin="&showminusmagin
'rw "showminusmagin15="&showminusmagin15
'rw "onlysoldout="&onlysoldout
'rw "expensive10x10="&expensive10x10
'rw "interyes10x10no="&interyes10x10no
'rw "onreginotmapping="&onreginotmapping
dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

// new API
function InterParkregIMSI(frm){
    var chkcnt=0;
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value==""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value==""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� ������� ���ɻ�ǰ�� �ƴմϴ�.');
        return;
    }
    
    frm.mode.value="regitemIMSIArr";
    frm.action="iParkAPI_Process.asp";
    if(confirm('���� ����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function InterParkdelIMSI(frm){
    var chkcnt=0;
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value==""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value==""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� �������� ���ɻ�ǰ�� �ƴմϴ�.');
        return;
    }
    
    frm.mode.value="delitemIMSIArr";
    frm.action="iParkAPI_Process.asp";
    if(confirm('���� �����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}


function InterParkregItemNewAPI(frm){
    var chkcnt=0;
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value==""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value==""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� �̵�� ��ǰ�� �ƴմϴ�.');
        return;
    }
    
    frm.mode.value="regitemONE";
    frm.action="iParkAPI_Process.asp";
    if(confirm('�̵�� ��ǰ�� ��� �˴ϴ�. ����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function InterParkEditItemNewAPI(frm){
    var chkcnt=0;
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value!=""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value!=""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� ��� ��ǰ�� �ƴմϴ�.');
        return;
    }
    
    frm.mode.value="edititemONE";
    frm.action="iParkAPI_Process.asp";
    if(confirm('��� ��ǰ�� ���� �˴ϴ�. �����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function InterParkSellYnProcess(frm, slYN){
    var chkcnt=0;
    
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value!=""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value!=""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� ��� ��ǰ�� �ƴմϴ�.');
        return;
    }
    
    
    
    if (slYN=="N"){
        slYNNm ="ǰ��";
        frm.mode.value="sellStatNONE";
    }else if(slYN=="X"){
        slYNNm ="�Ǹ����� �� ����";
        frm.mode.value="delitemONE";
    }
    
    frm.action="iParkAPI_Process.asp";
    if(confirm('���� ��ǰ�� '+ slYNNm +' ó�� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function InterParkSelectStatCheck(frm){
    var chkcnt=0;
    
    if (frm.cksel.length){
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                if (frm.xsiteitemno[i].value!=""){
                    chkcnt++;
                }else{
                    frm.cksel[i].checked=false;
                }
            }
        }
    }else{
        if (frm.cksel.checked){
            if (frm.xsiteitemno.value!=""){
                chkcnt++;
            }else{
                frm.cksel.checked=false;
            }
            chkcnt = 1
        }
    }
    
    if (chkcnt<1){
        alert('���õ� ��ǰ�� ���ų� ��� ��ǰ�� �ƴմϴ�.');
        return;
    }
    
    frm.mode.value="CheckItemStat";
    frm.action="iParkAPI_Process.asp";
    if(confirm('���� ��ǰ�� �Ǹ� ���¸� Ȯ���Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function InterParkSelectStatCheckBatch(frm){
    frm.mode.value="CheckItemStatBatch";
    frm.action="iParkAPI_Process.asp";
    frm.submit();
}

function InterParkItemInfoCheckBatch(frm){
    frm.mode.value="CheckItemInfo";
    if(document.frmReg.locNo.value == ""){
    	document.frmReg.locNo.focus;
    	alert('���ڸ� �Է��ϼ���');
    	return;
    }
    frm.locNo.value=document.frmReg.locNo.value;
    frm.action="iParkAPI_Process.asp";
    frm.submit();
}


function InterParkExpireItemAutoNewAPI(frm){
    frm.mode.value="delitemAuto";
    frm.action="iParkAPI_Process.asp";
    if(confirm('�����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function InterParkInfoFivNoneItemAutoNewAPI(frm){
    frm.mode.value="infoDivNone";
    frm.action="iParkAPI_Process.asp";
    if(confirm('ǰ������ ���Է� ��ǰ ǰ�� ó�� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function InterParkEditItemAutoNewAPI(frm){
    frm.mode.value="edititemAuto";
    frm.action="iParkAPI_Process.asp";
    if(confirm('�����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function InterParkregItemAutoNewAPI(frm){
    frm.mode.value="regitemAuto";
    frm.action="iParkAPI_Process.asp";
    if(confirm('����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=interpark&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}


function popItem2CategoryRedirect(itemid){
    var popwin = window.open('InterParkMatcheDispCateByitemRedirect.asp?itemid=' + itemid,'MatcheDispCate','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}


// �Ѱ����� ����..
function RegByItemID(frm){
    if (frm.itemidArr.value.length<1){
        alert('��ǰ��ȣ�� �Է��� �ּ���.');
        frm.itemidArr.focus();
        return;
    }
    
    if (confirm('������� �Ͻðڽ��ϱ�?')){
        //frm.mode.value = "regByItemIDarr";
        frm.cksel.value=frm.itemidArr.value;
        frm.mode.value="regitemIMSIArr";
        frm.action="iParkAPI_Process.asp";
    
        frm.submit();
    }
}

function RegByEventID(frm){
    if (frm.eventidArr.value.length<1){
        alert('�̺�Ʈ ��ȣ��  �Է��� �ּ���.');
        frm.eventidArr.focus();
        return;
    }
    
    if (confirm('������� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "regByEventIDarr";
        frm.action="iParkAPI_Process.asp";
        frm.submit();
    }
}

function RegByMakerID(frm){
    if (frm.makeridArr.value.length<1){
        alert('�귣�� ID��  �Է��� �ּ���.');
        frm.makeridArr.focus();
        return;
    }
    
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "regByMakerid";
        frm.action="iParkAPI_Process.asp";
        frm.submit();
    }
}

function RegByRecentSell(frm){
    if (confirm('��� �Ͻðڽ��ϱ�?')){
        frm.mode.value = "recentBestSeller";
        frm.action="iParkAPI_Process.asp";
        frm.submit();
    }
}

///---------------------------------------------------------------------------------






function InterParkRegProcess(){
    if (confirm('�ϰ� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegAll";
        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
}

function InterParkEditProcess(){
    if (confirm('�ϰ� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
}

function MakeInterParkEditFile(){
    if (confirm('���� ������ �ۼ� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditPrd";
        document.frmSvArr.action = "/admin/etc/interparkXML/newRegedItem.asp"
        document.frmSvArr.submit();
    }
}

function MakeInterParkRegFile(){
    if (confirm('��� ������ �ۼ� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "";
        document.frmSvArr.action = "/admin/etc/interparkXML/newRegedItem.asp"
        document.frmSvArr.submit();
    }
    
}

function InterParkDelSoldOutProcess(){
    //return;
    if (confirm('ǰ�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelSoldOut";
        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
    
}

function InterParkDelJaeHyuProcess(){
    //return;
    if (confirm('���޸� �ƴѰ��� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelJaeHyu";
        
        if (document.frmSvArr.jaehyupagegubun.value == "2")
        {
        	document.frmSvArr.jaehyupagegubun.value = "1";
        }
        else
        {
        	document.frmSvArr.jaehyupagegubun.value = "2";
        }

        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
    
}

function MakeInterParkDelFile(){
    //return;
    if (confirm('ǰ�� ��ǰ ���� ������ �ۼ� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelSoldOut";
        document.frmSvArr.action = "/admin/etc/interparkXML/newRegedItem.asp"
        document.frmSvArr.submit();
    }
    
}


//������� ��ǰ üũ �� ���� 
function checkNDelItem(iitemid){
    if (confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.frmDumiArr.target = "xLink";
        document.frmDumiArr.mode.value = "chkNdelitem";
        document.frmDumiArr.cksel.value = iitemid+",";
        document.frmDumiArr.action = "iParkAPI_Process.asp"
        document.frmDumiArr.submit();
    }
}

//������ũ ��ϻ�ǰ ����
function DelTenIparkItem(iitemid){
    if (confirm('���޻翡 ��ǰ�� ��� �Ǿ� ������� ����� ǰ��ó�� �Ͻñ� �ٶ��ϴ�. \n���� �Ͻðڽ��ϱ�?')){
        var popwin = window.open('','iDelTenIparkItem','width=100,height=100');
        
        
        frmDel.mode.value = "delitem";
        frmDel.delitemid.value = iitemid;
        frmDel.target = "iDelTenIparkItem";
        frmDel.submit();
    }
}


//������ũ ǰ�� ó�� �� ���� // not using
function DelIparkItem(iitemid){
    if (confirm('���޻翡 ��� �Ǿ� �ִ� ��ǰ�� ǰ�� ó�� �� ���� �մϴ�. \n���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelPrd";
        document.frmSvArr.delitemid.value = iitemid;
        //document.frmSvArr.action = "/admin/etc/interparkXML/newRegedItem.asp"
        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
}

function EditIParkSupplyCtrtSeq(iitemid){
    var popwin = window.open('EditIParkSupplyCtrtSeq.asp?itemid=' + iitemid,'EditIParkSupplyCtrtSeq','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=interpark","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// ������� ��ǰ
function NotInItemid()
{
	var popwin = window.open('JaehyuMall_Not_In_Itemid.asp?mallgubun=interpark','notinItem','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function category_manager()
{
	window.open('InterparkCategory.asp','category_manager','width=1100,height=700,scrollbars=yes');
}

function BrandUpdate()
{
	document.frmbrand.brandid.value = frm.makerid.value;
	
	if(document.frmbrand.brandid.value == "")
	{
		alert("�귣�带 �Է��ϼ���.");
		frm.makerid.focus();
		return;
	}
	
    if (confirm(''+document.frmbrand.brandid.value+' �귣�� ��� ���� �Ͻðڽ��ϱ�?')){
        document.frmbrand.target = "iframebrandupdate";
        document.frmbrand.action = "/admin/etc/interparkXML/brandupdate.asp"
        document.frmbrand.submit();
    }
}

function InterParkBrandUpdate(){

	document.frmSvArr.brandid.value = frm.makerid.value;
	
	if(document.frmSvArr.brandid.value == "")
	{
		alert("�귣�带 �Է��ϼ���.");
		frm.makerid.focus();
		return;
	}
	
    if (confirm('�ϰ� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "interparkItem_Process.asp"
        document.frmSvArr.submit();
    }
}

function checkComp(comp){
    if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
        if ((comp.name=="bestOrd")&&(comp.checked)){
            comp.form.bestOrdMall.checked=false;
        }
        
        if ((comp.name=="bestOrdMall")&&(comp.checked)){
            comp.form.bestOrd.checked=false;
        }
    }else if ((comp.name=="optAddprcExists")||(comp.name=="optAddprcExistsExcept")){
        if ((comp.name=="optAddprcExists")&&(comp.checked)){
            comp.form.optAddprcExistsExcept.checked=false;
        }
        
        if ((comp.name=="optAddprcExistsExcept")&&(comp.checked)){
            comp.form.optAddprcExists.checked=false;
        }
    }
}

function checkQuickClick(comp){

}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
    		�귣�� :
    		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
    		
    		������ũ��ǰ��ȣ:
    		<input type="text" name="extitemid" value="<%= extitemid %>" size="12" maxlength="10" class="input">
    		��ǰ��:
    		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
    		&nbsp;
    		<a href="http://ipss.interpark.com/member/login.do?_method=initial&GNBLogin=Y&wid1=wgnb&wid2=wel_login&wid3=seller" target="_blank">������ũAdmin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ coolhass | store10x10 ]</font>"
			End If
		%>
    		<br>
    		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
    		<br>
			��ǰ��ȣ: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
    		�̺�Ʈ��ȣ:
    		<input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
			&nbsp;
			�ֹ����ۿ��� :
			<select name="isMadeHand" class="select">
				<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
				<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
				<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
			</select>
    		<br>
    		��Ͽ��� : 
    		<select name="ExtNotReg">
    		<option value="V" <%= CHkIIF(ExtNotReg="V","selected","") %> >��Ͽ��� ���ɻ�ǰ
    		<option value=""  <%= CHkIIF(ExtNotReg="","selected","") %> >��Ͽ����̻�
    		<option value="M" <%= CHkIIF(ExtNotReg="M","selected","") %> >��Ͽ���
    		<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >������ũ ��ϿϷ�
    		<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >������ũ �������
    		</select>
    		&nbsp;
		    <input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %>  onClick="checkComp(this)"><b>����Ʈ��</b>
		    &nbsp;
		    <input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>
    		&nbsp;
    		ī�׸�Ī :
    		<select name="MatchCate">
    		<option value="">��ü
    		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
    		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
    		</select>
    		&nbsp;
    		�Ǹſ��� :
    		<select name="sellyn" class="select">
    		<option value="" <%= CHkIIF(sellyn="","selected","") %> >��ü
    		<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
    		<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
    		</select>
    		&nbsp;
    		���Ͽ��� :
    		<select name="sailyn" class="select">
    		<option value="" <%= CHkIIF(sailyn="","selected","") %> >��ü
    		<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
    		<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
    		</select>
    		&nbsp;
    		<input type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>������</font>��ǰ����
    		&nbsp;
    		�������� :
    		<select name="showminusmagin15" class="select">
    		<option value="" <%= CHkIIF(showminusmagin15="","selected","") %> >��ü
    		<option value="Y" <%= CHkIIF(showminusmagin15="Y","selected","") %> ><%=CMAXMARGIN%>�̻�
    		<option value="N" <%= CHkIIF(showminusmagin15="N","selected","") %> ><%=CMAXMARGIN%>����
    		</select>
    		&nbsp;
    		<input type="checkbox" name="onlysoldout" <%= ChkIIF(onlysoldout="on","checked","") %> ><font color=red>ǰ��</font>��ǰ����
    		&nbsp;
    		<input type="checkbox" name="onlynotusing" <%= ChkIIF(onlynotusing="on","checked","") %> ><font color=red>�������</font>��ǰ����
    		&nbsp;
    		<input type="checkbox" name="availreg" <%= ChkIIF(availreg="on","checked","") %> >��ϰ��ɻ�ǰ����
    		&nbsp;
		    <input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >��ϼ���������ǰ
    		<br>
    		<input type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>������ũ ����<�ٹ����� �ǸŰ�</font>��ǰ����
    		&nbsp;
    		<input type="checkbox" name="interyes10x10no" <%= ChkIIF(interyes10x10no="on","checked","") %> ><font color=red>������ũ�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
    		&nbsp;
    		<input type="checkbox" name="onreginotmapping" <%= ChkIIF(onreginotmapping="on","checked","") %> ><font color=red>registered interpark & not cate mapping</font>��ǰ����
			&nbsp;
			<input type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
    		<br>
    		<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (���޸� �����Ե�)
		    &nbsp;&nbsp;�����ǸŻ��� :
    		<select name="extsellyn" class="select">
    		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
    		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
    		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
    		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
    		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
    		</select>
    		&nbsp;&nbsp;ǰ�������Է¿��� :
    		<select name="infoDivYn" class="select">
    		<option value="" <%= CHkIIF(infoDivYn="","selected","") %> >��ü
    		<option value="Y" <%= CHkIIF(infoDivYn="Y","selected","") %> >�Է�
    		<option value="N" <%= CHkIIF(infoDivYn="N","selected","") %> >���Է�
    		</select>
    		
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmReg" method="post" action="interparkitem.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="cksel" value="">
<tr height="30" bgcolor="#FFFFFF">
    <td>
        ��ǰ�ڵ�� (����)��� &nbsp;&nbsp;&nbsp;&nbsp;: 
        <input class="input" type="text" name="itemidArr" value="" size="60"> <input class="button" type="button" value="���" onclick="RegByItemID(document.frmReg);">(�޸ӷ� ����)
        <br>
        �̺�Ʈ ��ȣ�� (����)��� : <input class="input" type="text" name="eventidArr" value="" size="60"> <input class="button" type="button" value="���" onclick="RegByEventID(frmReg);">(�޸ӷ� ����)
        <br>
        �귣��ID�� (����)��� &nbsp;&nbsp;&nbsp;&nbsp;: 
        <input class="input" type="text" name="makeridArr" value="" size="32" maxlength="32"> <input class="button" type="button" value="���" onclick="RegByMakerID(frmReg);">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
        <tr height="10"><td></td></tr>
        <tr>
        	<td>
        		<input class="button" type="button" value="�ֱ� ����Ʈ ���� ���" onclick="RegByRecentSell(frmReg);">
		        &nbsp;&nbsp;&nbsp;
		        <input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();">
		        &nbsp;
		        <input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
        	</td>
        	<td align="right"><input class="button" type="button" value="InterParkī�װ���Ī" onclick="category_manager();"></td>
        </tr>
        </table>
    </td>
</tr>
<% IF (FALSE) then %>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table class="a">
	    <tr>
	    	<td width="100%">
			    <input class="button" type="button" value="������ ��ǰ ������ũ�� �ϰ����� [70�� ��]" onClick="InterParkEditProcess();">
			    <input type="button" value="." onClick="MakeInterParkEditFile();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" value="�̵�� ��ǰ ������ũ�� �ϰ���� [70�� ��]" onClick="InterParkRegProcess();">
			    <input type="button" value="." onClick="MakeInterParkRegFile();">
			    &nbsp;&nbsp;
			    <br>�� "������ ��ǰ ������ũ�� �ϰ����� [70�� ��]" �� �ڵ�ȭ ó���Ǿ� �ֽ��ϴ�.<p>
			    <input class="button" type="button" value="���޸��ƴѰ� �ϰ����� [30�Ǿ�]" onClick="InterParkDelJaeHyuProcess();"> 0���϶�����
			    <!--
			    &nbsp;&nbsp;
			    <input class="button" type="button" value="����ó����ǰ �ϰ����" onClick="InterParkDelSoldOutProcess();">
			    //-->
			</td>
			<td width="50">
			    <input class="button" type="button" value="" onClick="InterParkDelSoldOutProcess();">
			    <input type="button" value="." onClick="MakeInterParkDelFile();">
			</td>
		</tr>
		</table>
    </td>
</tr>
<% end if %>


<tr >
    <td bgcolor="#FFFFFF">���ű� �޴� :: �ڵ�ȭ ���� �κ� :: �ǵ��(20��), �Ǽ���(ǰ��:5, ����:5, ����:20) /�� 30�д���</td>
</tr>
<tr >
    <td bgcolor="#FFFFFF" height="35">
        <table width="100%" cellpadding="0" cellspacing="0" class="a">
    	    <tr >
    	        <td bgcolor="#AAAA77" width="100" align="center">���� �޴�</td>
    	        <td bgcolor="#FFFFFF" width="10"></td>
    	        <td bgcolor="#FFFFFF" >
    	            <input class="button" type="button" value="���û�ǰ(����)���" onClick="InterParkregIMSI(document.frmSvArr);">
    	            &nbsp;&nbsp;
    	            <input class="button" type="button" value="���û�ǰ(����)����" onClick="InterParkdelIMSI(document.frmSvArr);">
    	        </td>
            </tr>
        </table> 
    </td>
</tr>
<tr >
    <td bgcolor="#FFFFFF" height="35">
        <table width="100%" cellpadding="0" cellspacing="0" class="a">
    	    <tr>
                <td bgcolor="#AAAA77" width="100" align="center">New API</td>
                <td bgcolor="#FFFFFF" width="10"></td>
                <td>
                <input class="button" type="button" value="���û�ǰ �� ���" onClick="InterParkregItemNewAPI(document.frmSvArr);">
                &nbsp;&nbsp;
                <input class="button" type="button" value="���û�ǰ �� ����" onClick="InterParkEditItemNewAPI(document.frmSvArr);">   
                &nbsp;&nbsp;
                <input class="button" type="button" value="���û�ǰ �ǸŻ���Ȯ��" onClick="InterParkSelectStatCheck(document.frmSvArr);">
                <% if session("ssBctID")="icommang" or session("ssBctID")="kjy8517" then %>
                <br>
                <input class="button" type="button" value="���� Auto TEST" onClick="InterParkEditItemAutoNewAPI(document.frmSvArr);">   
                &nbsp;&nbsp;
                <input class="button" type="button" value="��� Auto TEST" onClick="InterParkregItemAutoNewAPI(document.frmSvArr);"> 
                &nbsp;&nbsp;
                <input class="button" type="button" value="���� Auto TEST" onClick="InterParkExpireItemAutoNewAPI(document.frmSvArr);">  
                &nbsp;&nbsp;
                <input class="button" type="button" value="ǰ����Է� Auto TEST" onClick="InterParkInfoFivNoneItemAutoNewAPI(document.frmSvArr);">  
                &nbsp;&nbsp;
                <input class="button" type="button" value="�ǸŻ���Ȯ�� Auto TEST" onClick="InterParkSelectStatCheckBatch(document.frmSvArr);">
                &nbsp;&nbsp;
                <input type="text" name="locNo" value="" size="3">
                <input class="button" type="button" value="�ǸŻ���Ȯ�� Batch" onClick="InterParkItemInfoCheckBatch(document.frmSvArr);">
                <% end if %>
                
                </td>
                <td align="right">
                    ���û�ǰ�� 
    				<Select name="chgSellYn" class="select">
    				<option value="N">ǰ��</option>
    				<% if (True) or (reqExpire="on") then %>
    				<option value="X"  >�Ǹ�����(����)</option><!-- �����ϸ� ���� ���� �� �� ���� -->
    				<% end if %>
    				</Select>(��)��
    				<input class="button" type="button" id="btnSellYn" value="����" onClick="InterParkSellYnProcess(document.frmSvArr,frmReg.chgSellYn.value);">
                </td>
            </tr>
        </table> 
    </td>
</tr>
</form>
</table>
<br>
<!--
<form name="frmbrand" method="post">
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table class="a">
	    <tr>
	    	<td width="100%">
	    		<input type="hidden" name="brandid" value="">
			    <input class="button" type="button" value="�귣�� ���� ������Ʈ" onClick="BrandUpdate();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" value="�귣�� ������ũ�� �ϰ���� [20�� ��]" onClick="InterParkBrandUpdate();">
			    <iframe name="iframebrandupdate" id="iframebrandupdate" width="0" height="0"></iframe>
<%
	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
			    <input type="button" value="aaa" onClick="MakeInterParkDelFile_aaa();">
<script language="javascript">
function MakeInterParkDelFile_aaa(){
    if (confirm('���� ������ �ۼ� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "interparkItem_Process_n.asp"
        //document.frmSvArr.mode.value = "EditPrd";
        //document.frmSvArr.action = "/admin/etc/interparkXML/newRegedItem_n.asp"
        document.frmSvArr.submit();
    }
}
</script>
<%
	End If
%>
			</td>
		</tr>
		</table>
    </td>
</tr>
</form>
-->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="18" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oInterParkitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oInterParkitem.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">��ǰ��ȣ</td>   
	<td >�귣��<br>��ǰ��</td>
	<td width="120">��ǰ���������<br>��ǰ����������</td>
	<td width="120">������ũ�����<br>������ũ����������</td>
	<td width="70">�ǸŰ�</td>
	<td width="70">����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">������ũ<br>���ݹ��Ǹ�</td>
	<td width="50">����</td>
	<td width="70">����<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="40">ǰ��</td>
</tr>
<form name="frmSvArr" method="post" onSubmit="return false;" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="jaehyupagegubun" value="2">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="locNo" value="">
<% for i=0 to oInterParkitem.FResultCount - 1 %>
<input type="hidden" name="xsiteitemno" value="<%=CHKIIF(IsNULL(oInterParkitem.FItemList(i).FExtSiteItemno),"",oInterParkitem.FItemList(i).FExtSiteItemno)%>">
<input type="hidden" name="availexpire" value="<%=CHKIIF(oInterParkitem.FItemList(i).FSellyn<>"Y" and oInterParkitem.FItemList(i).Fisusing="N","1","") %>">
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oInterParkitem.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oInterParkitem.FItemList(i).Fsmallimage %>" width="50"></td>
    <td align="center"><%= oInterParkitem.FItemList(i).FItemID %><br><%= oInterParkitem.FItemList(i).getiParkRegStateName %></td>
    <td><%= oInterParkitem.FItemList(i).FMakerid %> <%= oInterParkitem.FItemList(i).getdeliverytypeName %><br><%= oInterParkitem.FItemList(i).FItemName %></td>
    <td align="center"><%= oInterParkitem.FItemList(i).FRegdate %><br><%= oInterParkitem.FItemList(i).FitemLastupdate %></td>
    <td align="center"><%= oInterParkitem.FItemList(i).FExtRegdate %><br><%= oInterParkitem.FItemList(i).FExtLastUpdate %></td>
    <td align="right">
        <% if oInterParkitem.FItemList(i).FSailYn="Y" then %>
        <strike><%= FormatNumber(oInterParkitem.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oInterParkitem.FItemList(i).FSellcash,0) %></font>
        <% else %>
        <%= FormatNumber(oInterParkitem.FItemList(i).FSellcash,0) %>
        <% end if %>
    </td>
    <td align="center">
        <% if oInterParkitem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oInterParkitem.FItemList(i).Fbuycash/oInterParkitem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oInterParkitem.FItemList(i).IsSoldOut then %>
            <% if oInterParkitem.FItemList(i).FSellyn="N" then %>
            <font color="red">ǰ��</font>
            <% else %>
            <font color="red">�Ͻ�<br>ǰ��</font>
            <% end if %>
        <% end if %>
        
        <% if oInterParkitem.FItemList(i).Fisusing="N" then %>
            <br><font color="blue">�������</font>
        <% end if %>
    </td>

    <td align="center">
	<%
		If oInterParkitem.FItemList(i).FItemdiv = "06" OR oInterParkitem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
    </td>

    <td align="center">
    <% if Not IsNULL(oInterParkitem.FItemList(i).FmayiParkPrice) then %>
        <% if (oInterParkitem.FItemList(i).Fsellcash<>oInterParkitem.FItemList(i).FmayiParkPrice) then %>
        <strong><%= formatNumber(oInterParkitem.FItemList(i).FmayiParkPrice,0) %></strong>
        <% else %>
        <%= formatNumber(oInterParkitem.FItemList(i).FmayiParkPrice,0) %>
        <% end if %>
        <br>
        <% if (oInterParkitem.FItemList(i).FmayiParkSellYn="X") then %>
        <a href="javascript:checkNDelItem('<%= oInterParkitem.FItemList(i).FItemID %>')">
        <% end if %>
        
        <% if (oInterParkitem.FItemList(i).FSellyn<>oInterParkitem.FItemList(i).FmayiParkSellYn) then %>
        <strong><%= oInterParkitem.FItemList(i).FmayiParkSellYn %></strong>
        <% else %>
        <%= oInterParkitem.FItemList(i).FmayiParkSellYn %>
        <% end if %>
        
        <% if (oInterParkitem.FItemList(i).FmayiParkSellYn="X") then %>
        </a>
        <% end if %>
    <% end if %>
    </td>
    <td align="center"><a href="javascript:EditIParkSupplyCtrtSeq('<%= oInterParkitem.FItemList(i).FItemID %>')"><%= oInterParkitem.FItemList(i).GetExtStoreSeqName %>(<%= oInterParkitem.FItemList(i).FExtStoreSeq %>)</a></td>
    <td align="center">
        <a target=_blank href="http://<%= chkIIF((application("Svr_Info")="Dev"),"sptest","www") %>.interpark.com/product/MallDisplay.do?_method=detail&sc.shopNo=0000100000&sc.prdNo=<%= oInterParkitem.FItemList(i).FExtSiteItemno %>"><%= oInterParkitem.FItemList(i).FExtSiteItemno %></a>
        
        <% if IsNULL(oInterParkitem.FItemList(i).FExtSiteItemno) then %>
        <a href="javascript:DelTenIparkItem('<%= oInterParkitem.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
        <% end if %>
    </td>
    <td align="center"><%= oInterParkitem.FItemList(i).Freguserid %></td>
    <td align="center">  <a href="javascript:popManageOptAddPrc('<%=oInterParkitem.FItemList(i).FItemID%>','0');"><%= oInterParkitem.FItemList(i).FoptionCnt %>:<%= oInterParkitem.FItemList(i).FregedOptCnt %></a></td>
    <td align="center"><%= oInterParkitem.FItemList(i).FrctSellCNT %></td>
    <td align="center">
    <% if IsNULL(oInterParkitem.FItemList(i).FExtdispcategory) then %>
    <font color="darkred">��Ī�ȵ�</font><br>
    <% else %>
    <a href="javascript:popItem2CategoryRedirect('<%= oInterParkitem.FItemList(i).FItemID %>');"><%= oInterParkitem.FItemList(i).FExtdispcategory %></a><br>
    <% end if %>
    
    <!-- ����� ī�װ� ������.
    <% if IsNULL(oInterParkitem.FItemList(i).FExtStorecategory) then %>Store X<% end if %>
    -->
    <% ''if NOT IsNULL(oInterParkitem.FItemList(i).FExtSiteItemno) and Not IsNULL(oInterParkitem.FItemList(i).FExtdispcategory) and  Not IsNULL(oInterParkitem.FItemList(i).FExtStorecategory) then %>
    <% if (FALSE) and NOT IsNULL(oInterParkitem.FItemList(i).FExtSiteItemno)  then %>
    <a href="javascript:DelIparkItem('<%= oInterParkitem.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
    <% end if %>
    
    <% if (oInterParkitem.FItemList(i).FaccFailCNT>0) then %>
        <br><font color="red" title="<%= oInterParkitem.FItemList(i).FlastErrStr %>">ERR:<%= oInterParkitem.FItemList(i).FaccFailCNT %></font>
    <% end if %>
    </td>
    <td align="center"><%= oInterParkitem.FItemList(i).FinfoDiv%></td>
</tr>
<% next %>
</form>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oInterParkitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oInterParkitem.StarScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>
    
    	<% for i=0 + oInterParkitem.StarScrollPage to oInterParkitem.FScrollCount + oInterParkitem.StarScrollPage - 1 %>
    		<% if i>oInterParkitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>
    
    	<% if oInterParkitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
set oInterParkitem = Nothing
%>
<form name="frmDel" method="post" action="interparkitem.asp" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
</form>

<form name="frmDumiArr" method="post" action="iParkAPI_Process.asp" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="cksel" value="">
</form>


<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr>
    <td><iframe name="xLink" id="xLink" width="100%" height="300"></iframe></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->