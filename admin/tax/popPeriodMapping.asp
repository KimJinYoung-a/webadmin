<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ���ڰ�꼭 ���� ���⼺ �ڷ� ����
' History : 2012.02.09 ������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->  

<%
Dim autoIcheIdx : autoIcheIdx =  requestCheckvar(request("autoIcheIdx"),10)
Dim sellType    : sellType    =  requestCheckvar(request("sellType"),10)
Dim isocno      : isocno    =  requestCheckvar(request("isocno"),20)
Dim mode        : mode    =  requestCheckvar(request("mode"),20)
Dim page    : page    =  requestCheckvar(request("page"),10)

IF (sellType="") then sellType="0"
IF (page="") then page=1
isocno = replace(isocno,"-","")

Dim clsPMapping
Set clsPMapping = new CEsero 
    clsPMapping.FtaxsellType= sellType  
	clsPMapping.FRectcorpNo = isocno  
	
	clsPMapping.FCurrPage 	= page
	clsPMapping.FPageSize 	= 100
	clsPMapping.fnGetAutoIcheMapDataList 	


Dim i
%>
<script language='javascript'>
function research(autoIcheIdx,mode){
    document.frm.autoIcheIdx.value = autoIcheIdx;
    document.frm.mode.value = mode;
    document.frm.submit();
}

function regPeriodMapping(isreg){
    var frm=document.frmReg;
    
    if (frm.TaxSellType.value==""){
        alert('����/���ⱸ���� �����ϼ���.');
        frm.TaxSellType.focus();
        return;
    }  
    
    if (frm.matchType.value==""){
        alert('���α����� �����ϼ���.');
        frm.matchType.focus();
        return;
    }   
    
    if (frm.autoIcheTitle.value==""){
        alert('���θ�Ī�� �Է��ϼ���.');
        frm.autoIcheTitle.focus();
        return;
    } 
    
    if (frm.corpNo.value==""){
        alert('����ڹ�ȣ�� �Է��ϼ���.');
        frm.corpNo.focus();
        return;
    } 
    
    if (frm.cust_cd.value==""){
        alert('�ŷ�ó �ڵ带 ���� �ϼ���.');
        
        return;
    } 
    
    
    if (frm.matchType.value=="900"){
        //�ڵ���ü
        //if (frm.mayPrice.value==""){
        //    alert('�ݾ��� �Է��ϼ���.');
        //    frm.mayPrice.focus();
        //    return;
        //}
    
        if (frm.mayPumok.value.length<3){
            alert('ǰ���� �Է��ϼ���. 3���̻�');
            frm.mayPumok.focus();
            return;
        }
        
        //if (frm.mayIcheDate.value==""){
        //    alert('��������� �Է��ϼ���.');
        //    return;
        //}
        
        //if (frm.mayAcctJukyo.value==""){
        //    alert('����� ���並 �Է��ϼ���.');
        //    frm.mayAcctJukyo.focus();
        //    return;
        //}
    }else{
        if (frm.mayPumok.value.length<3){
            alert('ǰ���� �Է��ϼ���. 3���̻�');
            frm.mayPumok.focus();
            return;
        }
    }
    //mayIcheDate
    
    if (frm.bizSecCd.value==""){
        alert('����κ��� ���� �ϼ���.');
        return;
    } 
    
    if (frm.arap_cd.value==""){
        alert('�����׸��� ���� �ϼ���.');
        return;
    } 


    var regMn='���';
    if (!isreg)  regMn='����';
    if (confirm(regMn + ' �Ͻðڽ��ϱ�?')){
        frm.mode.value="regPeriod";
        frm.submit();
    }
}

function delPeriodMapping(){
    var frm=document.frmReg;
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value="delPeriod";
        frm.submit();
    }
}

//�ڱݰ����μ� ����
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//�ڱݰ����μ� ���
function jsSetPart(bizSecCd, sPNM){ 
    var frm = document.frmReg;
    frm.bizSecCd.value = bizSecCd;
    frm.AssignBizSecName.value = sPNM;
}

//�����׸� �ҷ�����
function jsGetARAP(){
    var rdoGB = "2"; //����
	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp?rdoGB="+rdoGB,"popARAP1","width=800,height=600,resizable=yes, scrollbars=yes");
	winARAP.focus();
}

//���� �����׸� ��������
function jsSetARAP(dAC, sANM,sACC,sACCNM){ 
    var frm = document.frmReg;
    frm.arap_cd.value = dAC;
    frm.AssignArapNm.value = sANM;
	
}

//�ŷ�ó ���� ����
function jsGetCust(){
	var Strparm="";
	var cust_cd = ""; 
	var rdoCgbn = "2"; //����
	var corpNo = document.frmReg.corpNo.value;
	if (cust_cd!=""){
		Strparm = "?selSTp=1&sSTx="+ cust_cd;
    }else if(corpNo!=""){
        Strparm = "?selSTp=5&sSTx="+ corpNo;
	}else{
	    Strparm = "?rdoCgbn="+rdoCgbn;
	}
	Strparm = Strparm + "&opnType=eTax";
	var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1200, height=600,resizable=yes, scrollbars=yes");
	winC.focus();
}

//�ŷ�ó ����
function jsSetCust(custcd, custnm, ceonm, custno ){
    var frm = document.frmReg;
    frm.cust_cd.value = custcd;
    frm.corpNo.value = custno;
    
}

function CkeckAll(comp,cname){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")&&(e.name==cname)) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkSel(comp){
    AnCheckClick(comp)
}

function popErpSending(itaxkey){
    var winD = window.open("popRegfileHand.asp?taxkey="+itaxkey,"popErpSending","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

function mapPeriod(frm){
    var checkedExists = false;
    var eseroKey="";
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")&&(e.value!="")&&(e.name=="chk")) {
    	    if (e.checked==true){
    	        checkedExists = e.checked;
    	        eseroKey += e.value+",";
    	    }
		}
	}
	
	if (!checkedExists){
	    alert('���� ������ �����ϴ�.');
	    return;
	}
	
	if (confirm('���� ������ ��Ī ó�� �Ͻðڽ��ϱ�?')){
	    
	    frm.mode.value="modiTaxMapping";
	    frm.eseroKey.value=eseroKey;
	    frm.submit();
	}
}

function sendErpArr(frm){
    
    var checkedExists = false;
    var eseroKey="";
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")&&(e.value!="")&&(e.name=="chk2")) {
    	    if (e.checked==true){
    	        checkedExists = e.checked;
    	        eseroKey += e.value+",";
    	    }
		}
	}
	
	if (!checkedExists){
	    alert('���� ������ �����ϴ�.');
	    return;
	}
	//alert(eseroKey);
	
	if (confirm('���������� ERP�� �����Ͻðڽ��ϱ�?')){
        document.frmAct.mode.value="sendDocErp"
        document.frmAct.taxKeyArr.value = eseroKey;
        if (frm.chkPLANDATE.checked==true){
            document.frmAct.chkPLANDATE.value = "on";
        }else{
            document.frmAct.chkPLANDATE.value = "";
        }
        document.frmAct.submit();
    }
    
}

function popHandMapping(iselltype,iaccDt,itaxkey,isocno){
    var popURL = 'popHandMapping.asp?iselltype='+iselltype+'&iaccDt='+iaccDt+'&itaxkey='+itaxkey+'&isocno='+isocno;
    var popwin = window.open(popURL,'popHandMapping','width=1000, height=800, scrollbars=yes, resizable=yes');
	popwin.focus();
}
</script>

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value=""> 
	<input type="hidden" name="mode" value=""> 
	<input type="hidden" name="autoIcheIdx" value=""> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2"  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
		<td align="left"> 
			<input type="radio" name="sellType" value="0" <%= CHKIIF(sellType="0","checked","") %> >���� 
			<input type="radio" name="sellType" value="1" <%= CHKIIF(sellType="1","checked","") %> >����&nbsp;&nbsp;
			
			&nbsp;&nbsp;����ڵ�Ϲ�ȣ:
			<input type="text" name="isocno" value="<%=isocno%>" size="15">
			
			
		</td> 
		<td rowspan="2"  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr  bgcolor="#FFFFFF" >
	    <td >
	    
	    </td>
	</tr>
	</form>
</table>
<p>
<%

''Dim autoIcheIdx   
Dim matchType
Dim TaxSellType   
Dim corpNo        
               
Dim autoIcheTitle 
Dim mayPrice      
Dim mayAcctDate   
Dim mayPumok      
Dim mayIcheDate   
Dim mayAcctJukyo  
Dim bizSecCd  , AssignBizSecName
Dim arap_cd   , AssignArapNm
Dim corpName      
Dim cust_cd

Dim clsPOneMap
IF (autoIcheIdx<>"") then
    Set clsPOneMap = new CEsero 
    clsPOneMap.FRectautoIcheIdx= autoIcheIdx  
	clsPOneMap.fnGetAutoIcheMapOne 	
	IF (clsPOneMap.FResultCount>0) then
	    autoIcheIdx = clsPOneMap.FOneItem.FautoIcheIdx
	    matchType   = clsPOneMap.FOneItem.FmatchType
	    TaxSellType = clsPOneMap.FOneItem.FTaxSellType
	    corpNo      = clsPOneMap.FOneItem.FcorpNo
	    autoIcheTitle = clsPOneMap.FOneItem.FautoIcheTitle
        mayPrice      = clsPOneMap.FOneItem.FmayPrice
        mayAcctDate   = clsPOneMap.FOneItem.FmayAcctDate
        mayPumok      = clsPOneMap.FOneItem.FmayPumok
        mayIcheDate   = clsPOneMap.FOneItem.FmayIcheDate
        mayAcctJukyo  = clsPOneMap.FOneItem.FmayAcctJukyo
        bizSecCd      = clsPOneMap.FOneItem.FAssignBizSec
        arap_cd       = clsPOneMap.FOneItem.FAssignarap_cd
        AssignBizSecName = clsPOneMap.FOneItem.FAssignBizSecName
        AssignArapNm     = clsPOneMap.FOneItem.FAssignArapNm
        corpName         = clsPOneMap.FOneItem.FcorpName
        cust_cd          = clsPOneMap.FOneItem.Fcust_cd
    else
        autoIcheIdx =""
	end if
    Set clsPOneMap = Nothing
    
    
End IF
%>
<% IF ((mode="") or (autoIcheIdx="") or (mode="mapping")) and (mode<>"reg") THEN %>

<% ELSE %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmReg" method="post" action="eTax_Process.asp">
<input type="hidden" name="autoIcheIdx" value="<%= autoIcheIdx %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="FFFFFF">
    <td>
        <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <td rowspan="2">����/����</td>
            <td rowspan="2">���α���</td>
            <td rowspan="2">���θ�Ī</td>
            <td colspan="5">
                ��꼭����
            </td>
            <td colspan="2">
                ���������
            </td>
            <td colspan="2">
                ��Ī����
            </td>
        </tr>
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <td>����ڹ�ȣ</td>
            <td>�ŷ�ó�ڵ�</td>
            <td>�ݾ�</td>
            <td>��¥</td>
            <td>ǰ��</td>
            <td>��ü��</td>
            <td>����</td>
            <td>����κ�</td>
            <td>�����׸�</td>
        </tr>
        <tr bgcolor="FFFFFF" align="center"> 
        
            <td>
                <select Name="TaxSellType">
		        <option value="">����
		        <option value="0" <%= CHKIIF(TaxSellType="0","selected","") %> >����
		        <option value="1" <%= CHKIIF(TaxSellType="1","selected","") %> >����
		        </select>
            </td>
            <td>
                <select Name="matchType">
		        <option value="">����
		        <option value="900" <%= CHKIIF(matchType="900","selected","") %> >�ڵ���ü
		        <option value="910" <%= CHKIIF(matchType="910","selected","") %> >��Ÿ���
		        </select>
            </td>
            
            <td><input type="text" name="autoIcheTitle" value="<%=autoIcheTitle%>" size="20" maxlength="30"></td>
            <td><input type="text" name="corpNo" value="<%=corpNo%>" size="10" maxlength="10" readonly class="text_ro"></td>
            <td><input type="text" name="cust_cd" value="<%=cust_cd%>" size="10" maxlength="10" readonly class="text_ro">
            <img src="/images/icon_search.jpg" onClick="jsGetCust();" style="cursor:pointer"> 
            </td>
            <td><input type="text" name="mayPrice" value="<%=mayPrice%>" size="10" maxlength="10" style="text-align=right"></td>
            <td><input type="text" name="mayAcctDate" value="<%=mayAcctDate%>" size="2" maxlength="2"></td>
            <td><input type="text" name="mayPumok" value="<%=mayPumok%>" size="10" maxlength="20"></td>
            <td><input type="text" name="mayIcheDate" value="<%=mayIcheDate%>" size="2" maxlength="2"></td>
            <td><input type="text" name="mayAcctJukyo" value="<%=mayAcctJukyo%>" size="10" maxlength="20"></td>
            <td>
                <input type="text" name="AssignBizSecName" value="<%=AssignBizSecName%>" size="10" readonly style="border=0">
                <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer"> 
                <input type="hidden" name="bizSecCd" value="<%= bizSecCd %>">
            </td>
            <td>
                <input type="text" name="AssignArapNm" value="<%=AssignArapNm%>" size="10" readonly style="border=0">  
                <img src="/images/icon_search.jpg" onClick="jsGetARAP();" style="cursor:pointer"> 
                <input type="hidden" name="arap_cd" value="<%= arap_cd %>">
            </td>
        </tr>
        </table>
    </td>
</tr>
<tr bgcolor="FFFFFF">
    <td >
    * �ڵ���ü�ΰ�� ��꼭 �ݾ� �� ǰ���� �ʼ� / �׿��� ��� ǰ���� �ʼ� (3���̻�)
    <br>
    * ��¥�� ������ ��� (31)�� �Է�
    </td>
</tr>
<tr bgcolor="FFFFFF">
    <td align="center">
        <% if (CStr(autoIcheIdx)<>"") then %>
        <input type="button" value="����" onclick="regPeriodMapping(false)">
        &nbsp;&nbsp;
        <input type="button" value="����" onclick="delPeriodMapping()">
        <% else %>
        <input type="button" value="�űԵ��" onclick="regPeriodMapping(true)">
        <% end if %>
    </td>
</tr>
</form>
</table>
<p>
<% ENd IF %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBuf" method="get" action="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		�˻���� : <b><%=clsPMapping.FTotCnt%></b> &nbsp;
	</td>
	<td colspan="12" align="right">
	    <% IF (mode<>"reg") THEN %>
	    <input type="button" value="�űԵ��" onClick="research('','reg');">
	    <% end if %>
    </td>
    </form>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
    <td rowspan="2">�˻�</td>
    <td rowspan="2">����/����</td>
    <td rowspan="2">���α���</td>
    <td rowspan="2">���θ�Ī</td>
    <td colspan="5">
        ��꼭����
    </td>
    <td colspan="2">
        ���������
    </td>
    <td colspan="2">
        ��Ī����
    </td>
    <td rowspan="2">
        ����
    </td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
    <td>����ڹ�ȣ<br>�ŷ�ó�ڵ�</td>
    <td>�ŷ�ó��</td>
    <td>�ݾ�</td>
    <td>��¥</td>
    <td>ǰ��</td>
    <td>��ü��</td>
    <td>����</td>
    <td>����κ�</td>
    <td>�����׸�</td>
</tr>
<%  
IF clsPMapping.FResultCount>0 then
	For i = 0 To clsPMapping.FResultCount-1
	%> 
<tr align="center" bgcolor="#FFFFFF">
    <input type="hidden" name="socno" value="<%= clsPMapping.FItemList(i).FcorpNo %>">
    <td><img src="/images/icon_search.jpg" onClick="research('<%= clsPMapping.FItemList(i).FautoIcheIdx %>','mapping');" style="cursor:pointer"> </td>
    <td><%= getSellTypeName(clsPMapping.FItemList(i).FTaxSellType)%></td>
    <td><%= getMatchTypeName(clsPMapping.FItemList(i).FmatchType)%></td>
    <td><%= clsPMapping.FItemList(i).FautoIcheTitle%></td>
    <td><%= clsPMapping.FItemList(i).FcorpNo%><br>(<%= clsPMapping.FItemList(i).Fcust_cd%>)</td>
    <td><%= clsPMapping.FItemList(i).FcorpName%></td>
    <td><%= clsPMapping.FItemList(i).FmayPrice%></td>
    <td><%= clsPMapping.FItemList(i).FmayAcctDate%></td>
    <td><%= clsPMapping.FItemList(i).FmayPumok%></td>
    <td><%= clsPMapping.FItemList(i).FmayIcheDate%></td>
    <td><%= clsPMapping.FItemList(i).FmayAcctJukyo%></td>
    <td><%= clsPMapping.FItemList(i).FAssignBizSecName%></td>
    <td><%= clsPMapping.FItemList(i).FAssignArapNm%></td>
    <td><input type="button" value="����" onClick="research('<%= clsPMapping.FItemList(i).FautoIcheIdx %>','edit');"></td>
</tr>	
<%	Next
ELSE%>
<tr height=30 align="center" bgcolor="#FFFFFF">				
	<td colspan="19">�˻� ������ �����ϴ�.</td>	
</tr>
<%END IF%>
</table>

<% IF (mode="mapping") then %>
<%
Dim clsEsero, arrList, intLoop, TotCnt
set clsEsero = new CEsero
clsEsero.FCurrPage=1
clsEsero.FPageSize=100
clsEsero.FSDate=Left(dateAdd("m",-2,now()),7)+"-01"
''clsEsero.FEDate
clsEsero.FtaxsellType = TaxSellType
clsEsero.FRectCorpNo  = CorpNo
clsEsero.FsearchText = mayPumok  '''FRectDtlName
clsEsero.FTotSum     = mayPrice

arrList = clsEsero.fnGetEseroTaxList
TotCnt = clsEsero.FTotCnt

set clsEsero= Nothing

%>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEsero" method="post" action="eTax_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="eseroKey" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2">
		�˻���� : <b><%=TotCnt%></b> &nbsp;
	</td>
	<td align="right" colspan="2"> 
	<% if isPLAN_DATEDefaultSend(matchType, taxSellType, arap_cd) then %>
    <input type="checkbox" name="chkPLANDATE" value="" checked >(����/����)���������Է�
    <% else %>
    <input type="checkbox" name="chkPLANDATE" value=""  >(����/����)���������Է�
    <% end if %>
                
    <input type="button" value="�ϰ�����" onClick="sendErpArr(frmEsero)">
    </td>
	<td colspan="12" align="right">
	    <input type="hidden" name="matchType" value="<%= matchType %>">
	    ��ĪŸ�� : <%= getMatchTypeName(matchType) %>
	    &nbsp;
	    �ŷ�ó�ڵ� : <input type="text" name="cust_cd" value="<%= cust_cd %>" size="8" class="text_ro">
	    &nbsp;
	    ����ι�<input type="text" name="bizSecCd_nm" value="<%= AssignBizSecName %>" size="16" class="text_ro"><input type="hidden" name="bizSecCd" value="<%= bizSecCd %>" >
	    <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer">
	    &nbsp;
	    �����׸�<input type="text" name="arap_cd_nm" value="<%= AssignArapNm %>" size="20" class="text_ro"><input type="hidden" name="arap_cd" value="<%= arap_cd %>" >
	    <img src="/images/icon_search.jpg" onClick="jsGetARAP();" style="cursor:pointer">
	    <input type="button" value="�ϰ�����" onClick="mapPeriod(frmEsero);">
	    
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="40"><input type="checkbox" name="chkALL" value="" onClick="CkeckAll(this,'chk');"><br>(����)</td>
    <td rowspan="2" width="40"><input type="checkbox" name="chkALL2" value="" onClick="CkeckAll(this,'chk2');"><br>(����)</td>
	<td rowspan="2">��꼭<br>�ۼ�����</td>
	<td rowspan="2">���ι�ȣ</td>
	
	<td colspan="2"><%IF TaxSellType="0" THEN%>������<%ELSE%>���޹޴���<%END IF%></td> 
	<td rowspan="2">�հ�ݾ�</td>  
	<td rowspan="2">���ް���</td> 	
	<td rowspan="2">����</td> 
	<td rowspan="2">�з�</td> 
	<td rowspan="2">����</td>  
	<td rowspan="2">ǰ���</td>  
	<td rowspan="2">����<br>Ÿ��</td> 
	<td rowspan="2">����ι�</td> 
	<td rowspan="2">ERP<br>���ۻ���</td> 
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 	
	<td>����ڵ�Ϲ�ȣ</td>
	<!-- td>��</td -->
	<td>��ȣ</td>
</tr>

<%  
IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%> 
<tr align="center" bgcolor="#FFFFFF">
    <td>
        <% if IsNULL(arrList(29,intLoop)) THEN %>
        <input type="checkbox" name="chk" value="<%= arrList(0,intLoop) %>" onClick="checkSel(this);">
        <% else %>
        <input type="checkbox" name="chk" value="<%= arrList(0,intLoop) %>" disabled >
        <% end if %>
    </td>
    <td>
        <% if IsNULL(arrList(33,intLoop)) and (Not IsNULL(arrList(29,intLoop))) and (Not IsNULL(arrList(32,intLoop))) and (Not IsNULL(arrList(38,intLoop))) THEN %>
        <input type="checkbox" name="chk2" value="<%= arrList(0,intLoop) %>" onClick="checkSel(this);">
        <% else %>
        <input type="checkbox" name="chk2" value="<%= arrList(0,intLoop) %>" disabled >
        <% end if %>
    </td>
    <td><%= arrList(1,intLoop) %></td>
    <td><a href="javascript:popErpSending('<%= arrList(0,intLoop) %>')"><%= arrList(0,intLoop) %></a></td>
    <% if arrList(15,intLoop)=1 then %>
    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(7,intLoop) %>')"><%= arrList(7,intLoop) %></a></td>
    <td><%= arrList(9,intLoop) %></td>
    <% else %>
    <td><a href="javascript:popHandMapping('<%= arrList(15,intLoop) %>','<%= arrList(1,intLoop) %>','<%= arrList(0,intLoop) %>','<%= arrList(2,intLoop) %>')"><%= arrList(2,intLoop) %></a></td>
    <td><%= arrList(4,intLoop) %></td>
    <% end if %>
    <td align="right"><%= FormatNumber(arrList(12,intLoop),0) %></td>
    <td align="right"><%= FormatNumber(arrList(13,intLoop),0) %></td>
    <td align="right"><%= FormatNumber(arrList(14,intLoop),0) %></td>
    <td><%= getSellTypeName(arrList(15,intLoop)) %></td>
    <td><%= gettaxModiTypeName(arrList(16,intLoop)) %>/<%= gettaxTypeName(arrList(17,intLoop)) %></td>
    <td><%= arrList(22,intLoop) %></td>
    <td><%= getMatchTypeName(arrList(29,intLoop)) %></td>
    <td><%= getbizSecCdName(arrList(32,intLoop)) %>
        <% if arrList(35,intLoop)>0 then %>
        �� <%= arrList(35,intLoop) %>
        <% end if %>
    </td>
    <td>
        <% if Not IsNULL(arrList(33,intLoop)) then %>
	    [<%= arrList(33,intLoop) %>]
	    <%= arrList(34,intLoop) %>    
        <% end if %>
    </td>
    
</tr>	
<%	Next %>
<% end if %>
</form>
</table>

<% end if %>
<%
Set clsPMapping = nothing	
%>
<form name="frmAct" method="post" action="eTax_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="taxKey" value="">
<input type="hidden" name="bizSecCd" value="">
<input type="hidden" name="arap_cd" value="">
<input type="hidden" name="matchSeq" value="">
<input type="hidden" name="chkPLANDATE" value="">
<input type="hidden" name="taxKeyArr" value="">
</form>

<!-- #include virtual="/lib/db/dbclose.asp" --> 
<!-- #include virtual="/admin/lib/poptail.asp"-->