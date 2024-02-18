<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/BizProfitCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim page : page=requestCheckvar(request("page"),10)
Dim yyyymm : yyyymm=requestCheckvar(request("yyyymm"),7)
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd : accusecd=requestCheckvar(request("accusecd"),16)
Dim cust_cd  : cust_cd=requestCheckvar(request("cust_cd"),10)
Dim mode     : mode=requestCheckvar(request("mode"),10)
Dim divMastKey : divMastKey=requestCheckvar(request("divMastKey"),10)

Dim regbizSecCd : regbizSecCd=requestCheckvar(request("regbizSecCd"),16)
Dim regaccusecd : regaccusecd=requestCheckvar(request("regaccusecd"),16)
Dim regcust_cd  : regcust_cd=requestCheckvar(request("regcust_cd"),10)
Dim regbizSecNM : regbizSecNM=requestCheckvar(request("regbizSecNM"),32)

dim i, j, intLoop

Dim AssignYYYYMM
Dim AssignBizSecName, AssignbizSecCd
Dim AssignAccUseCD, Assigncust_cd

Dim oBizProfitOne
set oBizProfitOne = new CBizProfit
oBizProfitOne.FRectdivMastKey = divMastKey

IF (oBizProfitOne.FRectdivMastKey<>"") then
    oBizProfitOne.getOneBizProfitDivMaster
elseif (mode="searchby") then
    oBizProfitOne.FRECTYYYYMM   = yyyymm
    oBizProfitOne.FRectBizSecCD = regbizSecCd
    oBizProfitOne.FRectCustCD   = regcust_cd
    oBizProfitOne.FRectAccUseCd = regaccusecd
    
    if (yyyymm<>"") and (regbizSecCd<>"") and (regcust_cd<>"" or regaccusecd<>"") then
        oBizProfitOne.getOneBizProfitDivMasterBySearch
        if (oBizProfitOne.FResultCount>0) then
            divMastKey = oBizProfitOne.FOneItem.FdivMastKey
            mode =""
        else
            mode ="reg"
        end if
    end if
    
end if


'rw "divMastKey="&divMastKey


Dim oBizDivMaster
set oBizDivMaster = new CBizProfit
oBizDivMaster.FRECTYYYYMM = yyyymm
oBizDivMaster.FRectBizSecCD = bizSecCd
oBizDivMaster.FRectAccUseCd = accusecd
oBizDivMaster.getBizProfitDivMasterList

Dim oBizDivDetail
set oBizDivDetail = new CBizProfit
oBizDivDetail.FRECTdivMastKey = divMastKey
oBizDivDetail.getBizProfitDivDetail


''����ι�
Dim clsBS, arrBizList
Set clsBS = new CBizSection 
	clsBS.FUSE_YN = "Y"  
	clsBS.FOnlySub = "Y"  
	arrBizList = clsBS.fnGetBizSectionList  
Set clsBS = nothing



if (oBizProfitOne.FResultCount>0) then
    AssignYYYYMM = oBizProfitOne.FOneItem.FYYYYMM
    AssignBizSecName = oBizProfitOne.FOneItem.FpBIZSECTION_NM
    AssignbizSecCd = oBizProfitOne.FOneItem.FpBIZSECTION_CD
    AssignAccUseCD = oBizProfitOne.FOneItem.FpACC_USE_CD
    Assigncust_cd = oBizProfitOne.FOneItem.FpCUST_CD
end if
set oBizProfitOne = Nothing

IF (AssignYYYYMM="") then AssignYYYYMM=yyyymm
IF (AssignBizSecName="") then AssignBizSecName=regbizSecNM
IF (AssignbizSecCd="") then AssignbizSecCd=regbizSecCd
IF (AssignAccUseCD="") then AssignAccUseCD=regaccusecd
IF (Assigncust_cd="") then Assigncust_cd=regcust_cd
 
IF (AssignBizSecName="") then AssignBizSecName="����Ⱥ�"
IF (AssignbizSecCd="") then AssignbizSecCd="0000009010"
   

''����� �󼼳���
Dim oBizProfit
if (divMastKey<>"") and (mode="") then
    
    set oBizProfit = new CBizProfit
    oBizProfit.FPageSize = 1000
    oBizProfit.FCurrPage = 1
    oBizProfit.FRectBizSecCD=AssignbizSecCd
    oBizProfit.FRectStdt = REplace(yyyymm,"-","")+"01"
    oBizProfit.FRectEddt = Replace(dateAdd("d",-1,DAteAdd("m",1,Left(yyyymm,4)+"-"+Right(yyyymm,2)+"-01")),"-","")
    oBizProfit.FRectAccUseCd = AssignAccUseCD
    oBizProfit.FRectCUSTCD = Assigncust_cd
    'oBizProfit.FRectSANSTS = SANSTS
    'oBizProfit.FRectINTRANS = isINTrans
    'oBizProfit.FRectdivAssign = divAssign
    'oBizProfit.FRectdivdpType = divdpType
    oBizProfit.getBizProfitList
    
end if

Dim tmpsum
Dim odived
%>
<script language='javascript'>
function research(divMastKey,mode){
    document.frm.divMastKey.value = divMastKey;
    document.frm.mode.value = mode;
    document.frm.submit();
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


//�ŷ�ó ���� ����
var pfrmName = '';

function jsGetCust(ifrmName){
    pfrmName = ifrmName;
	var Strparm = "";
	var cust_cd = ""; 
	var rdoCgbn = "";
	var corpNo  = "";
	
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
    var frm = eval("document."+pfrmName);
    frm.cust_cd.value = custcd;
}    

//
function regDivMast(isreg){
    var frm=document.frmReg;
    
    if (frm.AssignYYYYMM.value==""){
        alert('���� ��/���� �����ϼ���.');
        frm.AssignYYYYMM.focus();
        return;
    }   
    
    if (frm.bizSecCd.value==""){
        alert('����ι��� �ʼ� �Դϴ�.');
        //frm.bizSecCd.focus();
        return;
    } 
    
    if (frm.bizSecCd.value*1<9010){
        alert('�Ⱥ� ����ι��� ����Ⱥ� �μ��� ����.');
        return;
    }
    
    if ((frm.cust_cd.value=="")&&(frm.AssignAccUseCD.value=="")){
        alert('�ŷ�ó �ڵ� �Ǵ� ���������ڵ�� �ʼ��Դϴ�.');
        return;
    } 

    if (frm.tmpsum.value*1!=100){
        alert('�Ⱥбݾ� �հ谡 100%�� �ƴմϴ� 100%�� ���߱� �ٶ��ϴ�.');
        return;
    }
    
    var regMn='���';
    if (!isreg)  regMn='����';
    if (confirm(regMn + ' �Ͻðڽ��ϱ�?')){
        frm.mode.value="regDivMast";
        frm.submit();
    }
}

function delDivMast(){
    var frm=document.frmReg;
    if (confirm('���� �Ͻðڽ��ϱ�? ����� �Ⱥе����Ͱ� ������� �������� �ʽ��ϴ�.')){
        frm.mode.value="delDivMast";
        frm.submit();
    }
}

function recalcuSubtotal(comp){
    var frm = comp.form;
    valsum = 0.00;
    for (i=0;i<frm.dtl_dlvPro.length;i++){
        valsum += frm.dtl_dlvPro[i].value*1.00;
    }
    
    valsum = (Math.round(valsum*100)) / 100;
    frm.tmpsum.value=valsum;
}

function showDiv(iSLTRKEY,iSLTRKEY_SEQ){
    var itr = document.getElementById("itr_"+iSLTRKEY+iSLTRKEY_SEQ);
    
    if (itr.style.display=="none"){
        itr.style.display="block";
    }else{
        itr.style.display="none";
    }
}

function delDivAssigned(iSLTRKEY,iSLTRKEY_SEQ){
    if (confirm('�Ⱥ� ������ ���� �Ͻðڽ��ϱ�?')){
        frmAct.mode.value="DelAssignDiv";
        frmAct.SLTRKEY.value = iSLTRKEY;
        frmAct.SLTRKEY_SEQ.value = iSLTRKEY_SEQ;
        
        frmAct.submit();
    }
}

function checkALL(comp){
    var frm = comp.form;
    
    if (frm.chk.length){
        for (i=0;i<frm.chk.length;i++){
            if (!frm.chk[i].disabled){
                frm.chk[i].checked = comp.checked;
            }
        }
    }else if(frm.chk){
        if (!frm.chk.disabled){
            frm.chk.checked = comp.checked;
        }
    }
}

function assignDiv(comp){
    var frm = comp.form;
    var chkExists = false;
    
    if (frm.chk.length){
        for (i=0;i<frm.chk.length;i++){
            chkExists = (chkExists||frm.chk[i].checked);
        }
    }else if(frm.chk){
        chkExists = frm.chk.checked;
    }
    
    if (!chkExists){
        alert('���� ������ �����ϴ�.');
        return;
    }
    
    if (confirm('���� ������ �Ⱥ� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="assignDiv";
        frm.submit();
    }
    
}

//������ ���� ������
	function jsLastGetReg(){  
		document.frmAct.mode.value = "regPreM";
    document.frmAct.submit();
	}

</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value=""> 
	<input type="hidden" name="mode" value=""> 
	<input type="hidden" name="divMastKey" value=""> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2"  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
		<td align="left"> 
			
			�����/��:
			<% CaLL DrawYYYYMMsimpleBox("yyyymm",yyyymm,"onChange=''") %>
					
			&nbsp;&nbsp;
			����ι�:
            <select name="bizSecCd">
            <option value="">--����--</option>
            <% For intLoop = 0 To UBound(arrBizList,2)	%>
        		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(bizSecCd) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
        	<% Next %>
            </select>
    
            &nbsp;&nbsp;
            
			
		</td> 
		<td rowspan="2"  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr  bgcolor="#FFFFFF" >
	    <td >
	        ���������ڵ�:
			<input type="text" name="accusecd" value="<%=accusecd%>" size="15">
			&nbsp;&nbsp;
			�ŷ�ó�ڵ�:
			<input type="text" name="cust_cd" value="<%=cust_cd%>" size="10" maxlength="10" >
            <img src="/images/icon_search.jpg" onClick="jsGetCust('frm');" style="cursor:pointer"> 
	    </td>
	</tr>
	</form>
</table>
<p>

<% IF ((mode="") or (divMastKey="")) and (mode<>"reg") THEN %>

<% ELSE %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmReg" method="post" action="bizProfit_Process.asp">
<input type="hidden" name="AssignYYYYMM" value="<%= yyyymm %>">
<input type="hidden" name="divMastKey" value="<%= divMastKey %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="FFFFFF">
    <td>
        <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <% if (divMastKey<>"") then %>
            <td >�Ⱥй�ȣ</td>
            <% end if %>
            <td >���� ��/��</td>
            <td >����ι�</td>
            <td >���������ڵ�</td>
            <td >�ŷ�ó</td>
        </tr>
        <tr bgcolor="FFFFFF" align="center"> 
            <% if (divMastKey<>"") then %>
            <td ><%= divMastKey %></td>
            <% end if %>
            <td>
                <% CaLL DrawYYYYMMsimpleBox("AssignYYYYMM",AssignYYYYMM,"onChange=''") %>    
            </td>
            <td>
                <input type="text" name="AssignBizSecName" value="<%=AssignBizSecName%>" size="10" readonly style="border=0">
                <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer"> 
                <input type="hidden" name="bizSecCd" value="<%= AssignbizSecCd %>">
            </td>
            <td><input type="text" name="AssignAccUseCD" value="<%=AssignAccUseCD%>" size="10" ></td>
            <td>
                <input type="text" name="cust_cd" value="<%=Assigncust_cd%>" size="10" maxlength="10" readonly class="text_ro">
                <img src="/images/icon_search.jpg" onClick="jsGetCust('frmReg');" style="cursor:pointer"> 
            </td>
        </tr>
        </table>
    </td>
</tr>
<tr>
    <td bgcolor="#FFFFFF">
        <table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <% For intLoop = 0 To oBizDivDetail.FREsultCount-1	%>
                <% if ((oBizDivDetail.FItemList(intLoop).FBIZSECTION_CD)) then %>
                <td>
                <%=oBizDivDetail.FItemList(intLoop).FBIZSECTION_NM%>
                </td>
                <% end if %>
        	<% Next %>
            <td>�հ�</td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <% For intLoop = 0 To oBizDivDetail.FREsultCount-1	%>
                <% if ((oBizDivDetail.FItemList(intLoop).FBIZSECTION_CD)) then %>
                <% tmpsum = tmpsum + NULL2Zero(oBizDivDetail.FItemList(intLoop).FdivPro) %>
                <td>
                <input type="hidden" name="dtl_bizSecCd" value="<%=oBizDivDetail.FItemList(intLoop).FBIZSECTION_CD %>">
                <input type="text" name="dtl_dlvPro" value="<%=oBizDivDetail.FItemList(intLoop).FdivPro %>" size="4" onKeyUp="recalcuSubtotal(this);">
                </td>
                <% end if %>
        	<% Next %>
        	<td><input type="text" name="tmpsum" value="<%= tmpsum %>" readonly class="text_ro" size="4"></td>
        </tr>
        </table>
    </td>
</tr>
<tr bgcolor="FFFFFF">
    <td >
    * ����ι� �ʼ�/ �������� �Ǵ� �ŷ�ó �ʼ�
    </td>
</tr>
<tr bgcolor="FFFFFF">
    <td align="center">
        <% if (CStr(divMastKey)<>"") then %>
        <!-- ���� ����/ ���� �Ǵ� ���.
        <input type="button" value="�Ⱥб��� ����" onclick="regDivMast(false)">
        &nbsp;&nbsp;
        -->
        <input type="button" value="�Ⱥб��� ����" onclick="delDivMast()">
        <% else %>
        <input type="button" value="�Ⱥб��� �űԵ��" onclick="regDivMast(true)">
        <% end if %>
    </td>
</tr>
</form>
</table>
<p>
<% ENd IF %>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="15">* �Ⱥб��� ��� ����</td>
</tr> 
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		�˻���� : <b><%=oBizDivMaster.FTotalcount%></b> &nbsp;
	</td>
	<td colspan="12" align="right">
	    <% IF (mode<>"reg") THEN %>
	    <input type="button" class="button" value="������ ���� ��������" onClick="jsLastGetReg()">
	    <input type="button"  class="button" value="�Ⱥб��� �űԵ��" onClick="research('','reg');">
	    <% end if %>
    </td> 
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
    <td width="30">�˻�</td>
    <td width="60">�Ⱥ��ڵ�</td>
    <td >����ι�</td>
    <td >��������</td>
    <td >�����ڵ�</td>
    <td >�ŷ�ó</td>
    <td >�ŷ�ó�ڵ�</td>
    <td width="100"></td>
</tr>
<% IF oBizDivMaster.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
	    <td colspan="10" align="center">�˻������ �����ϴ�.</td>
    </tr>
<% else %>
    <% For i = 0 To oBizDivMaster.FResultCount-1 %> 
	<tr align="center" bgcolor="<%= CHKIIF(CStr(oBizDivMaster.FItemList(i).FdivMastKey)=CStr(divMastKey),"#C7EEC7","#FFFFFF") %>">
	    <td><img src="/images/icon_search.jpg" onClick="research('<%= oBizDivMaster.FItemList(i).FdivMastKey %>','');" style="cursor:pointer"> </td>
	    <td><%= oBizDivMaster.FItemList(i).FdivMastKey %></td>
	    <td><%= oBizDivMaster.FItemList(i).FpBIZSECTION_NM %></td>
	    <td><%= oBizDivMaster.FItemList(i).FpACC_NM %></td>
	    <td><%= oBizDivMaster.FItemList(i).FpACC_USE_CD %></td>
	    <td><%= oBizDivMaster.FItemList(i).FpCUST_NM %></td>
	    <td><%= oBizDivMaster.FItemList(i).FpCUST_CD %></td>
	    <td><input type="button" value="����" onClick="research('<%= oBizDivMaster.FItemList(i).FdivMastKey %>','edit');"></td>
    </tr>
    <% if CStr(oBizDivMaster.FItemList(i).FdivMastKey)=CStr(divMastKey) then %>
    <tr align="center" >
        <td colspan="8" align="left" bgcolor="#FFFFFF">
        <table align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <% For intLoop = 0 To oBizDivDetail.FREsultCount-1	%>
                <% if (Not IsNULL(oBizDivDetail.FItemList(intLoop).FdivPro)) then %>
                <td width="100">
                <%=oBizDivDetail.FItemList(intLoop).FBIZSECTION_NM%>
                </td>
                <% end if %>
        	<% Next %>
        </tr>
        <tr bgcolor="#FFFFFF">
            <% For intLoop = 0 To oBizDivDetail.FREsultCount-1	%>
                <% if (Not IsNULL(oBizDivDetail.FItemList(intLoop).FdivPro)) then %>
                <% tmpsum = tmpsum + NULL2Zero(oBizDivDetail.FItemList(intLoop).FdivPro) %>
                <td align="center">
                <%=oBizDivDetail.FItemList(intLoop).FdivPro %>
                </td>
                <% end if %>
        	<% Next %>
        </tr>
        </table>
        </td>
    </tr>
    <% end if %>
    <% next %>
<% end if %>
</table>


<p><br/>
<% if (divMastKey<>"") and (mode="") then %>
<% dim debitSum, creditSum, ix %>
<form name="frmDtl" method="post" action="bizProfit_Process.asp" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="AssignYYYYMM" value="<%= yyyymm %>">
<input type="hidden" name="divMastKey" value="<%= divMastKey %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			�˻���� : <b><%=oBizProfit.FTotalCount%></b> &nbsp;
		</td>
		<td colspan="12" align="right">
		    <input type="button" value="���ó��� �Ⱥ�����" onClick="assignDiv(this);">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td ><input type="checkbox" name="chkALL" onClick="checkALL(this)":></td>
	    <td >��ǥ����</td>
		<td >����ι�</td>
		<td >����</td>
		<td >�����ڵ�</td>
		<td >�����з�</td>
		<td >��������</td>
		<td >��������</td>
		<td >�ŷ�ó</td>
		<td >����</td>
		<td >�뺯</td>
		<td >���</td>
		<td >�Ⱥ�<br>����</td>
		<td >�Ⱥ�<br>����</td>
	</tr>
	<% if oBizProfit.FResultCount>0 then %>
	<% For i = 0 To oBizProfit.FResultCount-1 %>
	<%
	debitSum    = debitSum + oBizProfit.FItemList(i).FdebitSum
	creditSum   = creditSum + oBizProfit.FItemList(i).FcreditSum
	%>
	<input type="hidden" name="SLTRKEY" value="<%= oBizProfit.FItemList(i).FSLTRKEY %>">
	<input type="hidden" name="SLTRKEY_SEQ" value="<%= oBizProfit.FItemList(i).FSLTRKEY_SEQ %>">
	<tr align="center" bgcolor="<%= CHKIIF(oBizProfit.FItemList(i).IsDivAssigned,"#C7EEC7","#FFFFFF") %>">
	    <td><input type="checkbox" name="chk" value="<%= i %>" <%= CHKIIF(oBizProfit.FItemList(i).FdivCnt>0,"disabled","") %> ></td>
	    <td><%= oBizProfit.FItemList(i).FSLDATE %></td>
	    <td><%= oBizProfit.FItemList(i).FBIZSECTION_NM %></td>
	    <td><%= oBizProfit.FItemList(i).FACC_GRP_NM %></td>
        <td><%= oBizProfit.FItemList(i).FACC_USE_CD %></td>    
        <td>
            <%= oBizProfit.FItemList(i).FACC_CD_UPNM %>
        </td>
        <td ><%= oBizProfit.FItemList(i).FACC_NM %></td>      
        <td align="left"><%= oBizProfit.FItemList(i).FACC_CD_RMK %></td>
        <td align="left"><%= oBizProfit.FItemList(i).Fcust_NM %>
        <% if Not IsNULL(oBizProfit.FItemList(i).Fcust_cd) then %>
            (<%= oBizProfit.FItemList(i).Fcust_cd %>)
        <% end if %>
        </td>      
        <td align="right" width="70"><%= FormatNumber(oBizProfit.FItemList(i).FdebitSum,0) %></td> 
        <td align="right" width="70"><%= FormatNumber(oBizProfit.FItemList(i).FcreditSum,0) %></td> 
        <td>
            <%= CHKIIF(oBizProfit.FItemList(i).IsINTERNALTRANS,"����","") %>
            
            <%= CHKIIF(oBizProfit.FItemList(i).FSLTR_SAN_STS="0","�̽���","") %>
        </td>
		<td >
		<% if (oBizProfit.FItemList(i).FdivCnt>0) then %>
		<img src="/images/icon_plus.gif" onClick="showDiv('<%= oBizProfit.FItemList(i).FSLTRKEY %>','<%= oBizProfit.FItemList(i).FSLTRKEY_SEQ %>');" style="cursor:pointer">
		<% end if %>
		</td>
		<td >
		<% if (oBizProfit.FItemList(i).FdivCnt>0) then %>
		<img src="/images/i_delete.gif" onClick="delDivAssigned('<%= oBizProfit.FItemList(i).FSLTRKEY %>','<%= oBizProfit.FItemList(i).FSLTRKEY_SEQ %>');" style="cursor:pointer">
		<% end if %>
		</td>
	</tr>
	<% if (oBizProfit.FItemList(i).FdivCnt>0) then %>
	<%
	set odived = new CBizProfit
	odived.FRectSLTRKEY = oBizProfit.FItemList(i).FSLTRKEY
	odived.FRectSLTRKEY_SEQ = oBizProfit.FItemList(i).FSLTRKEY_SEQ
	odived.getBizProfitDivedList    
	%>
	
	<tr style="display:none" id="itr_<%= oBizProfit.FItemList(i).FSLTRKEY %><%= oBizProfit.FItemList(i).FSLTRKEY_SEQ %>">
	    <td colspan="2" bgcolor="#FFFFFF"></td>
	    <td colspan="9" bgcolor="#FFFFFF">
	        <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
            <% for j=0 to odived.FResultCount -1 %>
            <tr bgcolor="#FFFFFF" >
                <td width="120"><%= odived.FItemList(j).FBIZSECTION_NM %></td>
                <td width="70" align="center"><%= odived.FItemList(j).FdivPro %></td>
                <td><%= odived.FItemList(j).getDivTypeName %> <%= odived.FItemList(j).FdivKey %></td>
                <td align="right" width="68"><%= FormatNumber(odived.FItemList(j).FdebitSum,0) %></td>
                <td align="right" width="68"><%= FormatNumber(odived.FItemList(j).FcreditSum,0) %></td>
            </tr>
            <% next %>
            </table>
	    </td>
	    <td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>
	
	<%
	set odived=Nothing
	%>
	<% end if %>
    <%	Next %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="9"></td>
        <td align="right"><%= FormatNumber(debitSum,0) %></td>
        <td align="right"><%= FormatNumber(creditSum,0) %></td>
        <td></td>
	    <td></td>
	    <td></td>
    </tr>
	<% ELSE%>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<td colspan="19">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<%END IF%>
	
	<tr bgcolor="#FFFFFF" >
        <td colspan="24" align="center">
            <% if oBizProfit.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oBizProfit.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + oBizProfit.StartScrollPage to oBizProfit.FScrollCount + oBizProfit.StartScrollPage - 1 %>
				<% if ix>oBizProfit.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>
	
			<% if oBizProfit.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
    </tr>
</table>
</form>
<% end if %>
<form name="frmAct" method="" action="bizProfit_Process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="AssignYYYYMM" value="<%= yyyymm %>">
<input type="hidden" name="divMastKey" value="<%= divMastKey %>">
<input type="hidden" name="SLTRKEY" value="">
<input type="hidden" name="SLTRKEY_SEQ" value="">
</form>
<%
set oBizDivMaster = Nothing
set oBizDivDetail = Nothing
set oBizProfit = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->