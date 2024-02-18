<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ���ڰ�꼭 ���� ���� ����
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
Dim iselltype : iselltype = requestCheckvar(request("iselltype"),10)
Dim iaccDt   : iaccDt   = requestCheckvar(request("iaccDt"),10)
Dim itaxkey  : itaxkey  = requestCheckvar(request("itaxkey"),10)
Dim isocno   : isocno   = requestCheckvar(request("isocno"),32)
Dim targetGb : targetGb = Trim(requestCheckvar(request("targetGb"),10))

Dim dSDate : dSDate   = requestCheckvar(request("dSDate"),10)
Dim dEDate : dEDate   = requestCheckvar(request("dEDate"),10)

Dim groupid

if (dSDate="") then
    dSDate = Left(DateAdd("m",-1,iaccDt),7)+"-01"
    dEDate = Left(DateAdd("d",-1,DateAdd("m",3,dSDate)),10)
end if

if (iselltype="1") and (targetGb="") then targetGb="11"
if (targetGb="") then targetGb=9

Dim clsEsero, arrList, iTotCnt, intLoop

Set clsEsero = new CEsero 
    clsEsero.FSDate      = dSDate
	clsEsero.FEDate      = dEDate
	clsEsero.FRectCorpNo = isocno    
	clsEsero.FtaxsellType= iselltype  
	''clsEsero.FtaxModiType= itaxModiType  
	''clsEsero.FtaxType    = itaxType      
	''clsEsero.FMappingTypeYN = iMapTpYn
	''clsEsero.FMappingType   = iMapTp
	clsEsero.FCurrPage 	= 1
	clsEsero.FPageSize 	= 100
	arrList = clsEsero.fnGetEseroTaxList 	
	iTotCnt = clsEsero.FTotCnt 

Set clsEsero = nothing	


Dim sqlStr
Dim sArr
Dim sTotCnt : sTotCnt=0

Dim cust_cd 
Dim retVal
if (isocno<>"") then
    retVal = fnGetCustCDByCorpNo(isocno,cust_cd)
end if
%>

<script language='javascript'>
function popTargetDetail(itargetGb,iidx,iridx){
    var popURL ='';
    if (itargetGb=="1"){
        popURL = "/admin/upchejungsan/nowjungsanmasteredit.asp?id="+iidx;
    }else if (itargetGb=="2"){
        popURL = "/admin/offupchejungsan/off_jungsanstateedit.asp?idx="+iidx;
    }else if (itargetGb=="9"){
        popURL = "/admin/approval/eapp/modeappPayDoc.asp?ipridx="+iidx+"&iridx="+iridx;
    }else if (itargetGb=="11"){
        popURL = "/cscenter/taxsheet/Tax_view.asp?taxIdx="+iidx;
    }
    
    var popWin = window.open(popURL,'popTargetDetail','width=900,height=600,scrollbars=yes,resizable=yes');
    popWin.focus();
}

function onlyOneCheck(frm,comp){
    var compArr = eval(frm.name+'.'+comp.name);

    if (!compArr.length){ return }

    for (var i=0;i<compArr.length;i++){
        if(compArr[i].value!=comp.value){
            compArr[i].checked=false;
        }
    }
}

function checkSel(comp){
    if (comp.form.name=="frmEsero"){
        reCalcuETaxSum();
        //onlyOneCheck(comp.form,comp);
    }else if(comp.form.name=="frmTarget"){
        //������ ���� ����
    }
}

function reCalcuETaxSum(){
    var frm = document.frmEsero;
    var isumval=0;

    if (!frm.chk.length){
        if(frm.chk.checked){
            isumval   = frmEsero.totprice.value*1;
        }
    }else{
        for (var i=0;i<frm.chk.length;i++){
            if ((frm.chk[i].checked)&&(frm.chk[i].name=="chk")){
                isumval += frm.totprice[i].value*1;
            }
        }
    }
    
    frmEsero.esubtotSum.value=isumval;
}

//���� �ϰ� ����
function jsHandMappingArr(){
    var esero_ChkCNT = 0;
    var esero_taxkey="";
    if (!frmEsero.chk.length){
        if(frmEsero.chk.checked){
            esero_ChkCNT++;
            esero_taxkey   = frmEsero.taxkey.value;
        }
    }else{
        for (var i=0;i<frmEsero.chk.length;i++){
            if(frmEsero.chk[i].checked){
                esero_ChkCNT++;
                esero_taxkey     = esero_taxkey + frmEsero.taxkey[i].value + ",";
            }
        }
    }
    
    if (esero_ChkCNT<1){
        alert('���� ������ �����ϴ�.');
        return;
    }
    
    if (!confirm('��ǰ���Ա� �Ǵ� ������û������ ������ �� �ִ� �ڷ�� ���� ���� ���ϴ°��� ��Ģ�Դϴ�.\n\n������û �ڷ�� ������ �ڷᰡ ���°�쿡�� ���.\n\n�׷��� ��� ���� �Ͻðڽ��ϱ�?')){
        return;
    }
    
    if (frmBuf.bizSecCd.value==''){
        alert('��� �ι��� ���� �ϼ���.');
        return;
    }
    
    if (frmBuf.arap_cd.value==""){
        alert('���� �׸��� ���� �ϼ���.');
        return;
    }
    
    if (confirm('���� ������ ���� ��꼭 �����۾����� ó�� �Ͻðڽ��ϱ�?')){
        frmMap.action ="eTax_process.asp";
        frmMap.mode.value="modiTaxMapping";
        frmMap.eseroKey.value = esero_taxkey;
        frmMap.cust_cd.value = frmBuf.cust_cd.value;
        frmMap.bizSecCd.value=frmBuf.bizSecCd.value;
        frmMap.arap_cd.value=frmBuf.arap_cd.value;
        frmMap.matchType.value="0";
        frmMap.submit();
    }
}

//������꼭 ó��
function jsMinusMapping(){
    //¦���� �̰� �հ�ݾ��� ���ƾ� ��.
    var esero_ChkCNT = 0;
    var esero_socno ='';
    var esero_taxkey='';
    var esero_totprice =0;
    var esero_suplyprice =0;
    var esero_vatprice =0;
    var esero_bizSecCd='';
    
    if (!frmEsero.chk.length){
        if(frmEsero.chk.checked){
            esero_ChkCNT++;
            esero_socno = frmEsero.socno.value;
            esero_bizSecCd = frmEsero.bizSecCd.value;
            esero_taxkey   = frmEsero.taxkey.value;
            esero_totprice = frmEsero.totprice.value;
            esero_suplyprice = frmEsero.suplyprice.value;
            esero_vatprice = frmEsero.vatprice.value;
        }
    }else{
        for (var i=0;i<frmEsero.chk.length;i++){
            if(frmEsero.chk[i].checked){
                esero_ChkCNT++;
                if (esero_socno==''){
                    esero_socno = frmEsero.socno[i].value;
                }else if (esero_socno!=frmEsero.socno[i].value){
                    esero_socno='X';
                }
                if (esero_bizSecCd==''){
                    esero_bizSecCd = frmEsero.bizSecCd[i].value;
                }else if (esero_bizSecCd!=frmEsero.bizSecCd[i].value){
                    esero_bizSecCd='X';
                }
                esero_taxkey     = esero_taxkey + frmEsero.taxkey[i].value + ",";
                esero_totprice   += frmEsero.totprice[i].value*1;
                esero_suplyprice += frmEsero.suplyprice[i].value*1;
                esero_vatprice   += frmEsero.vatprice[i].value*1;
            }
        }
    }
    
    
    if (esero_ChkCNT%2!=0){
       // alert('���� ��꼭 ó���� ¦������ ���� �ϼ���.');  
       // return;
    }
    
    if ((esero_totprice!=0)||(esero_suplyprice!=0)||(esero_vatprice!=0)){
        alert('���� ��꼭 ó���� �հ� �ݾ��� 0���� ó�� �Ǿ�� �մϴ�.'+esero_totprice);  
        return;
    }
    
    if (esero_socno=='X'){
        alert('���� ������ ����� ��ȣ�� ��ġ���� �ʽ��ϴ�.');  
        return;
    }
    
    if (esero_bizSecCd=='X'){
        alert('���� ������ ��� �ι��� ��ġ���� �ʽ��ϴ�.');  
        return;
    }
    
    if (esero_bizSecCd==''){
        //alert('��� �ι��� ���� �ϼ���.');
        //jsGetPart(0);
        //return;
    }
    
    if (esero_ChkCNT<1){
        alert('���� ������ �����ϴ�.');
        return;
    }
    
    if (frmBuf.bizSecCd.value==""){
        if (!confirm('����ι� ���� ���� ���� �Ͻðڽ��ϱ�?')){ 
            return;
        }
    }
    
    if (frmBuf.arap_cd.value==""){
        if (!confirm('�����׸� ���� ���� ���� �Ͻðڽ��ϱ�?')){ 
            return;
        }
    }
    
    if (confirm('���� ������ ���� ��꼭 �����۾����� ó�� �Ͻðڽ��ϱ�?')){
        frmMap.action ="eTax_process.asp";
        frmMap.mode.value="modiTaxMapping";
        frmMap.eseroKey.value = esero_taxkey;
        frmMap.cust_cd.value = frmBuf.cust_cd.value;
        frmMap.bizSecCd.value=frmBuf.bizSecCd.value;
        frmMap.arap_cd.value=frmBuf.arap_cd.value;
        frmMap.matchType.value="999";
        frmMap.submit();
    }
}

function jsMatch(){
    var esero_ChkCNT = 0;
    var esero_socno ='';
    var esero_taxkey='';
    var esero_totprice =0;
    var esero_suplyprice =0;
    var esero_vatprice =0;
    var taxkeyArr ='';
    
    var tg_ChkCNT = 0;
    var tg_socno ='';
    var tg_taxkey='';
    var tg_totprice =0;
    var tg_suplyprice =0;
    var tg_vatprice =0;
    var tg_Arr ='';
    
    if (!frmEsero.chk.length){
        if(frmEsero.chk.checked){
            esero_ChkCNT++;
            esero_socno = frmEsero.socno.value;
            esero_taxkey = frmEsero.taxkey.value; 
            esero_totprice = frmEsero.totprice.value;
            esero_suplyprice = frmEsero.suplyprice.value;
            esero_vatprice = frmEsero.vatprice.value;
        }
    }else{
        for (var i=0;i<frmEsero.chk.length;i++){
            if(frmEsero.chk[i].checked){
                esero_ChkCNT++;
                if (esero_socno==''){
                    esero_socno = frmEsero.socno[i].value;
                }else if (esero_socno!=frmEsero.socno[i].value){
                    esero_socno='X';
                }
                
                if (esero_taxkey==''){
                    esero_taxkey = frmEsero.taxkey[i].value; 
                }else if (esero_taxkey!=frmEsero.taxkey[i].value){
                    esero_taxkey='X';
                }
                esero_totprice += frmEsero.totprice[i].value*1;
                esero_suplyprice += frmEsero.suplyprice[i].value*1;
                esero_vatprice += frmEsero.vatprice[i].value*1;
                taxkeyArr += frmEsero.taxkey[i].value+',';
            }
        }
    }
    
    if (!frmTarget.chk.length){
        if(frmTarget.chk.checked){
            tg_ChkCNT++;
            tg_socno = frmTarget.socno.value;
            tg_taxkey = frmTarget.taxkey.value; 
            tg_totprice = frmTarget.totprice.value;
            tg_suplyprice = frmTarget.suplyprice.value;
            tg_vatprice = frmTarget.vatprice.value;
            tg_Arr = frmTarget.chk.value;
        }
    }else{
        for (var i=0;i<frmTarget.chk.length;i++){
            if(frmTarget.chk[i].checked){
                tg_ChkCNT++;
                if (tg_socno==''){
                    tg_socno = frmTarget.socno[i].value;
                }else if (tg_socno!=frmTarget.socno[i].value){
                    tg_socno='X';
                }
                
                if (tg_taxkey==''){
                    tg_taxkey = frmTarget.taxkey[i].value; 
                }else if (tg_taxkey!=frmTarget.taxkey[i].value){
                    tg_taxkey='X';
                }
                tg_totprice += frmTarget.totprice[i].value*1;
                tg_suplyprice += frmTarget.suplyprice[i].value*1;
                tg_vatprice += frmTarget.vatprice[i].value*1;
                tg_Arr += frmTarget.chk[i].value+',';
            }
        }
    }
    
    if (esero_ChkCNT<1){
        alert('������ �̼��� ������ �����ϼ���.');
        return;
    }
    
    //**
    //if (esero_ChkCNT!=1){
    //    alert('�̼��� ������ 1�Ǹ� ���� �����մϴ�.');
    //    return;
    //}
    
    if (tg_ChkCNT<1){
        alert('������ ����/���� ������ �����ϼ���.');
        return;
    }
    
    if (esero_totprice!=tg_totprice){
        alert('�ѱݾ��� ��ġ ���� �ʽ��ϴ�.' + esero_totprice + ':' + tg_totprice);
        return;
    }
    
    if (esero_socno=='X'){
        alert('����ڹ�ȣ ����ġ -�̼���');
        return;
    }
    
    /*
    if (esero_taxkey=='X'){
        alert('����û ���ι�ȣ ����ġ -�̼���');
        return;
    }
    */
    
    if (tg_socno=='X'){
        alert('����ڹ�ȣ ����ġ -����/���� ����');
        return;
    }
    
    if (tg_taxkey=='X'){
        alert('����û ���ι�ȣ ����ġ-����/���� ����');
        return;
    }
    
    if (esero_socno!=tg_socno){
        alert('����ڹ�ȣ ����ġ �̼���:����/���� ���� '+ esero_socno + ':' + tg_socno);
        return;
    }
    
    if ((esero_taxkey!='X')&&(esero_taxkey!=tg_taxkey)){
        alert('����û ���ι�ȣ ����ġ �̼���:����/���� ����'+ esero_taxkey + ':' + tg_taxkey);
        return;
    }
    
   if (confirm('���� ���� ó�� �Ͻðڽ��ϱ�?')){
        frmMap.action ="eTax_process.asp";
        frmMap.mode.value="handTaxMapping";
        if (esero_ChkCNT==1){
            frmMap.eseroKey.value = esero_taxkey; 
        }else{
            frmMap.taxkeyArr.value = taxkeyArr; 
        }
        
        frmMap.targetArr.value = tg_Arr;
        frmMap.targetCnt.value = tg_ChkCNT;
        
        //frmMap.bizSecCd.value=frmBuf.bizSecCd.value;
        //frmMap.arap_cd.value=frmBuf.arap_cd.value;
        frmMap.submit();
   }
}

function popErpSending(itaxkey){
    var winD = window.open("popRegfileHand.asp?taxkey="+itaxkey,"popErpSending","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

//�ڱݰ����μ� ����
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//�ڱݰ����μ� ���
function jsSetPart(bizSecCd, sPNM){ 
    var frm = document.frmBuf;
    frm.bizSecCd.value = bizSecCd;
    frm.bizSecCd_nm.value = sPNM;
}

//�����׸� �ҷ�����
function jsGetARAP(){
    var rdoGB = "<%= CHKIIF(iselltype="0","2","1") %>";
	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp?rdoGB="+rdoGB,"popARAP1","width=800,height=600,resizable=yes, scrollbars=yes");
	winARAP.focus();
}

//���� �����׸� ��������
function jsSetARAP(dAC, sANM,sACC,sACCNM){ 
    var frm = document.frmBuf;
    frm.arap_cd.value = dAC;
    frm.arap_cd_nm.value = sANM;
}


</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value=""> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2"  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
		<td align="left"> 
			<input type="radio" name="iselltype" value="0" <%= CHKIIF(iselltype="0","checked","") %> >���� 
			<input type="radio" name="iselltype" value="1" <%= CHKIIF(iselltype="1","checked","") %> >����&nbsp;&nbsp;
			 �ۼ���:
			<input type="text" name="dSDate" size="10" value="<%=dSDate%>" onClick="calendarOpen(frm.dSDate);"  style="cursor:hand;">
			-
			<input type="text" name="dEDate" size="10" value="<%=dEDate%>" onClick="calendarOpen(frm.dEDate);"  style="cursor:hand;">
			&nbsp;&nbsp;����ڵ�Ϲ�ȣ:
			<input type="text" name="isocno" value="<%=isocno%>" size="15">
			
			
		</td> 
		<td rowspan="2"  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr  bgcolor="#FFFFFF" >
	    <td >
	    ���� �˻� ����:
	    <input type="radio"  name="targetGb" value="1" <%= CHKIIF(targetGb="1","checked","") %> >�¶��� ����   
	    <input type="radio"  name="targetGb" value="2" <%= CHKIIF(targetGb="2","checked","") %> >�������� ����   
	    <input type="radio"  name="targetGb" value="9" <%= CHKIIF(targetGb="9","checked","") %> >��Ÿ����  
	    &nbsp;&nbsp;
	    <input type="radio"  name="targetGb" value="11" <%= CHKIIF(targetGb="11","checked","") %> >���� 
	     
	    </td>
	</tr>
	</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBuf" method="get" action="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		�̼��� ���� �˻���� : <b><%=iTotCnt%></b> &nbsp;
	</td>
	<td><input type="button" value="����" class="button" onClick="jsMatch();"></td>
	<td colspan="10" align="right">
	    <input type="hidden" name="matchType" value="">
	    
	    &nbsp;
	    �ŷ�ó�ڵ� : <input type="text" name="cust_cd" value="<%= cust_cd %>" size="8" class="text_ro">
	    
	    ����ι�<input type="text" name="bizSecCd_nm" value="" size="10"><input type="hidden" name="bizSecCd" value="" >
	    <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer">
	    &nbsp;
	    �����׸�<input type="text" name="arap_cd_nm" value="" size="10"><input type="hidden" name="arap_cd" value="" >
	    <img src="/images/icon_search.jpg" onClick="jsGetARAP();" style="cursor:pointer">
	    <input type="button" value="������꼭ó��" class="button" onClick="jsMinusMapping();">
	    <br>
	    <input type="button" value=" ���� �ϰ����� " class="button" onClick="jsHandMappingArr();">
    </td>
    </form>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="20"></td>
	<td rowspan="2">��꼭<br>�ۼ�����</td>
	<td rowspan="2">���ι�ȣ</td>
	
	<td colspan="2"><%IF iselltype="0" THEN%>������<%ELSE%>���޹޴���<%END IF%></td> 
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
<form name="frmEsero" method="post">
<%  
IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%> 
<tr align="center" bgcolor="#FFFFFF">
    <input type="hidden" name="socno" value="<%= CHKIIF(arrList(15,intLoop)=1,arrList(7,intLoop),arrList(2,intLoop)) %>">
    <input type="hidden" name="taxkey" value="<%= arrList(0,intLoop) %>">
    <input type="hidden" name="totprice" value="<%= arrList(12,intLoop) %>">
    <input type="hidden" name="suplyprice" value="<%= arrList(13,intLoop) %>">
    <input type="hidden" name="vatprice" value="<%= arrList(14,intLoop) %>">
    <input type="hidden" name="bizSecCd" value="<%= arrList(32,intLoop) %>">
    
    <td>
        <% if IsNULL(arrList(29,intLoop)) THEN %>
        <input type="checkbox" name="chk" value="<%= arrList(0,intLoop) %>" onClick="checkSel(this);">
        <% else %>
        <input type="checkbox" name="chk" value="<%= arrList(0,intLoop) %>" disabled >
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
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="5" align="right">���ó����հ�</td>
    <td align="right"><input type="text" name="esubtotSum" size="10" class="text" style="text-align=right"></td>
    <td colspan="8"></td>
</tr>
<% ELSE%>
<tr height=30 align="center" bgcolor="#FFFFFF">				
	<td colspan="19">�˻� ������ �����ϴ�.</td>	
</tr>
<%END IF%>
</form>
</table>
<p>

<%
''' �¶��� ���곻��.
''dSDate = "2010-01" '�ӽ�
Dim pDate : pDate = DateAdd("m",-3,dSDate)

sArr = fnGetmappingTargetInfo(targetGb,pDate,isocno,"")

If IsArray(sArr) then
    sTotCnt = UBound(sArr,2) +1
end if
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
	<% IF (targetGb="1") then %>
	    �¶������� ���� �˻���� : <b><%=sTotCnt%></b> &nbsp;
	<% ELSEIF (targetGb="2") then %>
		������������ ���� �˻���� : <b><%=sTotCnt%></b> &nbsp;
    <% ELSEIF (targetGb="9") then %>
		��Ÿ���� ���� �˻���� : <b><%=sTotCnt%></b> &nbsp;
	<% ELSEIF (targetGb="11") then %>
		���� ���� �˻���� : <b><%=sTotCnt%></b> &nbsp;
	<% END IF %>
	</td>
	<td colspan="13" align="right"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="20"></td>
	<td rowspan="2">�ۼ�����</td>
	<td rowspan="2">���ι�ȣ</td>
	<td colspan="2"><%IF iselltype="0" THEN%>������<%ELSE%>���޹޴���<%END IF%></td> 
	<% if (targetGb="9") then %>
	<td rowspan="2">�հ�ݾ�</td>  
	<td rowspan="2">���ް���</td> 	
	<td rowspan="2">����</td> 
	<% else %>
	<td rowspan="2">�հ�ݾ�</td>  
	<% end if %>
	<td rowspan="2">�з�</td> 
	<td rowspan="2">����</td>  
	<td rowspan="2">ǰ���</td>  
	<td rowspan="2">����</td> 
	<% if (targetGb="9") then %>
	<td rowspan="2">������</td> 
	<% end if %>
	<td rowspan="2">ERP<br>����(����)</td> 
	<td rowspan="2">ERP<br>����(��꼭)</td> 
	<td rowspan="2">����</td> 
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 	
	<td>����ڵ�Ϲ�ȣ</td>
	<!-- td>��</td -->
	<td>��ȣ</td>
</tr>
<form name="frmTarget">
<%  
IF isArray(sArr) THEN
	For intLoop = 0 To UBound(sArr,2)
	%> 
<tr align="center" bgcolor="#FFFFFF">
    <input type="hidden" name="socno" value="<%= sArr(9,intLoop) %>">
    <input type="hidden" name="taxkey" value="<%= sArr(7,intLoop) %>">
    <% if (targetGb="9") then %>
        <% if (sArr(19,intLoop)=8) then %> <!--��꼭����-->
        <input type="hidden" name="totprice" value="<%= sArr(18,intLoop) %>"> <!--������û��-->
        <% else %>
        <input type="hidden" name="totprice" value="<%= sArr(4,intLoop) %>">
        <% end if %>
    <% else %>
    <input type="hidden" name="totprice" value="<%= sArr(4,intLoop) %>">
    <% end if %>
    <% if (targetGb="9") then %>
    <input type="hidden" name="suplyprice" value="<%= sArr(12,intLoop) %>">
    <input type="hidden" name="vatprice" value="<%= sArr(13,intLoop) %>">
    <% else %>
    <input type="hidden" name="suplyprice" value="0">
    <input type="hidden" name="vatprice" value="0">
    <% end if %>
    <td><input type="checkbox" name="chk" value="<%= sArr(0,intLoop) %>" onClick="checkSel(this);"></td>
    <td><%= sArr(6,intLoop) %></td>
    <td><%= sArr(7,intLoop) %></td>
    <td><%= sArr(9,intLoop) %></td>
    <td><%= sArr(10,intLoop) %></td>
    
    <% if (targetGb="9") then %>
        <% if (sArr(19,intLoop)=8) then %>
        <td >������û��</td>
        <td align="right" colspan="2"><%= FormatNumber(sArr(18,intLoop),0) %></td>
        <% else %>
        <td align="right"><%= FormatNumber(sArr(4,intLoop),0) %></td>
        <td align="right"><%= FormatNumber(sArr(12,intLoop),0) %></td>
        <td align="right"><%= FormatNumber(sArr(13,intLoop),0) %></td>
        <% end if %>
    <% else %>
    <td align="right"><%= FormatNumber(sArr(4,intLoop),0) %></td>
    <% end if %>
    <td>
        <% if (CLNG(targetGb)<10) then %>
        ����
        <% else %>
        ����
        <% end if %>
    </td>
    <td>
        <% if (targetGb="9") then %>
        <%= GetEAppTaxtypeName(sArr(11,intLoop)) %>
        <% elseif (targetGb="11") then %>
        <%= gettaxTypeName(sArr(11,intLoop)) %>
        <% else %>
        <%= GetJungsanTaxtypeName(sArr(11,intLoop)) %>
        <% end if %>
    </td>
    <td>
        <% if (targetGb="9") or (targetGb="11") then %>
        <%= sArr(14,intLoop) %>
        <% else %>
        <%= sArr(1,intLoop) %>&nbsp;<%= sArr(2,intLoop) %>
        <% end if %>
    </td>
    <td>
        <% if (targetGb="9") then %>
        <%= fnGetPayRequestState(sArr(5,intLoop)) %>
        <% elseif (targetGb="11") then %>
         <%= chkiif(sArr(5,intLoop)="Y","�߱�","�̹߱�") %>
        <% else %>
        <font color="<%= GetJungsanStateColor(sArr(5,intLoop)) %>"><%= GetJungsanStateName(sArr(5,intLoop)) %></font>
        <% end if %>
    </td>
    <% if (targetGb="9") then %>
    <td><%=sArr(22,intLoop)%></td>
    <% end if %>
    <td>
        <% if (targetGb="9") then %>
            <% if Not IsNULL(sArr(16,intLoop)) then %>
		    [<%=sArr(15,intLoop)%>]<%=sArr(16,intLoop)%>
		    <% end if %>
        <% else %> 
        
        <% end if %>
    </td>
    <td>
        <% if (targetGb="9") then %>
            <% if Not IsNULL(sArr(20,intLoop)) then %>
		    [<%=sArr(20,intLoop)%>]<%=sArr(21,intLoop)%>
		    <% end if %>
        <% else %> 
        
        <% end if %>
    </td>
    <td>
        <img src="/images/icon_arrow_link.gif" onClick="popTargetDetail('<%= targetGb %>','<%=sArr(0,intLoop)%>'<%IF targetGb="9" then%>,'<%=sArr(17,intLoop)%>'<%END IF%>)" style="cursor:pointer">
    </td>
</tr>	
<%	Next
ELSE%>

<tr height=30 align="center" bgcolor="#FFFFFF">				
	<td colspan="19">�˻� ������ �����ϴ�.</td>	
</tr>
<%END IF%>
</form>
</table>

<form name="frmMap" method="post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="eseroKey" value="">
<input type="hidden" name="targetKey" value="">
<input type="hidden" name="targetArr" value="">
<input type="hidden" name="targetCnt" value="">
<input type="hidden" name="targetGb" value="<%= targetGb %>">
<input type="hidden" name="arap_cd" value="">
<input type="hidden" name="bizSecCd" value="">
<input type="hidden" name="cust_cd" value="">
<input type="hidden" name="matchType" value="">
<input type="hidden" name="taxkeyArr" value="">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<!-- #include virtual="/admin/lib/poptail.asp"-->