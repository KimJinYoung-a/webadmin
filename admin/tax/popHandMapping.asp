<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 전자계산서 관리 수기 매핑
' History : 2012.02.09 서동석
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
        //여러개 선택 가능
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

//수기 일괄 매핑
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
        alert('선택 내역이 없습니다.');
        return;
    }
    
    if (!confirm('상품매입금 또는 결제요청내역과 매핑할 수 있는 자료는 수기 매핑 않하는것이 원칙입니다.\n\n결제요청 자료와 매핑할 자료가 없는경우에만 사용.\n\n그래도 계속 진행 하시겠습니까?')){
        return;
    }
    
    if (frmBuf.bizSecCd.value==''){
        alert('사업 부문을 선택 하세요.');
        return;
    }
    
    if (frmBuf.arap_cd.value==""){
        alert('수지 항목을 선택 하세요.');
        return;
    }
    
    if (confirm('선택 내역을 수정 계산서 매핑작업으로 처리 하시겠습니까?')){
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

//수정계산서 처리
function jsMinusMapping(){
    //짝수개 이고 합계금액이 같아야 함.
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
       // alert('수정 계산서 처리는 짝수개를 선택 하세요.');  
       // return;
    }
    
    if ((esero_totprice!=0)||(esero_suplyprice!=0)||(esero_vatprice!=0)){
        alert('수정 계산서 처리는 합계 금액이 0으로 처리 되어야 합니다.'+esero_totprice);  
        return;
    }
    
    if (esero_socno=='X'){
        alert('선택 내역의 사업자 번호가 일치하지 않습니다.');  
        return;
    }
    
    if (esero_bizSecCd=='X'){
        alert('선택 내역의 사업 부문이 일치하지 않습니다.');  
        return;
    }
    
    if (esero_bizSecCd==''){
        //alert('사업 부문을 선택 하세요.');
        //jsGetPart(0);
        //return;
    }
    
    if (esero_ChkCNT<1){
        alert('선택 내역이 없습니다.');
        return;
    }
    
    if (frmBuf.bizSecCd.value==""){
        if (!confirm('사업부문 지정 없이 진행 하시겠습니까?')){ 
            return;
        }
    }
    
    if (frmBuf.arap_cd.value==""){
        if (!confirm('수지항목 지정 없이 진행 하시겠습니까?')){ 
            return;
        }
    }
    
    if (confirm('선택 내역을 수정 계산서 매핑작업으로 처리 하시겠습니까?')){
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
        alert('매핑할 이세로 내역을 선택하세요.');
        return;
    }
    
    //**
    //if (esero_ChkCNT!=1){
    //    alert('이세로 내역은 1건만 선택 가능합니다.');
    //    return;
    //}
    
    if (tg_ChkCNT<1){
        alert('매핑할 매입/매출 내역을 선택하세요.');
        return;
    }
    
    if (esero_totprice!=tg_totprice){
        alert('총금액이 일치 하지 않습니다.' + esero_totprice + ':' + tg_totprice);
        return;
    }
    
    if (esero_socno=='X'){
        alert('사업자번호 불일치 -이세로');
        return;
    }
    
    /*
    if (esero_taxkey=='X'){
        alert('국세청 승인번호 불일치 -이세로');
        return;
    }
    */
    
    if (tg_socno=='X'){
        alert('사업자번호 불일치 -매입/매출 내역');
        return;
    }
    
    if (tg_taxkey=='X'){
        alert('국세청 승인번호 불일치-매입/매출 내역');
        return;
    }
    
    if (esero_socno!=tg_socno){
        alert('사업자번호 불일치 이세로:매입/매출 내역 '+ esero_socno + ':' + tg_socno);
        return;
    }
    
    if ((esero_taxkey!='X')&&(esero_taxkey!=tg_taxkey)){
        alert('국세청 승인번호 불일치 이세로:매입/매출 내역'+ esero_taxkey + ':' + tg_taxkey);
        return;
    }
    
   if (confirm('수기 매핑 처리 하시겠습니까?')){
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

//자금관리부서 선택
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//자금관리부서 등록
function jsSetPart(bizSecCd, sPNM){ 
    var frm = document.frmBuf;
    frm.bizSecCd.value = bizSecCd;
    frm.bizSecCd_nm.value = sPNM;
}

//수지항목 불러오기
function jsGetARAP(){
    var rdoGB = "<%= CHKIIF(iselltype="0","2","1") %>";
	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp?rdoGB="+rdoGB,"popARAP1","width=800,height=600,resizable=yes, scrollbars=yes");
	winARAP.focus();
}

//선택 수지항목 가져오기
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
		<td rowspan="2"  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
		<td align="left"> 
			<input type="radio" name="iselltype" value="0" <%= CHKIIF(iselltype="0","checked","") %> >매입 
			<input type="radio" name="iselltype" value="1" <%= CHKIIF(iselltype="1","checked","") %> >매출&nbsp;&nbsp;
			 작성일:
			<input type="text" name="dSDate" size="10" value="<%=dSDate%>" onClick="calendarOpen(frm.dSDate);"  style="cursor:hand;">
			-
			<input type="text" name="dEDate" size="10" value="<%=dEDate%>" onClick="calendarOpen(frm.dEDate);"  style="cursor:hand;">
			&nbsp;&nbsp;사업자등록번호:
			<input type="text" name="isocno" value="<%=isocno%>" size="15">
			
			
		</td> 
		<td rowspan="2"  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr  bgcolor="#FFFFFF" >
	    <td >
	    매핑 검색 구분:
	    <input type="radio"  name="targetGb" value="1" <%= CHKIIF(targetGb="1","checked","") %> >온라인 매입   
	    <input type="radio"  name="targetGb" value="2" <%= CHKIIF(targetGb="2","checked","") %> >오프라인 매입   
	    <input type="radio"  name="targetGb" value="9" <%= CHKIIF(targetGb="9","checked","") %> >기타매입  
	    &nbsp;&nbsp;
	    <input type="radio"  name="targetGb" value="11" <%= CHKIIF(targetGb="11","checked","") %> >매출 
	     
	    </td>
	</tr>
	</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBuf" method="get" action="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		이세로 내역 검색결과 : <b><%=iTotCnt%></b> &nbsp;
	</td>
	<td><input type="button" value="매핑" class="button" onClick="jsMatch();"></td>
	<td colspan="10" align="right">
	    <input type="hidden" name="matchType" value="">
	    
	    &nbsp;
	    거래처코드 : <input type="text" name="cust_cd" value="<%= cust_cd %>" size="8" class="text_ro">
	    
	    사업부문<input type="text" name="bizSecCd_nm" value="" size="10"><input type="hidden" name="bizSecCd" value="" >
	    <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer">
	    &nbsp;
	    수지항목<input type="text" name="arap_cd_nm" value="" size="10"><input type="hidden" name="arap_cd" value="" >
	    <img src="/images/icon_search.jpg" onClick="jsGetARAP();" style="cursor:pointer">
	    <input type="button" value="수정계산서처리" class="button" onClick="jsMinusMapping();">
	    <br>
	    <input type="button" value=" 수기 일괄지정 " class="button" onClick="jsHandMappingArr();">
    </td>
    </form>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="20"></td>
	<td rowspan="2">계산서<br>작성일자</td>
	<td rowspan="2">승인번호</td>
	
	<td colspan="2"><%IF iselltype="0" THEN%>공급자<%ELSE%>공급받는자<%END IF%></td> 
	<td rowspan="2">합계금액</td>  
	<td rowspan="2">공급가액</td> 	
	<td rowspan="2">세액</td> 
	<td rowspan="2">분류</td> 
	<td rowspan="2">종류</td>  
	<td rowspan="2">품목명</td>  
	<td rowspan="2">매핑<br>타입</td> 
	<td rowspan="2">사업부문</td> 
	<td rowspan="2">ERP<br>전송상태</td> 
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 	
	<td>사업자등록번호</td>
	<!-- td>종</td -->
	<td>상호</td>
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
        외 <%= arrList(35,intLoop) %>
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
    <td colspan="5" align="right">선택내역합계</td>
    <td align="right"><input type="text" name="esubtotSum" size="10" class="text" style="text-align=right"></td>
    <td colspan="8"></td>
</tr>
<% ELSE%>
<tr height=30 align="center" bgcolor="#FFFFFF">				
	<td colspan="19">검색 내용이 없습니다.</td>	
</tr>
<%END IF%>
</form>
</table>
<p>

<%
''' 온라인 정산내역.
''dSDate = "2010-01" '임시
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
	    온라인정산 내역 검색결과 : <b><%=sTotCnt%></b> &nbsp;
	<% ELSEIF (targetGb="2") then %>
		오프라인정산 내역 검색결과 : <b><%=sTotCnt%></b> &nbsp;
    <% ELSEIF (targetGb="9") then %>
		기타매입 내역 검색결과 : <b><%=sTotCnt%></b> &nbsp;
	<% ELSEIF (targetGb="11") then %>
		매출 내역 검색결과 : <b><%=sTotCnt%></b> &nbsp;
	<% END IF %>
	</td>
	<td colspan="13" align="right"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="20"></td>
	<td rowspan="2">작성일자</td>
	<td rowspan="2">승인번호</td>
	<td colspan="2"><%IF iselltype="0" THEN%>공급자<%ELSE%>공급받는자<%END IF%></td> 
	<% if (targetGb="9") then %>
	<td rowspan="2">합계금액</td>  
	<td rowspan="2">공급가액</td> 	
	<td rowspan="2">세액</td> 
	<% else %>
	<td rowspan="2">합계금액</td>  
	<% end if %>
	<td rowspan="2">분류</td> 
	<td rowspan="2">종류</td>  
	<td rowspan="2">품목명</td>  
	<td rowspan="2">상태</td> 
	<% if (targetGb="9") then %>
	<td rowspan="2">결제일</td> 
	<% end if %>
	<td rowspan="2">ERP<br>전송(결제)</td> 
	<td rowspan="2">ERP<br>전송(계산서)</td> 
	<td rowspan="2">보기</td> 
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 	
	<td>사업자등록번호</td>
	<!-- td>종</td -->
	<td>상호</td>
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
        <% if (sArr(19,intLoop)=8) then %> <!--계산서차후-->
        <input type="hidden" name="totprice" value="<%= sArr(18,intLoop) %>"> <!--결제요청액-->
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
        <td >결제요청액</td>
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
        매입
        <% else %>
        매출
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
         <%= chkiif(sArr(5,intLoop)="Y","발급","미발급") %>
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
	<td colspan="19">검색 내용이 없습니다.</td>	
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