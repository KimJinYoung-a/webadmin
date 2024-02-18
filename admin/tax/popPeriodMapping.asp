<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 전자계산서 관리 정기성 자료 매핑
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
        alert('매입/매출구분을 선택하세요.');
        frm.TaxSellType.focus();
        return;
    }  
    
    if (frm.matchType.value==""){
        alert('매핑구분을 선택하세요.');
        frm.matchType.focus();
        return;
    }   
    
    if (frm.autoIcheTitle.value==""){
        alert('매핑명칭을 입력하세요.');
        frm.autoIcheTitle.focus();
        return;
    } 
    
    if (frm.corpNo.value==""){
        alert('사업자번호를 입력하세요.');
        frm.corpNo.focus();
        return;
    } 
    
    if (frm.cust_cd.value==""){
        alert('거래처 코드를 선택 하세요.');
        
        return;
    } 
    
    
    if (frm.matchType.value=="900"){
        //자동이체
        //if (frm.mayPrice.value==""){
        //    alert('금액을 입력하세요.');
        //    frm.mayPrice.focus();
        //    return;
        //}
    
        if (frm.mayPumok.value.length<3){
            alert('품목을 입력하세요. 3자이상');
            frm.mayPumok.focus();
            return;
        }
        
        //if (frm.mayIcheDate.value==""){
        //    alert('입출금일을 입력하세요.');
        //    return;
        //}
        
        //if (frm.mayAcctJukyo.value==""){
        //    alert('입출금 적요를 입력하세요.');
        //    frm.mayAcctJukyo.focus();
        //    return;
        //}
    }else{
        if (frm.mayPumok.value.length<3){
            alert('품목을 입력하세요. 3자이상');
            frm.mayPumok.focus();
            return;
        }
    }
    //mayIcheDate
    
    if (frm.bizSecCd.value==""){
        alert('사업부분을 선택 하세요.');
        return;
    } 
    
    if (frm.arap_cd.value==""){
        alert('수지항목을 선택 하세요.');
        return;
    } 


    var regMn='등록';
    if (!isreg)  regMn='수정';
    if (confirm(regMn + ' 하시겠습니까?')){
        frm.mode.value="regPeriod";
        frm.submit();
    }
}

function delPeriodMapping(){
    var frm=document.frmReg;
    if (confirm('삭제 하시겠습니까?')){
        frm.mode.value="delPeriod";
        frm.submit();
    }
}

//자금관리부서 선택
function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//자금관리부서 등록
function jsSetPart(bizSecCd, sPNM){ 
    var frm = document.frmReg;
    frm.bizSecCd.value = bizSecCd;
    frm.AssignBizSecName.value = sPNM;
}

//수지항목 불러오기
function jsGetARAP(){
    var rdoGB = "2"; //지출
	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp?rdoGB="+rdoGB,"popARAP1","width=800,height=600,resizable=yes, scrollbars=yes");
	winARAP.focus();
}

//선택 수지항목 가져오기
function jsSetARAP(dAC, sANM,sACC,sACCNM){ 
    var frm = document.frmReg;
    frm.arap_cd.value = dAC;
    frm.AssignArapNm.value = sANM;
	
}

//거래처 정보 보기
function jsGetCust(){
	var Strparm="";
	var cust_cd = ""; 
	var rdoCgbn = "2"; //매입
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

//거래처 선택
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
	    alert('선택 내역이 없습니다.');
	    return;
	}
	
	if (confirm('선택 내역을 매칭 처리 하시겠습니까?')){
	    
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
	    alert('선택 내역이 없습니다.');
	    return;
	}
	//alert(eseroKey);
	
	if (confirm('증빙서류를 ERP로 전송하시겠습니까?')){
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
		<td rowspan="2"  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
		<td align="left"> 
			<input type="radio" name="sellType" value="0" <%= CHKIIF(sellType="0","checked","") %> >매입 
			<input type="radio" name="sellType" value="1" <%= CHKIIF(sellType="1","checked","") %> >매출&nbsp;&nbsp;
			
			&nbsp;&nbsp;사업자등록번호:
			<input type="text" name="isocno" value="<%=isocno%>" size="15">
			
			
		</td> 
		<td rowspan="2"  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="document.frm.submit();">
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
            <td rowspan="2">매입/매출</td>
            <td rowspan="2">매핑구분</td>
            <td rowspan="2">매핑명칭</td>
            <td colspan="5">
                계산서정보
            </td>
            <td colspan="2">
                입출금정보
            </td>
            <td colspan="2">
                매칭정보
            </td>
        </tr>
        <tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
            <td>사업자번호</td>
            <td>거래처코드</td>
            <td>금액</td>
            <td>날짜</td>
            <td>품목</td>
            <td>이체일</td>
            <td>적요</td>
            <td>사업부분</td>
            <td>수지항목</td>
        </tr>
        <tr bgcolor="FFFFFF" align="center"> 
        
            <td>
                <select Name="TaxSellType">
		        <option value="">선택
		        <option value="0" <%= CHKIIF(TaxSellType="0","selected","") %> >매입
		        <option value="1" <%= CHKIIF(TaxSellType="1","selected","") %> >매출
		        </select>
            </td>
            <td>
                <select Name="matchType">
		        <option value="">선택
		        <option value="900" <%= CHKIIF(matchType="900","selected","") %> >자동이체
		        <option value="910" <%= CHKIIF(matchType="910","selected","") %> >기타등록
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
    * 자동이체인경우 계산서 금액 및 품목은 필수 / 그외의 경우 품목은 필수 (3자이상)
    <br>
    * 날짜는 말일의 경우 (31)로 입력
    </td>
</tr>
<tr bgcolor="FFFFFF">
    <td align="center">
        <% if (CStr(autoIcheIdx)<>"") then %>
        <input type="button" value="수정" onclick="regPeriodMapping(false)">
        &nbsp;&nbsp;
        <input type="button" value="삭제" onclick="delPeriodMapping()">
        <% else %>
        <input type="button" value="신규등록" onclick="regPeriodMapping(true)">
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
		검색결과 : <b><%=clsPMapping.FTotCnt%></b> &nbsp;
	</td>
	<td colspan="12" align="right">
	    <% IF (mode<>"reg") THEN %>
	    <input type="button" value="신규등록" onClick="research('','reg');">
	    <% end if %>
    </td>
    </form>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
    <td rowspan="2">검색</td>
    <td rowspan="2">매입/매출</td>
    <td rowspan="2">매핑구분</td>
    <td rowspan="2">매핑명칭</td>
    <td colspan="5">
        계산서정보
    </td>
    <td colspan="2">
        입출금정보
    </td>
    <td colspan="2">
        매칭정보
    </td>
    <td rowspan="2">
        수정
    </td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
    <td>사업자번호<br>거래처코드</td>
    <td>거래처명</td>
    <td>금액</td>
    <td>날짜</td>
    <td>품목</td>
    <td>이체일</td>
    <td>적요</td>
    <td>사업부분</td>
    <td>수지항목</td>
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
    <td><input type="button" value="수정" onClick="research('<%= clsPMapping.FItemList(i).FautoIcheIdx %>','edit');"></td>
</tr>	
<%	Next
ELSE%>
<tr height=30 align="center" bgcolor="#FFFFFF">				
	<td colspan="19">검색 내용이 없습니다.</td>	
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
		검색결과 : <b><%=TotCnt%></b> &nbsp;
	</td>
	<td align="right" colspan="2"> 
	<% if isPLAN_DATEDefaultSend(matchType, taxSellType, arap_cd) then %>
    <input type="checkbox" name="chkPLANDATE" value="" checked >(수입/지출)예정정보입력
    <% else %>
    <input type="checkbox" name="chkPLANDATE" value=""  >(수입/지출)예정정보입력
    <% end if %>
                
    <input type="button" value="일괄전송" onClick="sendErpArr(frmEsero)">
    </td>
	<td colspan="12" align="right">
	    <input type="hidden" name="matchType" value="<%= matchType %>">
	    매칭타입 : <%= getMatchTypeName(matchType) %>
	    &nbsp;
	    거래처코드 : <input type="text" name="cust_cd" value="<%= cust_cd %>" size="8" class="text_ro">
	    &nbsp;
	    사업부문<input type="text" name="bizSecCd_nm" value="<%= AssignBizSecName %>" size="16" class="text_ro"><input type="hidden" name="bizSecCd" value="<%= bizSecCd %>" >
	    <img src="/images/icon_search.jpg" onClick="jsGetPart();" style="cursor:pointer">
	    &nbsp;
	    수지항목<input type="text" name="arap_cd_nm" value="<%= AssignArapNm %>" size="20" class="text_ro"><input type="hidden" name="arap_cd" value="<%= arap_cd %>" >
	    <img src="/images/icon_search.jpg" onClick="jsGetARAP();" style="cursor:pointer">
	    <input type="button" value="일괄지정" onClick="mapPeriod(frmEsero);">
	    
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
    <td rowspan="2" width="40"><input type="checkbox" name="chkALL" value="" onClick="CkeckAll(this,'chk');"><br>(지정)</td>
    <td rowspan="2" width="40"><input type="checkbox" name="chkALL2" value="" onClick="CkeckAll(this,'chk2');"><br>(전송)</td>
	<td rowspan="2">계산서<br>작성일자</td>
	<td rowspan="2">승인번호</td>
	
	<td colspan="2"><%IF TaxSellType="0" THEN%>공급자<%ELSE%>공급받는자<%END IF%></td> 
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