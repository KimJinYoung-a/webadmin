<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 전자계산서 관리
' History : 2012.02.07 정윤정 생성
'			2022.09.29 한용민 수정(매칭구분 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim isUseSerp : isUseSerp = true
Dim clsEsero, arrList, intLoop, iStartPage,iEndPage,iX,iPerCnt, iTotCnt,iPageSize, iTotalPage,page, arapCD, arapNM
Dim dSDate,dEDate,ssearchText,itaxsellType,itaxModiType,itaxType, iMapTpYn, iMapTp, iErpSnd, selBiz, mcExpt, exptp
	iPageSize = 150
	page = requestCheckvar(Request("page"),10)
	dSDate = requestCheckvar(Request("dSD"),10)
	dEDate = requestCheckvar(Request("dED"),10)
	ssearchText = requestCheckvar(Request("sST"),200)
	itaxsellType = requestCheckvar(Request("iTST"),10)
	itaxModiType = requestCheckvar(Request("iTMT"),10)
	itaxType = requestCheckvar(Request("iTT"),10)
    iMapTpYn   = requestCheckvar(Request("iMapTpYn"),10)
    iMapTp     = requestCheckvar(Request("iMapTp"),10)
    iErpSnd    = requestCheckvar(Request("iErpSnd"),10)
    selBiz    = requestCheckvar(Request("selBiz"),16)
    mcExpt    = requestCheckvar(Request("mcExpt"),10)
    exptp     = requestCheckvar(Request("exptp"),10)
	arapCD     = requestCheckvar(Request("arapCD"),5)
	arapNM     = requestCheckvar(Request("arapNM"),25)

    if (itaxsellType="") then itaxsellType="0"
	if page="" then page=1
    ''사업자번호.- 치환
    if Len(ssearchText)=12 and InStr(ssearchText,"-")>0 then
        ssearchText = replace(ssearchText,"-","")
    end if

Set clsEsero = new CEsero
    clsEsero.FSDate      =dSDate
	clsEsero.FEDate      =dEDate
	clsEsero.FsearchText =ssearchText
	clsEsero.FtaxsellType=itaxsellType
	clsEsero.FtaxModiType=itaxModiType
	clsEsero.FtaxType    =itaxType
	clsEsero.FMappingTypeYN = iMapTpYn
	clsEsero.FMappingType   = iMapTp
	clsEsero.FErpSendType   = iErpSnd
	clsEsero.FRectArapCD   = arapCD
	clsEsero.FRectBizSecCd = selBiz
	clsEsero.FCurrPage 	= page
	clsEsero.FPageSize 	= iPageSize

	IF (mcExpt="on") then
	    ''수정중.
	    clsEsero.FExpectType = exptp
	    arrList = clsEsero.fnGetEseroTaxMatchExpectList
	    iTotCnt = clsEsero.FTotCnt
	ELSE
	    arrList = clsEsero.fnGetEseroTaxList
	    iTotCnt = clsEsero.FTotCnt
    ENd IF
Set clsEsero = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

''사업부문
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<script type='text/javascript'>

// 페이지 이동
function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}


//새로등록
function jsNewReg(){
	var winD = window.open("popRegFile.asp","popD","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

function jsNewRegNoTax(){
	var winD = window.open("popRegfileNoTax.asp","popRegfileNoTax","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

function jsNewRegXML(){
    var winD = window.open("popRegfileXML.asp","popDXML","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}


function jsNewRegHand(){
    var winD = window.open("popRegfileHand.asp","popDHand","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

function jsAutoIcheMapping(){
    var winD = window.open("popPeriodMapping.asp","popPeriodMapping","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

//미전송 계산서 엑셀 출력
function jssendnottax(){
    var jssendnottax = window.open("popsendnottax.asp?menupos=<%=menupos%>","jssendnottax","width=400, height=300, resizable=yes, scrollbars=yes");
	jssendnottax.focus();
}

//월별 다운로드
function jsDnMonthTax(){
    var stdt=document.frm.dSD.value;
    if (stdt.length<1){
        alert('작성일 시작일을 입력하세요');
        return;
    }
    var iyyyymm=stdt.substring(0,7);
    if (!confirm(iyyyymm + ' 내역을 다운로드 하시겠습니까?')){ return }

    var popwin = window.open("/admin/tax/popMonthTaxList.asp?yyyymm="+iyyyymm,"jssendnottax","width=400, height=300, resizable=yes, scrollbars=yes");
	popwin.focus();
}

function mapByTaxKey(itaxkey){

    if (confirm('매핑처리 하시겠습니까?')){
        var MapByTaxKey = window.open("MapByTaxKey","MapByTaxKey","width=200, height=200, resizable=yes, scrollbars=yes");
	    MapByTaxKey.focus();

        document.frmActLocal.mode.value="MapByTaxKey"
        document.frmActLocal.taxKey.value = itaxkey;
        document.frmActLocal.target="MapByTaxKey";
        document.frmActLocal.submit();
    }
}

function popErpSending(itaxkey){
    var winD = window.open("popRegfileHand.asp?taxkey="+itaxkey,"popErpSending","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}

	//검색
	function jsSearch(){
		document.frm.submit();
	}

	//달력보기
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsMatch(){
	    var frm = document.frmAct;
	    frm.mode.value="autoMapp";
	    frm.submit();
	}

	function popHandMapping(iselltype,iaccDt,itaxkey,isocno){
	    var popURL = 'popHandMapping.asp?iselltype='+iselltype+'&iaccDt='+iaccDt+'&itaxkey='+itaxkey+'&isocno='+isocno;
	    <% if (mcExpt="on") and (exptp="ON") then %>
	    popURL = popURL+"&targetGb=1";
	    <% elseif (mcExpt="on") and (exptp="OF") then %>
	    popURL = popURL+"&targetGb=2";
	    <% end if %>

	    var popwin = window.open(popURL,'popHandMapping','width=1000, height=800, scrollbars=yes, resizable=yes');
		popwin.focus();
	}

function jsFillCal(comp1,comp2,val1,val2){
    comp1.value = val1;
    comp2.value = val2;
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

function chkComp(comp){
    comp.form.exptp.disabled = (!comp.checked);
}

function sendErpArr(frm){

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
	//alert(eseroKey);

	if (confirm('증빙서류를 ERP로 전송하시겠습니까?')){
        document.frmSendErp.mode.value="sendDocErp"
        document.frmSendErp.taxKeyArr.value = eseroKey;
        if (frm.chkPLANDATE.checked==true){
            document.frmSendErp.chkPLANDATE.value = "on";
        }else{
            document.frmSendErp.chkPLANDATE.value = "";
        }
        document.frmSendErp.submit();
    }

}

function sendErpArr_sERP_unlock(frm){
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    e.disabled=false;
		}
	}
}

function sendErpArr_sERP(frm){

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
	//alert(eseroKey);

	if (confirm('증빙서류를 sERP로 전송하시겠습니까?')){
        document.frmSendErp_sERP.mode.value="sendDocErp"
        document.frmSendErp_sERP.taxKeyArr.value = eseroKey;
        //if (frm.chkPLANDATE.checked==true){
        //    document.frmSendErp_sERP.chkPLANDATE.value = "on";
        //}else{
            document.frmSendErp_sERP.chkPLANDATE.value = "";
        //}
        document.frmSendErp_sERP.submit();
    }

}

function jsGetARAP(){
 	var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=600,height=600,resizable=yes, scrollbars=yes");
 	winARAP.focus();
}

function jsReSetARAP(){
 	document.frm.arapCD.value = "";
 	document.frm.arapNM.value = "";
}

function jsSetARAP(dAC, sANM,sACC,sACCNM){
 	document.frm.arapCD.value = dAC;
 	document.frm.arapNM.value = sANM;
}

</script>

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
		<input type="hidden" name="page" value="">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td  rowspan="3" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
			<td align="left">
				<input type="radio" name="iTST" value="0" <%= CHKIIF(itaxsellType="0","checked","") %> >매입
				<input type="radio" name="iTST" value="1" <%= CHKIIF(itaxsellType="1","checked","") %> >매출&nbsp;&nbsp;
				 작성일:
				<input type="text" name="dSD" size="10" value="<%=dSDate%>" onClick="jsPopCal('dSD');"  style="cursor:hand;">
				-
				<input type="text" name="dED" size="10" value="<%=dEDate%>" onClick="jsPopCal('dED');"  style="cursor:hand;">
				<input type="button" value="전전달" class="button" onClick="jsFillCal(document.frm.dSD,document.frm.dED,'<%= Left(DateAdd("m",-2,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",-1,now())),7)+"-01" ),10) %>')";>
				<input type="button" value="전달" class="button" onClick="jsFillCal(document.frm.dSD,document.frm.dED,'<%= Left(DateAdd("m",-1,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",0,now())),7)+"-01" ),10) %>')";>
				<input type="button" value="이번달" class="button" onClick="jsFillCal(document.frm.dSD,document.frm.dED,'<%= Left(DateAdd("m",0,now()),7)+"-01" %>','<%= Left(DateAdd("d",-1,Left(CStr(dateadd("m",1,now())),7)+"-01" ),10) %>')";>

				&nbsp;
				<input type="checkbox" name="mcExpt" <%= CHKIIF(mcExpt="on","checked","") %> onClick="chkComp(this)">매칭예상건검색

				<select name="exptp" <%= CHKIIF(mcExpt="on","","disabled") %>>
				<option value="ON" <%= CHKIIF(exptp="ON","selected","") %> >온라인매입
				<option value="OF" <%= CHKIIF(exptp="OF","selected","") %> >오프라인매입
				<option value="ET" <%= CHKIIF(exptp="ET","selected","") %> >기타매입
				</select>
			</td>
			<td  rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="검색" onClick="jsSearch();">
			</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" >
		    <td align="left">
		        검색어:
				<input type="text" name="sST" value="<%=ssearchText%>" size="30" onKeyPress="if (event.keyCode == 13) document.frm.submit();"><font color="Gray">(사업자등록번호,승인번호,상호,품목명)</font>
				&nbsp;&nbsp;
		        계산서구분:
		        <select Name="iTMT">
		        <option value="">전체
		        <option value="0" <%= CHKIIF(itaxModiType="0","selected","") %> >전자(일반)
		        <option value="1" <%= CHKIIF(itaxModiType="1","selected","") %> >전자(수정)
		        <option value="9" <%= CHKIIF(itaxModiType="9","selected","") %> >기타(수기)
		        </select>
		        &nbsp;&nbsp;
		        과세구분:
		        <select Name="iTT">
		        <option value="">전체
		        <option value="1" <%= CHKIIF(itaxType="1","selected","") %> >과세
		        <option value="2" <%= CHKIIF(itaxType="2","selected","") %> >영세
		        <option value="3" <%= CHKIIF(itaxType="3","selected","") %> >면세
		        </select>
		    </td>
		</tr>
		<tr>
		    <td  bgcolor="#FFFFFF">
		        매칭상태 :
		        <select Name="iMapTpYn">
		        <option value="">전체
		        <option value="Y" <%= CHKIIF(iMapTpYn="Y","selected","") %> >매칭
		        <option value="N" <%= CHKIIF(iMapTpYn="N","selected","") %> >비매칭
		        <option value="E" <%= CHKIIF(iMapTpYn="E","selected","") %> >오류예상건
		        </select>
		        &nbsp;&nbsp;
		        매칭구분 :
				<%= drawSelectBoxMatchTypeName("iMapTp",iMapTp,"") %>
		        <!--<select Name="iMapTp">
		        <option value="">전체
		        <option value="0" <%'= CHKIIF(iMapTp="0","selected","") %> >수기매칭
		        <option value="1" <%'= CHKIIF(iMapTp="1","selected","") %> >온라인 매입
		        <option value="2" <%'= CHKIIF(iMapTp="2","selected","") %> >오프라인 매입
				<option value="4" <%'= CHKIIF(iMapTp="4","selected","") %> >원가
		        <option value="9" <%'= CHKIIF(iMapTp="9","selected","") %> >기타 매입
		        <option value="11" <%'= CHKIIF(iMapTp="11","selected","") %> >매출
		        <option value="900" <%'= CHKIIF(iMapTp="900","selected","") %> >자동이체 매칭
		        <option value="910" <%'= CHKIIF(iMapTp="910","selected","") %> >기타등록 매칭
		        <option value="999" <%'= CHKIIF(iMapTp="999","selected","") %> >수정계산서 매칭
		        </select>-->


		        &nbsp;&nbsp;
		        erp입력구분:
		        <select Name="iErpSnd">
		        <option value="">전체
		        <option value="NN" <%= CHKIIF(iErpSnd="NN","selected","") %> >미입력(온/오프상품매입제외)
		        <option value="N" <%= CHKIIF(iErpSnd="N","selected","") %> >미입력(전체)
		        <option value="Y" <%= CHKIIF(iErpSnd="Y","selected","") %> >입력완료
		        </select>
		        &nbsp;&nbsp;
				사업부문:
                <select name="selBiz">
                <option value="">--선택--</option>
                <% For intLoop = 0 To UBound(arrBizList,2)	%>
            		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(selBiz) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
            	<% Next %>
                </select>
				&nbsp;&nbsp;
				수지항목 :
				<input type="hidden" name="arapCD" value="<%= arapCD %>" >
				<input type="text" name="arapNM" value="<%= arapNM %>" size="13" onClick="jsGetARAP();" readonly>
				<input class="button" type="button" value="X" onClick="jsReSetARAP()">
		    </td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td>
	    <table border="0" cellspacing="0" cellpadding="0" width="100%">
	    <tr>
	        <td align="left">
	            <input type="button" class="button" value="이세로 XL(대량)" onClick="jsNewReg();">
	            <input type="button" class="button" value="면세매입 XL(대량)" onClick="jsNewRegNoTax();">
	            <input type="button" class="button" value="신규등록 XML(1건)" onClick="jsNewRegXML();">
	            <input type="button" class="button" value="신규등록 수기(1건)" onClick="jsNewRegHand();">
	            &nbsp;&nbsp;
	            <input type="button" class="button" value="정기성 매핑설정" onClick="jsAutoIcheMapping();">
	            <input type="button" class="button" value="미전송계산서엑셀출력" onClick="jssendnottax();">
	            <input type="button" class="button" value="월별계산서엑셀출력" onClick="jsDnMonthTax();">
	        </td>
	        <td align="right"><input type="button" class="button" value="자동매핑" onClick="jsMatch();"></td>
	    </tr>
	    </table>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
	    <form name="frmEsero" style="margin:0px;">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="3">
				검색결과 : <b><%=iTotCnt%></b> &nbsp;
				페이지 : <b><%= page %> / <%=iTotalPage%></b>
			</td>
			<td colspan="12" align="right">
			    <% if (isUseSerp) then %>
			        <input type="button" value="선택내역 sERP전송" onClick="sendErpArr_sERP(frmEsero)">
			    <% else %>
                    <input type="checkbox" name="chkPLANDATE" value="" <%= CHKIIF(iMapTp="999","","checked") %> >(수입/지출)예정정보입력
                    <input type="button" value="일괄전송" onClick="sendErpArr(frmEsero)">
                    
                    <% if session("ssBctID")="icommang" or session("ssBctID")="ju1209" then %>
                        <font color=red>sERP[</font> 
                        <input type="button" value="unlock" onClick="sendErpArr_sERP_unlock(frmEsero)">
                        <input type="button" value="sERP 전송" onClick="sendErpArr_sERP(frmEsero)"> <font color=red>]</font>
                    <% end if %>
                <% end if %>    
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    <td rowspan="2" width="20"><input type="checkbox" name="chkALL" value="" onClick="CkeckAll(this,'chk');"></td>
			<td rowspan="2">작성일자</td>
			<td rowspan="2">승인번호</td>
			<td colspan="2"><%IF itaxsellType="0" THEN%>공급자<%ELSE%>공급받는자<%END IF%></td>
			<td rowspan="2">합계금액</td>
			<td rowspan="2">공급가액</td>
			<td rowspan="2">세액</td>
			<td rowspan="2">분류</td>
			<td rowspan="2">종류</td>
			<td rowspan="2">품목명</td>
			<td rowspan="2">매핑<br>타입</td>
			<td rowspan="2">사업부문</td>
			<td rowspan="2">수지항목</td>
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
		<tr align="center" bgcolor="<%=CHKIIF(IsNULL(arrList(39,intLoop)),"#FFFFFF","#CCFFFF")%>" >
		    <td>
		    <% if IsNULL(arrList(33,intLoop)) and (Not IsNULL(arrList(29,intLoop))) and (Not IsNULL(arrList(32,intLoop)))  THEN %> <% ''and (Not IsNULL(arrList(38,intLoop))) %>
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
		    <td><%= getMatchTypeName(arrList(29,intLoop)) %>
		    <% if (mcExpt="on") and (exptp<>"ET") then %>
    		    <% if IsNULL(arrList(29,intLoop)) then %>
                <input type="button" value="매핑" onClick="mapByTaxKey('<%= arrList(0,intLoop) %>')" <%=CHKIIF(arrList(40,intLoop)="CC","disabled","") %>>
                <% if arrList(40,intLoop)="CC" then %>
                <b><font color=red><br>수수료분오발행</font></b>
                <% end if %>
    		    <% end if %>
		    <% end if %>
		    </td>
		    <td><%= getbizSecCDName(arrList(32,intLoop)) %>
		    <% if arrList(35,intLoop)>0 then %>
		    외 <%= arrList(35,intLoop) %>
		    <% end if %>
		    </td>
		    <td><%= arrList(38,intLoop) %></td>
		    <td>
		        <% if Not IsNULL(arrList(33,intLoop)) then %>
			    [<%= arrList(33,intLoop) %>]
			    <%= arrList(34,intLoop) %>
		        <% end if %>
		    </td>
		</tr>
		<%
		Next

		ELSE
		%>
		<tr height=30 align="center" bgcolor="#FFFFFF">
			<td colspan="19">등록된 내용이 없습니다.</td>
		</tr>
		<%END IF%>
		</table>
		</form>
	</td>
</tr>
<!-- 페이지 시작 -->
<%
iPerCnt = 10

iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr height="26" >
	<td colspan="15" align="center">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr valign="bottom" height="25">
	        <td valign="bottom" align="center">
	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(page) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
	        </td>
	    </tr>
		</table>
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
<form name="frmSendErp_sERP" method="post" action="eTax_sERP_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="taxKey" value="">
	<input type="hidden" name="bizSecCd" value="">
	<input type="hidden" name="arap_cd" value="">
	<input type="hidden" name="matchSeq" value="">
	<input type="hidden" name="chkPLANDATE" value="">
	<input type="hidden" name="taxKeyArr" value="">
</form>

<form name="frmSendErp" method="post" action="eTax_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="taxKey" value="">
	<input type="hidden" name="bizSecCd" value="">
	<input type="hidden" name="arap_cd" value="">
	<input type="hidden" name="matchSeq" value="">
	<input type="hidden" name="chkPLANDATE" value="">
	<input type="hidden" name="taxKeyArr" value="">
</form>
<form name="frmAct" method="post" action="eTax_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="stDt" value="">
	<input type="hidden" name="edDt" value="">
</form>

<form name="frmActLocal" method="post" action="eTax_processLocal.asp" style="margin:0px;">
    <input type="hidden" name="mode" value="">
	<input type="hidden" name="taxKey" value="">
</form>
</body>
</html>
