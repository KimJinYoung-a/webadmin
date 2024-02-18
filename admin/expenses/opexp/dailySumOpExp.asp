<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비관리 상세   리스트
' History : 2011.06.03 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"--> 
<%
Dim clsPart,clsOpExp,arrPart, arrList, arrType, intLoop 
Dim clsAccount, arrAccount  
Dim dYear, dMonth, iPartTypeIdx, iOpExpPartIdx, iarap_cd
Dim intY, intM
Dim isearchType
Dim iOpExpIdx,dyyyymm, mLastMonthExp,mInExp,mOutExp,mTotExp,sOpExpPartName
Dim blnAdmin, blnWorker, blnReg 
 	dYear			= requestCheckvar(Request("selY"),4)
 	dMonth			= requestCheckvar(Request("selM"),2)
 	isearchType		= requestCheckvar(Request("rdoST"),1) 
 	IF isearchType = "" THEN isearchType =1
 	IF isearchType = 1 THEN
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10) 
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10) 
 	ELSE
 	iarap_cd		= requestCheckvar(Request("selA"),10)
	END IF
 
	iOpExpIdx		= requestCheckvar(Request("hidOE"),10)
 	IF dYear = "" THEN dYear = year(date())
 	IF dMonth = "" THEN dMonth = month(date())	
 		dyyyymm =  dYear&"-"&Format00(2,dMonth) 
 	IF 	iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF 	iOpExpPartIdx = "" THEN iOpExpPartIdx = 0
 	IF 	iarap_cd = "" THEN iarap_cd = 0
		
	'권한초기값 설정-------------- 
 	blnWorker = 0 '담당자
 	blnReg = 0 	'등록권한
  	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '어드민권한	
  	
  	IF blnAdmin and iOpExpPartIdx > 0 THEN blnReg = 1 '어드민권한 있을 경우 등록처리 항상 가능
	 				


 '수지항목 리스트 
set clsAccount = new COpExpAccount
	arrAccount = clsAccount.fnGetArapAllList
set clsAccount = nothing  
	
'운영비 리스트	
Set clsOpExp = new OpExp
	
	IF isearchtype =1 then  
	clsOpExp.FYYYYMM 		=dyyyymm
	clsOpExp.FOpExpPartIdx 	= iOpExpPartIdx   
	clsOpExp.FOpExpIdx 	= iOpExpIdx   
	clsOpExp.fnGetOpExpMonthlyData
	iOpExpidx 	   =  clsOpExp.FOpExpidx 	  
	dyyyymm		   =  clsOpExp.Fyyyymm		 
	dYear				= year(dyyyymm) 
	dMonth			= month(dyyyymm) 
	iOpExpPartIdx   =  clsOpExp.FOpExpPartIdx 
	mLastMonthExp   =  clsOpExp.FLastMonthExp 
	mInExp		   =  clsOpExp.FInExp		 
	mOutExp		   =  clsOpExp.FOutExp		 
	mTotExp 	    =  clsOpExp.FTotExp 	 
	sOpExpPartName  =  clsOpExp.FOpExpPartName 
	iPartTypeIdx	= clsOpExp.FPartTypeIdx
	'권한체크----------------------------
	IF iOpExpPartIdx > 0  THEN	'운영비 사용처 구분값 잇을 경우에만 체크 
	clsOpExp.Fyyyymm		=	dyyyymm
	clsOpExp.FOpExpPartIdx	= iOpExpPartIdx 
	clsOpExp.FadminID 		= session("ssBctId") 
	blnWorker = clsOpExp.fnGetOpExpAuth  
	
	 IF blnWorker = 1  THEN blnReg =1 	'담당자이거나 어드민권한을 가진 경우 등록처리 가능 
	END IF
	'/권한체크----------------------------
	end if

	clsOpExp.FYYYYMM 		= dyyyymm 
	clsOpExp.FpartTypeidx 	= iPartTypeIdx  
	clsOpExp.FOpExpPartIdx 	= iOpExpPartIdx  
	clsOpExp.Farap_cd 		= iarap_cd  
	arrList = clsOpExp.fnGetOpExpDailySumList 
	
	
Set clsOpExp = nothing	

 '운영비관리 팀 구분 리스트		
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList  
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx  
	arrPart = clsPart.fnGetOpExppartAllList  
	END IF
Set clsPart = nothing 
%> 
 <script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!-- 
	//팀 관리 
	// ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");
    
    var _divName = "divP";
    
    function processAjax(){
        var reTxt = xmlHttp.responseText;   
        eval("document.all."+_divName).innerHTML = reTxt; 
    }
    
    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }
    
     
    function jsChangePart(iValue){   
        initializeURL('/admin/expenses/part/ajaxPart.asp?iPTIdx='+iValue);
    	startRequest();
    }	
    		   
//새로등록
function jsNewReg(){
	var winNew = window.open("about:blank","popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	document.frm.target = "popNew";
	document.frm.action = "regOpExp.asp";
	document.frm.submit();
	winNew.focus();
}  
 
//상세보기
function jsDetail(iST, ipt, ivalue){
	var ioidx, iccd;
	
	if (iST==2){
		iccd = "<%=iarap_cd%>"; 
		ioidx = ivalue;
	}else{
		ipt = "<%=iPartTypeIdx%>";
		ioidx = "<%=iOpExpPartIdx%>";
		iccd = ivalue;
	}
	 location.href = "dailyOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT="+ipt+"&selP="+ioidx+"&selA="+iccd+"&menupos=<%=menupos%>";
} 

function jsTotDetail(){
  location.href = "dailyOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>&selA=<%=iarap_cd%>&menupos=<%=menupos%>";
}

//선택 활성화
function jsSetST(iValue){
	if (iValue==1){
		document.frm.selPT.disabled = false;
		document.frm.selP.disabled = false; 
	}else{ 
		document.frm.selA.disabled = false;
	}
}

//검색
function jsSearch(){
	if(document.frm.rdoST[0].checked ==true){
		if(document.frm.selPT.value==0){
	 	alert("운영비사용처을 선택해주세요");
	 	return;
	 	}
	 	if(document.frm.selP.value==0){
	 	alert("운영비사용처을 선택해주세요");
	 	return;
	 	}
	}else{ 
	 	if(document.frm.selA.value==0){
	 	alert("수지항목을 선택해주세요");
	 	return;
	 	}
	}
	document.frm.target = "_self";
	document.frm.action = "dailySumOpExp.asp";
	document.frm.submit();
}

//리스트로 이동
function jsGoList(){
	location.href = "index.asp?selSY=<%=dyear%>&selSM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>&menupos=<%=menupos%>";
}

//프린트
	function jsPrint(){
		var winP = window.open("printDailySumOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>","popP","width=1024, height=600,resizable=yes, scrollbars=yes");
		winP.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<tr>
	<td>+ <a href="javascript:jsGoList();">운영비관리 월별 팀 리스트</a> > 월별 상세 리스트</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
				날짜:
					<select name="selY">
				<%For intY = Year(date()) To 2011 STEP -1%>
				<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>년
				 <select name="selM">
				<%For intM = 1 To 12%>
				<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
				<%Next%>
				</select>월
				&nbsp;&nbsp;
					<input type="radio" name="rdoST" value="1" <%IF isearchType =1 THEN%>checked<%END IF%> onClick="jsSetST(1);">운영비사용처:
					<select name="selPT" onChange="jsChangePart(this.value);" <%IF isearchType=2 THEN%>disabled<%END IF%>>
					<option value="0">--선택--</option>
					<% sbOptPartType arrType,ipartTypeIdx%>
					</select>
					<span id="divP"> 
					<select name="selP" onChange="jsSearch();" <%IF isearchType=2 THEN%>disabled<%END IF%>>
					<option value="0">--선택--</option>
					<% sbOptPart arrPart,iOpExpPartIdx%>
					</select> 
					</span> 
					&nbsp;&nbsp;
					<input type="radio" name="rdoST" value="2" <%IF isearchType =2 THEN%>checked<%END IF%>  onClick="jsSetST(2);">수지항목:	
					<select name="selA" <%IF isearchType=1 THEN%>disabled<%END IF%>>
					<option value="0">--선택--</option>
					<% sbOptAccount arrAccount, iarap_cd%> 
					</select>
				</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<tr>
	<td>
		<table border="0" cellpadding="5" cellspacing="0" width="100%">
		<tr>
		<%IF blnReg = 1 THEN%>
		<td><input type="button" class="button" value="운영비상세내역 등록" onClick="jsNewReg();"></td>
		<%END IF%>
		<%IF iOpExpPartIdx > 0 THEN%> 
		 <td align="right"><input type="button" class="button" value="프린트" onClick="jsPrint();"></td> 
		<%END IF%> 
		</table>
	</td>
</tr> 
<%IF isearchType =1 THEN%>
<tr>
	<td><table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">전월잔액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">지급액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">당월잔액</td>
			
		</tr>
		<tr> 
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mLastMonthExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mInExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mOutExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mTotExp,0)%></td>
		</tr>
		</table>
</td>
</tr>
<%END IF%>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">    
				<%IF iPartTypeIdx > 0 THEN%>
				<td>수지항목</td> 
				<%ELSE%>
				<td>운영비사용처</td>   
				<%END IF%> 
				<td>지급액</td>    
				<td>사용액</td>   
				<td>공급가액</td>   
				<td>부가세</td>  
				<td>건수</td>  	
				<td>링크</td>  	
			</tr>
			<%    Dim sumIn, sumOut, sumSup, sumVat,sumCnt
			IF isArray(arrList) THEN
				sumIn = 0
				sumOut = 0
				sumSup = 0
				sumVat = 0
				sumCnt = 0
				For intLoop = 0 To UBound(arrList,2)  
			 %>  
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><%IF isearchType="2" THEN%><%=arrList(7,intLoop)%> > <%END IF%><%=arrList(6,intLoop)%></td>
				<td><%=formatnumber(arrList(0,intLoop),0)%></td>
				<td><%=formatnumber(arrList(1,intLoop),0)%></td>
				<td><%=formatnumber(arrList(2,intLoop),0)%></td>
				<td><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td><%=formatnumber(arrList(4,intLoop),0)%></td>
				<td><a href="javascript:jsDetail('<%=isearchType%>','<%IF isearchType= 2 THEN%><%=arrList(8,intLoop)%><%END IF%>','<%=arrList(5,intLoop)%>')">>>상세보기</a></td>
			</tr>	
			<%	sumIn = sumIn + arrList(0,intLoop)
				sumOut = sumOut + arrList(1,intLoop)
				sumSup = sumSup + arrList(2,intLoop)	
				sumVat = sumVat + arrList(3,intLoop)
				sumCnt = sumCnt + arrList(4,intLoop)
			Next  
			ELSE%>
			<tr height="30" align="center" bgcolor="#FFFFFF">				
				<td colspan="7">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
			<tr height=30 align="center" bgcolor="<%=adminColor("sky")%>">	
				<td>총합</td>
				<td><%=formatnumber(sumIn,0)%></td>
				<td><%=formatnumber(sumOut,0)%></td>
				<td><%=formatnumber(sumSup,0)%></td>
				<td><%=formatnumber(sumVat,0)%></td>
				<td><%=formatnumber(sumCnt,0)%></td>
				<td><a href="javascript:jsTotDetail('<%=isearchType%>',<%IF isearchType= 1 THEN%>'<%=iPartTypeIdx%>','<%=iOpExpPartIdx%>'<%ELSE%>'','<iarap_cdcd%>'<%END IF%>)">>>상세보기</a></td>
			</tr>
		</table>	
	</td> 
</tr> 	 
</table> 
</body>
</html> 



	