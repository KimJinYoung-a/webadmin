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
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"--> 
<%
Dim clsPart,clsOpExp,arrPart, arrList, arrType, intLoop 
Dim clsAccount, arrAccount  
Dim dYear, dMonth, iPartTypeIdx, iOpExpPartIdx, iarap_cd
Dim intY, intM
Dim isearchType, ipartsn, sadminid
Dim iOpExpIdx,dyyyymm, mLastMonthExp,mInExp,mOutExp,mTotExp,sOpExpPartName
Dim blnAdmin, blnWorker, blnReg 
 	dYear			= requestCheckvar(Request("selY"),4)
 	dMonth			= requestCheckvar(Request("selM"),2)
 	isearchType		= requestCheckvar(Request("rdoST"),1) 
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10) 
 	IF isearchType = "" THEN isearchType =1
 	IF isearchType = 1 THEN
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10) 
 	ELSE
 	iarap_cd		= requestCheckvar(Request("selA"),10)
	END IF
 
	iOpExpIdx		= requestCheckvar(Request("hidOE"),10)
 	IF dYear = "" THEN dYear = year(date())
 	IF dMonth = "" THEN dMonth = month(date())	
 		dyyyymm =  dYear&"-"&Format00(2,dMonth) 
 	IF 	iOpExpPartIdx = "" THEN iOpExpPartIdx = 0
 	IF 	iPartTypeIdx = "" THEN iPartTypeIdx = 0
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
	dYear			= year(dyyyymm) 
	dMonth			= month(dyyyymm) 
	iOpExpPartIdx   =  clsOpExp.FOpExpPartIdx 
	mLastMonthExp   =  clsOpExp.FLastMonthExp 
	mInExp		   =  clsOpExp.FInExp		 
	mOutExp		   =  clsOpExp.FOutExp		 
	mTotExp 	    =  clsOpExp.FTotExp 	 
	sOpExpPartName  =  clsOpExp.FOpExpPartName 
	iPartTypeIdx	= clsOpExp.FPartTypeIdx
	end if
	clsOpExp.FYYYYMM 		= dyyyymm 
	clsOpExp.FPartTypeIdx 	= iPartTypeIdx  
	clsOpExp.FOpExpPartIdx 	= iOpExpPartIdx  
	clsOpExp.Farap_cd 		= iarap_cd  
	arrList = clsOpExp.fnGetOpExpDailySumList 
Set clsOpExp = nothing	

	
 '운영비관리 팀 구분 리스트		
Set clsPart = new COpExpPart  
	IF not blnAdmin THEN  '리스트 권한을 가진 사람을 제외하고 담당자와 담당부서  view 가능
		ipartsn  =  session("ssAdminPsn")
 		sadminid = 	session("ssBctId")
 	END IF	
	clsPart.FRectPartsn = ipartsn
	clsPart.FRectUserid = sadminid  
	arrType = clsPart.fnGetOpExpPartTypeCardList 
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx   
	arrPart = clsPart.fnGetOpExppartAllList   
	END IF   
Set clsPart = nothing 
%> 
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script> 
<script language="javascript">
<!--     
//팀 관리 
// =========================================================================================================
$(document).ready(function(){
	$("#selPT").change(function(){
		var iValue = $("#selPT").val();
		var url="/admin/expenses/part/ajaxPart.asp";
		 var params = "iPTIdx="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divP").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});

//새로등록
function jsNewReg(){
	var winNew = window.open("about:blank","popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	document.frm.target = "popNew";
	document.frm.action = "regOpExp.asp";
	document.frm.submit();
	winNew.focus();
}  
 
//상세보기
function jsDetail(iST, ivalue){
	var ioidx, iccd;
	
	if (iST==2){
		iccd = "<%=iarap_cd%>"; 
		ioidx = ivalue;
	}else{ 
		ioidx = "<%=iOpExpPartIdx%>";
		iccd = ivalue;
	}
	 location.href = "dailyOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP="+ioidx+"&selA="+iccd+"&menupos=<%=menupos%>";
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

    function jsChangePart(iValue){   
        initializeURL('/admin/expenses/part/ajaxPart.asp?iPTIdx='+iValue);
    	startRequest();
    	
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
	<td>+ <a href="javascript:jsGoList('index');">법인카드관리 월별 리스트</a> > 월별 상세 리스트</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td   width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
				날짜(청구일):
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
					<select name="selPT"  id="selPT"   class="select" <%IF isearchType=2 THEN%>disabled<%END IF%>>
					<option value="0">--선택--</option>
					<% sbOptPartType arrType,ipartTypeIdx%>
					</select>
					<span id="divP"> 
					<select name="selP"  id="selP" class="select" <%IF isearchType=2 THEN%>disabled<%END IF%>>
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
				<td    width="50" bgcolor="<%= adminColor("gray") %>">
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
		<%IF iOpExpPartIdx > 0 THEN%> 
		 <!--td align="right"><input type="button" class="button" value="프린트" onClick="jsPrint();"></td--> 
		<%END IF%> 
		</table>
	</td>
</tr>  
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">    
				<%IF iOpExpPartIdx > 0 THEN%>
				<td>수지항목</td> 
				<%ELSE%>
				<td>운영비사용처</td>   
				<%END IF%>  
				<td>사용액</td>   
				<td>공급가액</td>   
				<td>부가세</td>  
				<td>봉사료</td>
				<td>건수</td>  	
				<td>링크</td>  	
			</tr>
			<%    Dim  sumOut, sumSup, sumVat, sumSev,sumCnt
			IF isArray(arrList) THEN 
				sumOut = 0
				sumSup = 0
				sumVat = 0
				sumSev = 0
				sumCnt = 0
				For intLoop = 0 To UBound(arrList,2)  
			 %>  
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><%=arrList(6,intLoop)%></td>
				<td><%=formatnumber(arrList(0,intLoop),0)%></td>
				<td><%=formatnumber(arrList(1,intLoop),0)%></td>
				<td><%=formatnumber(arrList(2,intLoop),0)%></td>
				<td><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td><%=formatnumber(arrList(4,intLoop),0)%></td>
				<td><a href="javascript:jsDetail('<%=isearchType%>','<%=arrList(5,intLoop)%>')">>>상세보기</a></td>
			</tr>	
			<%	
				sumOut = sumOut + arrList(0,intLoop)
				sumSup = sumSup + arrList(1,intLoop)	
				sumVat = sumVat + arrList(2,intLoop)
				sumSev = sumSev + arrList(3,intLoop)
				sumCnt = sumCnt + arrList(4,intLoop) 
			Next  
			ELSE%>
			<tr height="30" align="center" bgcolor="#FFFFFF">				
				<td colspan="7">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
			<tr height=30 align="center" bgcolor="<%=adminColor("sky")%>">	
				<td>총합</td> 
				<td><%=formatnumber(sumOut,0)%></td>
				<td><%=formatnumber(sumSup,0)%></td>
				<td><%=formatnumber(sumVat,0)%></td>
				<td><%=formatnumber(sumSev,0)%></td>
				<td><%=formatnumber(sumCnt,0)%></td>
				<td><a href="javascript:jsTotDetail('<%=isearchType%>',<%IF isearchType= 1 THEN%>'<%=iOpExpPartIdx%>'<%ELSE%>'<%=iarap_cd%>'<%END IF%>)">>>상세보기</a></td>
			</tr>
		</table>	
	</td> 
</tr> 	 
</table> 
</body>
</html> 



	