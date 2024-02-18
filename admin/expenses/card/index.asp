<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비관리    리스트
' History : 2011.06.03 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim isUseSerp : isUseSerp = true

Dim clsPart, arrType,arrPart,clsOpExp
Dim arrList, intLoop
Dim dSYear, dSMonth,dEYear, dEMonth, iPartTypeIdx ,iOpExpPartIdx
Dim intY, intM
Dim blnAdmin, blnWorker ,blnReg, blnSearch, sadminid, ipartsn, department_id
Dim iState
dim sBankName,sBankAccNo,arrIK,intK

''// ===========================================================================
''관리자 = 마스터권한 or 경영지원팀
''
''담당자, 사용부서, 관리자 : 조회가능
''담당자, 관리자 : 작성가능
''// ===========================================================================

 	dSYear			= requestCheckvar(Request("selSY"),4)
 	dSMonth			= requestCheckvar(Request("selSM"),2)
 	dEYear			= requestCheckvar(Request("selEY"),4)
 	dEMonth			= requestCheckvar(Request("selEM"),2)
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx	= requestCheckvar(Request("iPS"),10)
 	iState			= requestCheckvar(Request("selSt"),1)

 	IF dSYear = "" THEN dSYear = year(dateadd("m",-1,date()))
 	IF dSMonth = "" THEN dSMonth = month(dateadd("m",-1,date()))
 	IF dEYear = "" THEN dEYear = year(date())
 	IF dEMonth = "" THEN dEMonth = month(date())
 	IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF iOpExpPartIdx ="" THEN iOpExpPartIdx = 0

 	'권한초기값 설정--------------
 	blnWorker = 0 '담당자
 	blnReg = 0 	'등록권한
  	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '어드민권한

  	IF blnAdmin THEN blnReg = 1 '어드민권한 있을 경우 등록처리 항상 가능

 '운영비관리 팀 구분 리스트
Set clsPart = new COpExpPart
	IF not blnAdmin THEN  '리스트 권한을 가진 사람을 제외하고 담당자와 담당부서  view 가능
		ipartsn  =  session("ssAdminPsn")
		department_id = GetUserDepartmentID("",session("ssBctID"))
 		sadminid = 	session("ssBctId")
 	END IF
	''clsPart.FRectPartsn = ipartsn
	clsPart.FRectDepartmentID = department_id
	clsPart.FRectUserid = sadminid
	arrType = clsPart.fnGetOpExpPartTypeCardListNew
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx
	arrPart = clsPart.fnGetOpExppartAllListNew
	END IF
Set clsPart = nothing


Set clsOpExp = new OpExp
	'운영비 리스트
	''clsOpExp.FRectPartsn = ipartsn
	clsOpExp.FRectDepartmentID = department_id
	clsOpExp.FRectUserid = sadminid
	clsOpExp.FSYYYYMM	= dSYear&"-"&Format00(2,dSMonth)
	clsOpExp.FEYYYYMM	= dEYear&"-"&Format00(2,dEMonth)
	clsOpExp.FPartTypeIdx	=iPartTypeIdx
	clsOpExp.FOpExpPartIdx	=iOpExpPartIdx
	clsOpExp.FState = iState
	arrList = clsOpExp.fnGetOpExpMonthlyList

    IF isArray(arrList) THEN
        sBankName = arrList(14,0)
        sBankAccNo = replace(arrList(15,0),"-","")
    END IF

    IF  sBankName <> "" and    sBankAccNo <>"" THEN
    clsOpExp.FRectBankNM    = sBankName
    clsOpExp.FRectBankAccNo = sBankAccNo
    arrIK = clsOpExp.fnGetIpkumList
    END IF
    
	'권한체크------------------------
	IF iOpExpPartIdx > 0  THEN	'운영비 사용처 구분값 잇을 경우에만 체크
	clsOpExp.FOpExpPartIdx	= iOpExpPartIdx
	clsOpExp.FadminID 		= session("ssBctId")
	blnWorker = clsOpExp.fnGetOpExpPartAuth '담당자 여부 확인
 
	IF  blnWorker =1  THEN	blnReg =1 '담당자이거나 어드민권한을 가진 경우 등록처리 가능
	END IF
	'/권한체크------------------------
	
Set clsOpExp = nothing

%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
//팀 관리
// =========================================================================================================
$(document).ready(function(){
	$("#selPT").change(function(){
		var iValue = $("#selPT").val();
		var url="/admin/expenses/part/ajaxDepartment.asp";
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

//파일등록
function jsNewRegFile(){
			var winF = window.open('/admin/expenses/card/popRegFile.asp?selP=<%=iOpExpPartIdx%>','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winF.focus();
	}

 //상세보기
 function jsDetail(sPage, dyear, dmonth, ipartypeidx, iopexppartidx){
 	location.href = sPage +".asp?selY="+dyear+"&selM="+dmonth+"&selPT="+ipartypeidx+"&selP="+iopexppartidx+"&menupos=<%=menupos%>";
 }

 	//전자결재 품의서 등록
	function jsRegEapp(dyyyymm, iOpexpPartidx, iPartTypeIdx){
		var winEapp = window.open("eappOpExp.asp?dyyyymm="+dyyyymm+"&hidP="+iOpexpPartidx+"&hidPT="+iPartTypeIdx,"popE","width=1200,height=600,scrollbars=yes,resizable=yes");
		winEapp.focus();
	}

	//전자결재 품의서 내용보기
	function jsViewEapp(reportidx,reportstate){
		var winEapp = window.open("/admin/approval/eapp/modeapp.asp?blnP=1&iRS="+reportstate+"&iridx="+reportidx,"popE","");
		winEapp.focus();
	}

	//상태변경처리
	function jsOpExpConfirm(strMsg,sY,sM,iOpExp,istate){
		if(confirm(strMsg)){
		document.frmC.hidOE.value = iOpExp;
		document.frmC.hidS.value = istate;
		document.frmC.selY.value = sY;
		document.frmC.selM.value = sM;
		document.frmC.submit();
		}
		}

	//검색
	function jsSearch(){
		document.frm.target = "_self";
		document.frm.action = "index.asp";
		document.frm.iPS.value = $("#selP").val();
		document.frm.submit();
	}

function jsLinkERP(frm){
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}
	//alert('작업중.. 10/17일 이전 작업하겠음.');
	//return;
	if (confirm('선택 내역을 ERP로 전송하시겠습니까?')){
	    frm.LTp.value="D";
	    frm.action="/admin/approval/payreqList/erpLink_Process.asp";
	    frm.submit();
	}
}

function jsLink_SERP_unlock(frm){
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    e.disabled=false;
		}
	}
}

function jsLinkERP_sERP(frm){
    alert('사용중지메뉴');
    return;
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}

	if (confirm('선택 내역을 ERP로 전송하시겠습니까?')){
	    frm.LTp.value="D";
	    frm.action="/admin/approval/payreqList/S_erpLink_Process.asp";
	    frm.submit();
	}
}


function jsMakeMonth(frm){
    var cstr = frm.selY.value+'-'+frm.selM.value+' 월별 데이터를 생성하시겠습니까?'
    if (confirm(cstr)){
        frm.submit();
    }
}

function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp){
    AnCheckClick(comp)
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>+ 법인카드관리 월별 리스트 </td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="index.asp" >
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iPS" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					청구일:
					<select name="selSY" class="select">
					<%For intY = Year(date()) To 2011 STEP -1%>
					<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dSYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					</select>년
					 <select name="selSM"  class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dSMonth) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
					</select>월
					-
					<select name="selEY" class="select">
					<%For intY = Year(date()) To 2011 STEP -1%>
					<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dEYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					</select>년
					 <select name="selEM" class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dEMonth) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
					</select>월
						&nbsp;&nbsp;
				  운영비사용처:
					<select name="selPT"  id="selPT"   class="select">
					<option value="0">--선택--</option>
					<% sbOptPartType arrType,ipartTypeIdx%>
					</select>
					<span id="divP">
					<select name="selP"  id="selP" class="select">
					<option value="0">--선택--</option>
					<% sbOptPart arrPart,iOpExpPartIdx%>
					</select>
					</span>
					&nbsp;&nbsp;
					상태:
					<select name="selSt" id="selSt" class="select">
					<% SbOptState iState%>
					</select>
				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%IF    blnReg =1    THEN %>
<tr>
    <td> + 계좌 마지막 입출금내역
        <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
            <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                <td>IDX</td>
            	<td>은행명</td>
            	<td>계좌번호</td>
            	<td>마지막입출금일</td>
            	<td>사용처</td>
            	<td>거래구분</td>
              	<td>입금금액</td>
            	<td>출금금액</td>
            	<td>잔액</td>
            	<td>업데이트시간</td> 
            </tr>
            
            <%IF isArray(arrIK) THEN
                For intK = 0 To UBound(arrIK,2)
                %>
             <tr height="30" align="center" bgcolor="#FFFFFF">
                <td><%=arrIK(0,intK)%></td>
            	<td><%=arrIK(1,intK)%></td>
            	<td><%=arrIK(2,intK)%></td>
            	<td><%=arrIK(3,intK)%></td>
            	<td><%=arrIK(4,intK)%></td>
            	<td><%=arrIK(5,intK)%></td>
              	<td><%IF arrIK(6,intK) = 2 THEN%><%=formatnumber(arrIK(7,intK),0)%><%ELSE%>0<%END IF%></td>
            	<td><%IF arrIK(6,intK) = 1 THEN%><%=formatnumber(arrIK(7,intK),0)%><%ELSE%>0<%END IF%></td>
            	<td><%=formatnumber(arrIK(8,intK),0)%></td>
            	<td><%=arrIK(9,intK)%></td> 
            </tr>    
            <%  Next 
            ELSE%>
            <tr  height="30" align="center" bgcolor="#FFFFFF">
                <td colspan="10">등록된 내역이 존재하지 않습니다.</td>
            </tr>
            <%END IF%>
        </table>
    </td>
</tr>
<%END IF%>
<tr>
	<td>
	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	    	<%IF  FALSE and blnReg =1    THEN%>
	    	<td>
					<input type="button" class="button" value="운영비상세내역 신규등록" onClick="jsNewReg();">
					<input type="button" class="button" value="파일등록" onClick="jsNewRegFile();">
			</td>
	    	<%END IF%>
	    	<% IF (blnAdmin) THEN %>
	    	<td align="left" ><input type="button" class="button" value="월별내역생성(<%=dSYear%>-<%=dSMonth%>)" onClick="jsMakeMonth(frmMnAct);"></td>
	        <td align="right" >
	            <% if (isUseSerp) then %>
	            <!-- 사용안함.
	                <input type="button" value="sERP 전송" onClick="jsLinkERP_sERP(frmAct)"> 
	             --> 
	            <% else %>
	            <input type="button" class="button" value="ERP 전송" onClick="jsLinkERP(frmAct);">
	            
    	        <% if session("ssBctID")="icommang" or session("ssBctID")="ju1209XXX" then %>
    	            <font color=red>sERP[</font> 
    	            <input type="button" value="unlock" onClick="jsLink_SERP_unlock(frmAct)">
                    <input type="button" value="sERP 전송" onClick="jsLinkERP_sERP(frmAct)"> 
                    <font color=red>]</font>
                <% end if %>
                <% end if %>
	        </td>
	        
	        
            
	      <% END IF %>
	    </tr>
	    </table>
	</td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <Form name="frmAct" method="post" action="/admin/approval/payreqList/erpLink_Process.asp">
		    <input type="hidden" name="LTp" value="C">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			    <% IF (blnAdmin) THEN %>
			    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
			    <% END IF %>
				<td>날짜</td>
				<td>구분</td>
				<td>운영비사용처</td>
				<td>당월사용액</td>
				<td>상태</td>
				<%IF blnReg=1  THEN%>
				<td>처리</td>
				<%END IF%>
				<td>경영지원팀<br>서류확인</td>
				<% IF (blnAdmin) THEN %><td>ERP<br>연동상태</td>  <% END IF %>
				<td>상세내역보기</td>
			</tr>
			<%   dim dRectY, dRectM
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				dRectY = year(arrList(1,intLoop))
				dRectM = month(arrList(1,intLoop))
			 %>
			<tr height=30 align="center" bgcolor="#FFFFFF">
			    <% IF (blnAdmin) THEN %>
			    <td><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%= CHKIIF((arrList(10,intLoop)="9") and (arrList(13,intLoop)=10) and (arrList(1,intLoop)>="2012-09") or (TRUE),"","disabled") %> ></td>
			    <% END IF %>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=fnGetPartTypeDesc(arrList(13,intLoop))%></td>
				<td><%=arrList(7,intLoop)%> </td>
				<td><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td><%=fnGetStateDesc(arrList(10,intLoop))%></td>
				<%IF ( blnReg=1 ) THEN%>
				<td>
				    <%=arrList(10,intLoop)%>
				    <% IF (arrList(10,intLoop)>9) THEN %>
				    수정불가 <input type="button" class="button" value="품의서보기 >" onClick="jsViewEapp('<%=arrList(8,intLoop)%>','<%=arrList(9,intLoop)%>')">
				    <% ELSE %>
    					<%IF (arrList(10,intLoop) = 1 and blnWorker = 1) OR (arrList(10,intLoop) >0 and arrList(10,intLoop) < 9 and blnAdmin ) THEN %>
    					<input type="button" class="button" style="color:gray;" value="< 작성중" onClick="jsOpExpConfirm('작성중 상태로 변경하시겠습니까?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,'<%=arrList(0,intLoop)%>',0)">
    					<%END IF%>
    					<%IF isNull(arrList(8,intLoop)) and  (arrList(10,intLoop) = 1 or arrList(10,intLoop) = 5) THEN %>
    						<input type="button" class="button"   value="품의서작성 >" onClick="jsRegEapp('<%=arrList(1,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(13,intLoop)%>')">
    					<%ELSEIF not isNull(arrList(8,intLoop))  THEN%>
    						<input type="button" class="button" value="품의서보기 >" onClick="jsViewEapp('<%=arrList(8,intLoop)%>','<%=arrList(9,intLoop)%>')">
    					<%ELSE%>
    						<%IF blnAdmin THEN%>
    						<input type="button" class="button" value="작성완료 >" onClick="jsOpExpConfirm('작성완료하시겠습니까?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,'<%=arrList(0,intLoop)%>',1)">
    						<%eND IF%>
    					<%END IF%>
					<% END IF %>
				</td>
				<%END IF%>
				<td>
					<%if  blnAdmin  and  (arrList(10,intLoop) >=7 ) and  (arrList(10,intLoop) <10 ) then%>
					<input type="radio" name="rdoC<%=arrList(0,intLoop)%>" value="1" <%IF arrList(10,intLoop) = 9 THEN%>checked<%END IF%> onClick="jsOpExpConfirm('서류확인상태로 변경하시겠습니까?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,<%=arrList(0,intLoop)%>,9)"><font color="blue">Y</font>
					<input type="radio" name="rdoC<%=arrList(0,intLoop)%>" value="0" <%IF arrList(10,intLoop) <> 9 THEN%>checked<%END IF%>  onClick="jsOpExpConfirm('서류확인을 취소하시겠습니까?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,<%=arrList(0,intLoop)%>,7)"><font color="red">N</font>
					<%else%>
						<%IF arrList(10,intLoop) >= 9 THEN %>
							<font color="blue">Y</font></a>
						<%ELSE%>
								<font color="red">N</font></a>
						<%END IF%>
					<%end if%>
				</td>
				<% IF (blnAdmin) THEN %>
				<td>
				    <% if Not IsNULL(arrList(12,intLoop)) then %>
				    [<%= arrList(11,intLoop) %>]<%= arrList(12,intLoop) %>
	                <% end if %>
  				</td>
				<% END IF %>
				<td>
					<a href="javascript:jsDetail('dailySumOpExp','<%=dRectY%>','<%=dRectM%>','<%=arrList(13,intLoop)%>','<%=arrList(2,intLoop)%>')">[월별상세]</a>
					<a href="javascript:jsDetail('dailyOpExp','<%=dRectY%>','<%=dRectM%>','<%=arrList(13,intLoop)%>','<%=arrList(2,intLoop)%>')">[일별상세]</a>
				</td>
			</tr>
		<%
			Next
			ELSE%>
			<tr height="30" align="center" bgcolor="#FFFFFF">
				<td colspan="13">등록된 내용이 없습니다.</td>
			</tr>
			<%END IF%>
			</form>
		</table>
	</td>
</tr>
</table>
<form name="frmC" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="C">
<input type="hidden" name="hidOE" value="">
<input type="hidden" name="hidS" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="selSY" value="<%= dSYear %>">
<input type="hidden" name="selSM" value="<%= dSMonth %>">
<input type="hidden" name="selY" value="" >
<input type="hidden" name="selM" value="">
<input type="hidden" name="selP" value="<%= iOpExpPartIdx %>">
<input type="hidden" name="selPT" value="<%= iPartTypeIdx %>">
</form>
<form name="frmMnAct" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="M">
<input type="hidden" name="hidOE" value="">
<input type="hidden" name="hidS" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="selY" value="<%= dSYear %>">
<input type="hidden" name="selM" value="<%= dSMonth %>">
<input type="hidden" name="selP" value="<%= iOpExpPartIdx %>">
<input type="hidden" name="selPT" value="<%= iPartTypeIdx %>">
</form>
</body>
</html>
