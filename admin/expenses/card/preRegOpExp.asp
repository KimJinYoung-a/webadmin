<%@ language=vbscript %>
<% option explicit  %>
<%
'###########################################################
' Description : 운영비관리  내용
' History : 2011.05.30 정윤정  생성
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
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim sMode, selA,arap_nm
Dim clsPart, clsAccount, arrAccount ,clsOpExp, clsPartMoney
Dim arrList, intLoop
Dim intY, dYear, intM, dMonth
Dim  dYYYYMM,iPartTypeIdx,iOpExpPartIdx, iOpExpDailyIdx, dauthDate,msevExp,blndeducttype
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim arrUsePart,sOpExpPartName, sPartTypeName
Dim  iarap_cd,minExp,mOutExp,sOpExpObj,sDetailCOnts,sbizsection_cd,sbizsection_nm,msupExp,mvatExp,sauthNo ,blnIntOut
Dim blnAdmin, blnWorker ,blnReg
Dim  ipartsn,sadminid
Dim idefaultArap_cd, intA

	dYear = requestCheckvar(Request("selY"),10)
	IF dYear = "" THEN dYear = year(date())
	dMonth= requestCheckvar(Request("selM"),10)
	IF dMonth = "" THEN dMonth = month(date())
	iPartTypeIdx = requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx = requestCheckvar(Request("selP"),10)
 		IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF iOpExpPartIdx = "" THEN iOpExpPartIdx = 0

 	iOpExpDailyIdx = 	requestCheckvar(Request("hidOED"),10)
 	IF iOpExpDailyIdx = "" THEN iOpExpDailyIdx = 0


 	'권한초기값 설정--------------
 	blnWorker = 0 '담당자
 	blnReg = 0 	'등록권한
 	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '어드민권한
 	IF blnAdmin THEN blnReg = 1 '어드민권한 있을 경우 등록처리 항상 가능
 	IF not blnAdmin THEN  '리스트 권한을 가진 사람을 제외하고 담당자와 담당부서  view 가능
		ipartsn  =  session("ssAdminPsn")
 		sadminid = 	session("ssBctId")
 	END IF
 '운영비 사용처
 	IF iOpExpPartIdx > 0 THEN
Set clsPart = new COpExpPart
		clsPart.FOpExpPartidx = iOpExpPartIdx
		clsPart.fnGetOpExpPartName
		sOpExpPartName =clsPart.FOpExpPartName
		sPartTypeName  =clsPart.FPartTypeName

Set clsPart = nothing
Set clsPart = new COpExpPart
		clsPart.FOpExpPartidx = iOpExpPartIdx
        clsPart.fnGetOpExpPartData
        idefaultArap_cd=clsPart.Farap_cd
Set clsPart = nothing
 END IF

'운영비 데일리 리스트
set clsOpExp = new OpExp
	clsOpExp.FSAuthDate 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FEAuthDate 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FPartTypeIdx = iPartTypeIdx
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx
	clsOpExp.FRectPartsn = ipartsn
	clsOpExp.FRectUserid = sadminid
	arrList = clsOpExp.fnGetOpExpDailyNoSetList
	iTotCnt = clsOpExp.FTotCnt

	clsOpExp.FadminID = session("ssBctId")
	clsOpExp.FPart_sn = session("ssAdminPsn")
  blnWorker = clsOpExp.fnGetOpExpPartAuth
  IF blnWorker = 1 THEN blnReg = 1
	IF blnReg=0 THEN
		set clsOpExp = nothing
			Call Alert_close ("수정권한이 없습니다. 확인 후 다시 시도해주세요")
		response.end
	END IF
IF iOpExpDailyIdx > 0 THEN
	sMode ="U"
	clsOpExp.FOpExpDailyIdx=iOpExpDailyIdx
	clsOpExp.fnGetOpExpDailyData
	dYYYYMM 		= clsOpExp.FYYYYMM
	dauthDate		= clsOpExp.Fauthdate
	iOpExpPartIdx = clsOpExp.FOpExpPartIdx
	iarap_cd			= clsOpExp.Farap_cd
	mOutExp 			= clsOpExp.FOutExp
	sOpExpObj 		= clsOpExp.FOpExpObj
	sDetailCOnts 	= clsOpExp.FDetailCOnts
	sbizsection_cd= clsOpExp.Fbizsection_cd
	msupExp 			= clsOpExp.FsupExp
	mvatExp 			= clsOpExp.FvatExp
	msevExp				= clsOpExp.FsevExp
	sauthNo				= clsOpExp.FauthNo
	blndeducttype	= clsOpExp.Fdeducttype
	blnIntOut			= clsOpexp.Finouttype
	sbizsection_nm= clsOpExp.Fbizsection_nm

END IF
set clsOpExp = nothing

 IF isNull(blndeducttype) THEN blndeducttype = False

 '수지항목 리스트
set clsAccount = new COpExpAccount
	clsAccount.FOpExpPartIdx = iOpExpPartIdx
	arrAccount = clsAccount.fnGetArapRegList
set clsAccount = nothing

'' 기본 수지 항목 입력
if (iarap_cd="") or (iarap_cd="0")then
    if (idefaultArap_cd="625") or (idefaultArap_cd="640") then
        iarap_cd = idefaultArap_cd
    end if
end if
%>
 <script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript" src="/js/datetime.js"></script>
<script language="javascript">
<!--
	//검색
	function jsSearch(){
		document.frmReg.action = "preRegOpExp.asp";
		document.frmReg.submit();
	}

 	//등록
 	function jsAddOpExp(){
 	  if((document.frmReg.selA.value==0)){
 		alert("수지항목을 선택해주세요");
 		return;
 		}
      <% 'if (idefaultArap_cd="928") then %> //주말야근식대
// 주석처리. 재무팀 최현희 요청
//      if((document.frmReg.selA.value!='<%'=idefaultArap_cd%>')){
//        alert('지정된 수지항목만 사용 가능합니다.');
//        document.frmReg.selA.focus();
//        return;
//      }
      <% 'end if %>
      <% 'if (idefaultArap_cd="625")  then %> //조직활성화비, 조식대 
// 주석처리. 재무팀 최현희 요청
//      if(!(document.frmReg.selA.value==625 || document.frmReg.selA.value==927 || document.frmReg.selA.value ==940)){
//        alert('지정된 수지항목만 사용 가능합니다.');
//        document.frmReg.selA.focus();
//        return;
//      }
      <% 'end if %>
 		document.frmReg.action ="procOpExp.asp"
 		document.frmReg.submit();
 	}

 	//수정
 	function jsModOpExp(idx){
// 	      <% if (idefaultArap_cd="625") or (idefaultArap_cd="640") then %>
//          if((document.frmReg.selA.value!='<%=idefaultArap_cd%>')){
//            alert('지정된 수지항목만 사용 가능합니다.');
//            document.frmReg.selA.focus();
//            return;
//          }
//          <% end if %>

 		document.frmReg.hidOED.value= idx;
 		document.frmReg.action ="preRegOpExp.asp" ;
 		document.frmReg.submit();
 	}

 	//삭제
 	function jsDelOpExp(idx){
 		if(confirm("삭제하시겠습니까?")){
 			document.frmDel.hidOED.value = idx;
 			document.frmDel.submit();
 		}
 	}


	//취소
	function jsReset(){
		document.frmReg.hidOED.value= 0;
		document.frmReg.action = "preRegOpExp.asp";
		document.frmReg.submit();
	}

	// 수지항목
	function jsGetarap_cd(iOpExpPartIdx){
		if (iOpExpPartIdx==''){
			alert('검색키가 없습니다.;');
			return;
		}

		var winarap_cdP = window.open('/admin/linkedERP/Biz/poparap_cdone.asp?selP='+ iOpExpPartIdx +'&menupos=<%= menupos %>','poparap_cd','width=600, height=500, resizable=yes, scrollbars=yes');
		winarap_cdP.focus();
	}

	// 수지항목 등록
	function jsSetarap_cd(arap_cd, arap_nm){
		document.frmReg.selA.value = arap_cd;
		document.frmReg.arap_nm.value = arap_nm;
	}

  	//자금관리부서 선택
	function jsGetPart(){
			var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winP.focus();
	}

	//자금관리부서 등록
	function jsSetPart(sBcd, sBnm){
			document.frmReg.sBcd.value = sBcd;
			document.frmReg.sBnm.value = sBnm;
	}
	//사용내역 파일등록
	function jsSetFile(){
			var sYear = document.frmReg.selY.options[document.frmReg.selY.selectedIndex].value;
			var sMonth = document.frmReg.selM.options[document.frmReg.selM.selectedIndex].value;
			var winF = window.open('/admin/expenses/opexp/popRegFile.asp?selY='+sYear+'&selM='+sMonth+'&selP=<%=iOpExpPartIdx%>','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winF.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" >
<form name="frmDel" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmReg" method="get" action="procOpExp.asp">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="hidOED" value="<%=iOpExpDailyIdx%>">
		<input type="hidden" name="iCP" value="<%=iCurrpage%>">
		<input type="hidden" name="hidNS" value="Y">
		<input type="hidden" name="hidRU" value="preRegOpExp.asp">
		<input type="hidden" name="mO"  value="<%=moutExp%>">
		<input type="hidden" name="mSP"  value="<%=msupExp%>">
		<input type="hidden" name="mV"  value="<%=mvatExp%>">
		<input type="hidden" name="mSV" value="<%=msevExp%>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
			<td align="left">
				승인일:
				 <%IF sMode="U" THEN%>
				<input type="hidden" name="selY" value="<%=dYear%>">
				<input type="hidden" name="selM" value="<%=dMonth%>">
				<%=dYear%>년 <%=dMonth%>월
				<%ELSE%>
				<select name="selY" class="select">
				<%For intY = Year(date()) To 2011 STEP -1%>
				<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>년
				 <select name="selM" class="select">
				<%For intM = 1 To 12%>
				<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
				<%Next%>
				</select>월
				<%END IF%>
				&nbsp;&nbsp;
				 운영비사용처:&nbsp;
				   <%=sPartTypeName%> > <%=sOpExpPartName%>
				  <input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
				<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
				</td>
				<%IF sMode="I" THEN%>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
				<%END IF%>
			</td>
		</tr>
		</table>
	</td>
</tr>
 <%IF  sMode="U"  THEN%>
<%IF ( blnReg = 1  ) THEN%>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td>수지항목</td>
			<td>업체명</td>
			<td>적요(상세내역)</td>
			<td>금액</td>
			<td>승인번호</td>
			<td>공제여부</td>
			<td>사용부서</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>
				<%
				If isArray(arrAccount) THEN
					For intA = 0 To UBound(arrAccount,2)
						IF Cstr(arrAccount(0,intA)) = Cstr(iarap_cd) THEN
							selA=arrAccount(0,intA)
							arap_nm=chkIIF(arrAccount(2,intA),"[사용]","[지급]") & arrAccount(1,intA)
						end if
					Next
				END IF
				%>
				<input type="hidden" name="selA" value="<%= selA %>">
				<input type="text" name="arap_nm" size="20" value="<%= arap_nm %>" class="text_ro" readonly>
				<a href="#" onclick="jsGetarap_cd('<%= iOpExpPartIdx %>'); return false;"><img src="/images/icon_search.jpg" border="0"></a>
			</td>
			<td><%=sOpExpObj%></td>
			<td><input type="text" name="sDC" size="50" maxlength="200" value="<%=sDetailCOnts%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><%=formatnumber(moutExp,0)%></td>
			<td><%=sauthNo%></td>
			<td><input type="radio"  name="rdoD" value="1" <%IF blndeducttype THEN%>checked<%END IF%>>Y &nbsp;
				 <input type="radio"  name="rdoD" value="0" <%IF not blndeducttype THEN%>checked<%END IF%>>N</td>
			<td><input type="hidden" name="sBcd" value="<%=sbizsection_cd%>"><input type="text" name="sBnm" size="10" value="<%=sbizsection_nm%>" class="text_ro" readonly>	<a href="javascript:jsGetPart();"><img src="/images/icon_search.jpg" border="0"></a></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td align="center">
			<input type="button" class="button" value="수정" style="width:80px;color:blue;" onClick="jsAddOpExp();">
			<input type="button" class="button" value="취소" style="width:80px;" onClick="jsReset();">
			</td>
		</tr>
		</table>
	</td>
	</form>
</tr>
<%ELSE%>
<tr>
	<td> <font color="red">- 작성완료되어 등록이 불가능하거나 등록 권한이 없습니다.</font></td>
</tr>
<%END IF%>
	<%END IF%>
<tr>
	<td>
		<div id="divList" style="height:600px;overflow:scroll;">
		<b> [ <%=dYear%>년 <%=dMonth%>월 법인카드사용 상세내역 - <%=sPartTypeName%> > <%=sOpExpPartName%>   ]</b>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
						<td width="50">순번</td>
						<td width="50">승인일</td>
						<td>운영비사용처</td>
						<td>수지항목</td>
						<td>업체명</td>
						<td>적요(상세내역)</td>
						<td>사용액</td>
						<td>공급가액</td>
						<td>부가세</td>
						<td>봉사료</td>
						<td>승인번호</td>
						<td>과세유형</td>
						<td>국내/외</td>
						<td>사용부서</td>
						<td>공제여부</td>
						<td width="100">처리</td>
					</tr>
				<%
					Dim  totOutExp, sumOutExp, iNum, sumSupExp, sumVatExp, sumSevExp, totSupExp, totVatExp, totSevExp
					totOutExp = 0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					sumSevExp=0
					totSupExp=0
					totVatExp=0
					totSevExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					 %>
					<tr height=30 bgcolor="<%IF Cstr(arrList(0,intLoop))= Cstr(iOpExpDailyIdx) THEN%><%=adminColor("green")%><%ELSE%><%= CHKIIF(arrList(22,intLoop)=0,"#CCCCCC","#FFFFFF") %><%END IF%>">
						<td align="center"><%=iNum%></td>
						<td align="center"><%=formatdate(arrList(2,intLoop),"0000-00-00")%></td>
						<td align="center"><%=arrList(15,intLoop)%></td>
						<td align="center"><%=arrList(5,intLoop)%></td>
						<td><%=arrList(11,intLoop)%></td>
						<td><%=arrList(12,intLoop)%></td>
						<td align="right"><%=formatnumber(arrList(6,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(7,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td>
						<td align="center"><%=arrList(10,intLoop)%></td>
						<td align="center"><%=arrList(16,intLoop)%></td>
						<td align="center"><%IF arrList(19,intLoop)=1 THEN%>국내<%ELSE%>국외<%END IF%></td>
						<td align="center"><%=arrList(14,intLoop)%></td>
						<td align="center"><%IF arrList(17,intLoop) THEN%><font color="red">Y</font><%ELSE%><font color="blue">N</font><%END IF%></td>
						<td align="center">
						<% if IsNULL(arrList(21,intLoop)) then %>
						<%IF blnReg = 1 THEN%>
						    <% if (arrList(22,intLoop)<>0) then %>
							<input type="button" class="button" value="수정" onClick="jsModOpExp(<%=arrList(0,intLoop)%>);">
							<% end if %>
							<%IF blnAdmin THEN%>
							<% if (arrList(22,intLoop)=0) then %>
						    <!-- input type="button" class="button" value="복구" onClick="jsLiveOpExp(<%=arrList(0,intLoop)%>)" -->
						    <% else %>
							<input type="button" class="button" value="삭제" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
							<% end if %>
							<% end if %>
						<%END IF%>
						<% else %>
						    <%= arrList(21,intLoop) %>
						<% end if %>
						</td>
					</tr>
					<%
					  iNum = iNum + 1
			 	Next
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">
						<td colspan="16">등록된 내용이 없습니다.</td>
					</tr>
					<%END IF%>

				</table>
			</td>
		</tR>
		</div>
	</td>
</tr>
</table>
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->



