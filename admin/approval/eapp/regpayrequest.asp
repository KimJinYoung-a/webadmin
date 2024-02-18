<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 결제요청서 등록
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim clsPay,clsMem,clseapp, clscomm,clsPM
Dim ireportidx,ipayrequestIdx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate ,scust_cd
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_use_cd,sacc_nm,  pcomment
Dim spartname ,sscmLink
Dim chkPayRequest
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate ,iBizNo
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrProc,arrPM,arrPart
Dim intA, intC, intF, intR, intRA, intP , iAuthCount, intPM, intPart
Dim blnMod
Dim mSumPrice, mSumRealPrice
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2), pmistate(2)
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	,scustnm,spayrequesttitle
Dim arrFName,arrF, sFName, intF2,intF3, iCount
Dim sMode
Dim iRectMenu ,sRectAuthId
Dim ipaytype, sCurrencyType, sCurrencyPrice,sACC_GRP_CD
Dim idepartmentid,sdepartmentname, icid1, icid2, icid3, icid4

 iRectMenu = requestCheckvar(Request("iRM"),10)
ireportidx 		=  requestCheckvar(Request("iridx"),10)
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
 blnMod = 0
 sMode = "I"
 IF ipayrequestIdx = "" THEN ipayrequestIdx = 0 '등록된 요청서가 없을 경우 품의서 폼의 내용 가져와서 default로 뿌려준다.
	sRectAuthId =  session("ssBctId")
'결재 기본 폼 정보 가져오기
set clsPay = new CPayRequest
	clsPay.Freportidx = ireportidx
	clsPay.FpayrequestIdx = ipayrequestIdx
	IF ipayrequestIdx <= 0 THEN '신규등록일때만 체크
		chkPayRequest = clsPay.fnCheckPayRequest

		IF chkPayRequest = 0 THEN
			set clsPay = nothing
	%>
	<!-- #include virtual="/lib/db/dbclose.asp" -->
	<%		Alert_return "결제요청서 등록이 불가능합니다. 데이터를 확인해주세요"
			response.end
		END IF
	ELSE
		sMode = "U"
	END IF

	clsPay.fnGetPayRequestData
	iarap_cd		 		= clsPay.Farap_cd
	sreportName      = clsPay.FreportName
	mreportPrice     = clsPay.FreportPrice
	iscmlinkno       = clsPay.Fscmlinkno
	sbigo            = clsPay.Fbigo
	ireportstate     = clsPay.Freportstate
	sadminid         = clsPay.Fadminid
	sarap_nm     		 = clsPay.Farap_nm
	sacc_cd					 = clsPay.Facc_cd
	sacc_use_cd			 = clsPay.Facc_use_cd
	sacc_nm		       = clsPay.Facc_nm

	dpayrequestdate   = clsPay.Fpayrequestdate
	mpayrequestprice  = clsPay.Fpayrequestprice
	iinBank           = clsPay.FinBank
	saccountNo        = clsPay.FaccountNo
	saccountHolder    = clsPay.FaccountHolder
	dpaydate          = clsPay.Fpaydate
	ioutBank          = clsPay.FoutBank
	dpayrealdate      = clsPay.Fpayrealdate
	mpayrealprice     = clsPay.Fpayrealprice
	syyyymm           = clsPay.Fyyyymm
	blnTakeDoc        = clsPay.FisTakeDoc
	ipayrequeststate  = clsPay.Fpayrequeststate
	dregdate          = clsPay.Fregdate
  pcomment					= clsPay.FpayComment
	scust_cd					= clsPay.Fcust_cd
	scustnm						= clsPay.Fcust_nm
	spayrequesttitle	= clsPay.FpayRequestTitle
	iBizNo						= clsPay.FBiz_no
	ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	sACC_GRP_CD				= clsPay.FACC_GRP_CD

	IF ipayrequestIdx = 0   THEN
		ipayrequestIdx = -1 '품의서 폼인지 결제요청서 폼인지 체크를 위해 (0=품의서)
		clsPay.FpayrequestIdx = ipayrequestIdx
	END IF
	'//기결제리스트
	arrProc			= clsPay.fnGetProcPayRequestList

	IF ipayrequestIdx > 0 THEN
	'//증빙서류
	clsPay.fnGetEappPayDoc
	ipayDocIdx   = clsPay.FpayDocIdx
	ipaydockind  = clsPay.Fpaydockind
	svatkind  	  = clsPay.Fvatkind
	dissuedate   = clsPay.Fissuedate
	sitemname  	= clsPay.Fitemname
	mtotprice 	  = clsPay.Ftotprice
	msupplyprice = clsPay.Fsupplyprice
	mvatprice  	= clsPay.Fvatprice
	setaxkey  	  = clsPay.Fetaxkey
	sDocbigo  	  = clsPay.FDocbigo
	sattachfile  = clsPay.Fattachfile
END IF
set clsPay = nothing


'결재라인, 코멘트, 파일 리스트 가져오기
set clseapp = new CEApproval
	clseapp.Freportidx 		= ireportidx
	clseapp.FpayrequestIdx = ipayrequestIdx

	IF ipayrequestIdx > 0  THEN
	arrAuth			= clseapp.fnGetAuthLineList
	END IF

	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList
	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing

 '결재선 리스트 지정
	IF isArray(arrAuth) THEN
		For intA = 0 To UBound(arrAuth,2)
			pmuserid(intA)  = arrAuth(2,intA)
			pmusername(intA)= arrAuth(7,intA)
		pmjobname(intA) = arrAuth(10,intA)
			pmstate(intA)	= fnGetPayAuthState(arrAuth(3,intA),intA+1)
			pmistate(intA)  = arrAuth(3,intA)
			pmdate(intA)	= arrAuth(6,intA)
			IF pmdate(intA) <> "" THEN pmdate(intA)	= formatdate(pmdate(intA),"0000-00-00")
		Next
	ELSE
		'재무회계팀 결제요청서 처리자정보
		Set clsPM	= new CPayManager
			clsPM.FisDef = 1
			arrPM	= clsPM.fnGetPayManager
		Set clsPM 	= nothing

		IF isArray(arrPM) THEN
			For intP = 0 To UBound(arrPM,2)
			pmuserid(intP)  = arrPM(1,intP)
			pmusername(intP)= arrPM(3,intP)
			pmjobname(intP) = arrPM(6,intP)
			pmistate(intP) = 0
			pmstate(intP)	= fnGetPayAuthState(0,intP+1)
			pmdate(intP)	= "&nbsp;"
			Next
		END IF
	END IF


'부서명 가져오기
set clsMem = new CTenByTenMember
	clsMem.Fuserid = session("ssBctId")
	clsMem.fnGetDepartmentInfo
	idepartmentid		= clsMem.Fdepartment_id
 	sdepartmentname = clsMem.FdepartmentNameFull
 	icid1						= clsMem.Fcid1
 	icid2						= clsMem.Fcid2
 	icid3						= clsMem.Fcid3
 	icid4						= clsMem.Fcid4
 set clsMem = nothing

'--------------------------------------------------
'문서 수정가능여부
 IF   ipayrequeststate = "" or isnull(ipayrequeststate) THEN ipayrequeststate = 0
 IF ( ipayrequeststate = 0  or ipayrequeststate=5) and sadminid = session("ssBctId")   THEN
	blnMod = 1
 END IF
%>
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="eapp.js"></script>
<script type="text/javascript">

	  //내용수정
  function jsPayReqUpdate(sMode){
  	var strMsg;
  	if(sMode=="D"){
  		strMsg = "삭제";
  	}else{
  		strMsg = "내용저장";
  	}

  	if(confirm(strMsg+" 하시겠습니까?")){
  		$("input[name='sFileP[]']").each( function(index,elem) {
		     var a = $(elem).val();
		     if( document.frm.sFile.value==""){
		     	document.frm.sFile.value = a;
		    }else{
		     document.frm.sFile.value = document.frm.sFile.value + ","+a;
		   }
		  });
  			document.frm.hidPRS.value=0;
  			document.all.mprp.value = document.all.mprp.value.replace(/\,/g,"");
				if(typeof(document.all.mPM) !="undefined"){
			  	if(typeof(document.all.mPM.length)!="undefined"){
			  		for(i=0;i<document.all.mPM.length;i++){
					 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
					 		if(document.frm.mP.value ==""){
							 			document.frm.mP.value =document.all.mPM[i].value;
							 }else{
							 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
							 }
						}
					}else{
							document.all.mPM.value = document.all.mPM.value.replace(/\,/g,"");
							document.frm.mP.value =	document.all.mPM.value;
					}
				}

				document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
				document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
				document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");
  			document.frm.hidM.value =sMode;
  			document.frm.submit();
  		}
  }

</script>
</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4">
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">
<tr>
	<td>
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<form name="frm" method="post" action="procpayrequest.asp">
		<input type="hidden" name="hidM" value="I">
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
		<input type="hidden" name="iprIdx" value="<%=ipayrequestidx%>">
		<input type="hidden" name="hidPRS" value="0">
		<input type="hidden" name="iAP" value="<%IF pmistate(1) = 5 THEN%>2<%ELSE%>1<%END IF%>"><!-- 반려일때만 결재위치가 2인 경우가 있다.-->
		<input type="hidden" name="blnL" value="1">
		<input type="hidden" name="iptt" value="1">
		<input type="hidden" name="hidRU" value="popIndex.asp">
		<input type="hidden" name="hidcustcd" value="<%=scust_cd%>">
		<input type="hidden" name="hidPDidx" value="<%=ipayDocIdx%>">
		<input type="hidden" name="iRM" value="<%=iRectMenu%>">
		<input type="hidden" name="sbizno" value="<%=iBizNo%>">
		<%IF ipayrequestidx>0 THEN%>
		<Tr>
			<td align="right"  style="border-bottom:1px dashed #cccccc;"><input type="button" value="프린트" class="button"  onClick="jsPopMPPrint(<%=ireportidx%>,<%=ipayrequestidx%>);"></td>
		</tr>
		<%END IF%>
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b>결제요청서(<%=sarap_nm%>)</b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">결제요청서Idx</td>
					<td bgcolor="#FFFFFF" ><%IF ipayrequestidx > 0 THEN%><%=ipayrequestidx%><%END IF%></td>
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="300">
						<!------결재자 리스트------------------------------------------------------------>
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0">
						<tr>
							<td  valign="top" width="150">
								<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>최종승인자</td> </tr>
									<tr align="Center"><td><%=pmstate(0)%></td> </tr>
									<tr align="Center"> <td><input type="hidden" name="hidAI1" value="<%=pmuserid(0)%>">   <%=pmusername(0)%>&nbsp;<%=pmjobname(0)%></td> </tr>
									<tr align="Center"><td><%=pmdate(0)%></td></tr>
								</table>
							</td>
							<td  valign="top">
								<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>재무회계담당</td> </tr>
									<tr align="Center"><td><font color="gray"><%=pmstate(1)%></font></td> </tr>
									<tr align="Center"> <td><input type="hidden" name="hidAI2" value="<%=pmuserid(1)%>"><%=pmusername(1)%>&nbsp;<%=pmjobname(1)%></td> </tr>
									<tr align="Center"><td><%=pmdate(1)%></td></tr>
								</table>
							</td>
						</tr>
						</table>
						<!------//결재자 리스트------------------------------------------------------------>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">팀/부서</td>
					<td bgcolor="#FFFFFF"><%=sdepartmentname%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성자</td>
					<td bgcolor="#FFFFFF"><%= session("ssBctCname")%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성일</td>
					<td bgcolor="#FFFFFF"><%IF dregdate <> "" THEN%><%=formatdate(dregdate,"0000-00-00")%><%ELSE%><%=date()%><%END IF%></td>
				</tr>
					<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(ipayrequeststate)%></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr><!--관련품의서-->
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="80" rowspan="2"  align="center" >관련품의서</td>
					<td>품의서 IDX</td>
					<td>품의서명</td>
					<td>품의금액(원)</td>
					<td>비고</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
					<td><%=ireportidx%></td>
					<td><%=sreportname%></td>
					<td><%=formatnumber(mreportprice,0)%></td>
					<td><a href="javascript:jsPopView('/admin/approval/eapp/confirmeapp.asp?iridx=<%=ireportidx%>');">상세보기>></a></td>
				</tr>
				</table>
			</td>
		</tr><!--//관련품의서-->
		<!--기결제내용-->
		<% Dim totPrice
		totPrice = 0
		IF isArray(arrProc) THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr align="center">
					<td  bgcolor="<%= adminColor("tabletop") %>" width="80" rowspan="<%=UBound(arrProc,2)+2%>">기결제내용</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">결제요청서 IDX</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">결제(입금)일</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">결제금액(원)</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">결제상태</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">비고</td>
				</tr>
				<%For intP = 0 To UBound(arrProc,2)
					totPrice = totPrice + arrProc(2,intP)
				%>
				<tr align="center">
					<td bgcolor="#FFFFFF"><%=arrProc(0,intP)%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(3,intP) <> "" THEN%><%=formatdate(arrProc(3,intP),"0000-00-00")%><%END IF%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(2,intP) <> "" THEN%><%=formatnumber(arrProc(2,intP),0)%><%END IF%></td>
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(arrProc(4,intP))%></td>
					<td bgcolor="#FFFFFF"><a href="javascript:jsPopView('regpayrequest.asp?iridx=<%=ireportidx%>&ipridx=<%=arrProc(0,intP)%>');">상세보기 >></a></td>
				</tr>
				<%Next%>
				</table>
			</td>
		</tr>
		<%END IF%>	<!--//기결제내용-->
		<input type="hidden" name="hidTP" value="<%=totPrice%>">
		<%IF mpayrequestprice = "" or isNull(mpayrequestprice) THEN 	 mpayrequestprice = mreportprice-totPrice  %>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="80" rowspan="6">결제요청내용</td>
					<td>결제요청일</td>
					<td>결제요청금액(원)</td>
					<td>결제방법</td>
					<td width="250" >비고</td>
				</tr>
				<tr align="center"  bgcolor="#FFFFFF">
					<td><input type="text" name="dprd" value="<%IF dpayrequestdate <> "" THEN%><%=formatdate(dpayrequestdate,"0000-00-00")%><%END IF%>" size="10" <%IF blnMod = 0 THEN%>style="border:0" readonly<%END IF%>><%IF blnMod = 1 THEN%><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('dprd');"  style="cursor:hand;"><%END IF%></td>
					<td><input type="text" name="mprp" value="<%IF mpayrequestprice <> "" THEN%><%=formatnumber(mpayrequestprice,0)%><%END IF%>" size="15" style="text-align:right;<%IF blnMod = 0 THEN%>border:0;" readonly<%ELSE%>"<%END IF%> onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" ></td>
					<td>
						<%
						IF blnmod= 0 THEN
							%><%=fnGetPayType(ipaytype)%>
					 <%ELSE%>
							<select name="selPT" onChange="jsChFC();" class="select">
								<%sboptPayType ipaytype%>
							</select>
						<%END IF%>
					</td>
					<td>
							<span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;">
							외화금액: <%IF blnMod=0 THEN%><%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%><%ELSE%><%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"><%END IF%>
							</span>
					</td>
				</tr>
				<tr>
					<td   align="center"  bgcolor="<%= adminColor("tabletop") %>">자금용도</td>
					<td colspan="3" bgcolor="#FFFFFF"><input type="text" class="text" id="sprt" name="sprt" value="<%=spayrequesttitle%>" size="60" <%IF blnMod = 0 THEN%>style="border:0" readonly<%END IF%>></td>
				</tr>
				<tr  align="center"   bgcolor="<%= adminColor("tabletop") %>">
					<td>거래처</td>
					<td>은행명</td>
					<td>예금주명</td>
					<td>계좌번호</td>
				</tr>
				<tr align="center"  bgcolor="#FFFFFF">
					<td><input type="text" name="scustnm" value="<%=scustnm%>" size="20"  readonly class="text_ro"> <%IF blnMod=1 THEN%><input type="button" class="button" value="선택" onClick="jsGetCust('<%=scust_cd%>')"><%END IF%></td>
					<td><input type="text" name="selIB" value="<%=iinBank%>" readonly class="text_ro"></td>
					<td><input type="text" name="sah" value="<%IF saccountholder <> "" THEN%><%=saccountholder%><%END IF%>" size="15"   readonly class="text_ro"></td>
					<td><input type="text" name="san" value="<%IF saccountno <> "" THEN%><%=saccountno%><%END IF%>" size="20" readonly class="text_ro"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td rowspan="8" width="80"  align="center" style="padding:5px;">증빙서류</td>
					<td colspan="2" bgcolor="#FFFFFF"  style="padding:5px;">
						 	<!-- #include virtual="/admin/approval/eapp/incDocInfo.asp" -->
						</td>
				</tr>
				<tr>
					<td width="90" bgcolor="<%= adminColor("tabletop") %>" align="center"  style="padding:5px;">서류종류</td>
					<td bgcolor="#FFFFFF"  style="padding:5px;" width="620">
						<%IF blnMod = 1 THEN%>
						<input type="radio" name="rdoDK" value="1" <%IF ipaydockind ="1" THEN%>checked<%END IF%> onClick="jsSetDocDis(1);">세금계산서-전자
						 <input type="radio" name="rdoDK" value="2" <%IF ipaydockind ="2"  THEN%>checked<%END IF%> onClick="jsSetDocDis(2);">세금계산서-수기
						<!--<input type="radio" name="rdoDK" value="3" <%IF ipaydockind ="3"  THEN%>checked<%END IF%> onClick="jsSetDocDis(3);">현금영수증-소득공제용
						<input type="radio" name="rdoDK" value="4" <%IF ipaydockind ="4"  THEN%>checked<%END IF%> onClick="jsSetDocDis(4);">현금영수증-지출증빙용 -->
						<!--<input type="radio" name="rdoDK" value="5" <%IF ipaydockind ="5"  THEN%>checked<%END IF%> onClick="jsSetDocDis(5);">기타영수증 [비타민등]//-->
						<input type="radio" name="rdoDK" value="9" <%IF ipaydockind ="9"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">서류없음 [전도금/운영비/비타민등]
						<br>
					 	<input type="radio" name="rdoDK" value="8" <%IF ipaydockind ="8"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">계산서 차후 수취(선급금처리)
					 	[선결제 후 계산서를 나중에 받을경우]
						<%ELSE
						Dim strDoc
							IF ipaydockind ="1" THEN
								strDoc = "세금계산서-전자"
							ELSEIF ipaydockind ="2" THEN
								strDoc = "세금계산서-수기"
							ELSEIF ipaydockind ="3" THEN
								strDoc = "현금영수증-소득공제용"
							ELSEIF ipaydockind ="4" THEN
								strDoc = "현금영수증-지출증빙용"
							ELSEIF ipaydockind ="5" THEN
								strDoc = "기타영수증"
							ELSEIF ipaydockind ="8" THEN
								strDoc = "계산서 차후 수취"
							ELSE
								strDoc = "서류없음"
							END IF
						%>
						<%=strDoc%>
						<%END IF%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" colspan="2">
						<div id="dSel1" style="display:<%IF not (ipaydockind = "1" or ipaydockind ="2") OR blnMod =0 THEN%>none<%END IF%>;padding:5px;">
						<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
						<tr>
							<td bgcolor="#FFFFFF" colspan="5">
								<input type="button" name="btnB1" class="button" style="color:blue;" value="세금계산서 검색" onClick="jsGetTax('<%=iBizNo%>','<%=mpayrequestprice%>');">
								<!--재무팀 요청으로 주석처리(2016.03.04)-->
								<!--<span id="spB2" style="display:<%IF ipaydockind <> "1" THEN%>none<%END IF%>;"><input type="button" name="btnB2" class="button" value="XML 등록" onClick="jsNewRegXML();" readonly></span>-->
								<!--//-->
								<span id="spB3" style="display:<%IF ipaydockind <> "2" THEN%>none<%END IF%>;"><input type="button" name="btnB3" class="button" value="종이세금계산서 등록" onClick="jsNewRegHand();" readonly></span></td>
						</tr>
						</table>
						</div>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" colspan="2">
						<div id="dView1" style="display:<%IF not (ipaydockind = "1" or ipaydockind ="2") or setaxkey = "" THEN%>none<%END IF%>;">
						<table border="0" cellpadding="5" cellspacing="0" class="a" width="100%">
						<tr>
							<td  width="90" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 과세구분 </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;">
									<%
								Dim strVat
									IF sVatKind ="0" THEN
										strVat = "과세(부가세 10%) "
									ELSEIF sVatKind ="2" THEN
										strVat = "면세"
									ELSEIF sVatKind ="3" THEN
										strVat = "영세"
									END IF
								%><input type="text" name="sVK" value="<%=strVat%>" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%> ><input type="hidden" name="rdoVK" value="<%=sVatKind%>">
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 발행일 </td>
							<td bgcolor="#FFFFFF"  colspan="3" style="border-bottom:1px solid #BABABA;"><input type="text" name="dID" value="<%=dissuedate%>" size="10" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%>  <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><!--%IF blnMod=1 THEN%--><!--img src="/images/calicon.gif" id="imgCal" align="absmiddle" border="0" onClick="jsPopCal('dID');"  style="cursor:hand;"  disabled --><!--%END IF%--></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 품목 </td>
							<td colspan="5" bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"> <input type="text" id="sINm" name="sINm" value="<%=sItemName%>" size="40"  <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%> <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>> </td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 총금액 </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mTP" value="<%=mTotPrice%>" size="10"   class="text_ro" style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mTotPrice,0)%><%END IF%>원</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 공급가 </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mSP" value="<%=mSupplyPrice%>" size="10" class="text_ro"  style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mSupplyPrice,0)%><%END IF%>원</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 부가세 </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mVP" value="<%=mVatPrice%>" size="10" class="text_ro"  style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mVatPrice,0)%><%END IF%>원</td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> 국세청승인번호 </td>
								<td bgcolor="#FFFFFF" style="border-right:1px solid #BABABA;"><input type="text" name="sEK" value="<%=setaxkey%>" size="30" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%>  <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>></td>
								<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> 비고 </td>
								<td bgcolor="#FFFFFF"  colspan="3"><input type="text" name="sDB" value="<%=sDocBigo%>" size="40" <%IF blnMod=0 THEN%>style="border:0" <%END IF%>   class="text" <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>> </td>
						</tr>
						</table>
			 			</div>
			 		</td>
			 	</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2" width="80">첨부서류</td>
					<td>첨부파일</td>
					<td>관련링크</td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<td align="center" valign="top">
						<%IF blnMod = 1 THEN %><input type="button" value="파일첨부" class="button" onClick="jsAttachFile('');"><%END IF%>
						<div id="dFile">
						<%
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2)
							IF arrFile(2,intF) = 0 THEN Exit For
								arrF = split(arrFile(1,intF),"/")
								arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0)
						%>
						<div id="dF<%=sFName%>"> <a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a><%IF blnMod = 1 THEN %>&nbsp;<input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"><%END IF%></a>
						<input type="hidden"  name="sFileP[]"   value="<%=arrFile(1,intF)%>"></div>
						<%Next
						END IF
						%>
						</div>
						<input type="hidden" name="sFile" value="">
					</td>
					<td>
						<% iCount = 0
						IF isArray(arrFile) THEN
						For intF2 = intF To UBound(arrFile,2)%>
						<input type="text" name="sL" size="60" maxlength="120" value="<%=arrFile(1,intF2)%>" <%IF blnMod = 0 THEN%>style="border:0;cursor:hand;" readonly onClick="jsFileLink('<%=arrFile(1,intF2)%>');"<%END IF%> ><br>
						<% iCount = iCount + 1
						Next
						END IF
						For intF3= iCount To 4
						%>
						<input type="text" name="sL" size="60" maxlength="120" <%IF blnMod = 0 THEN%>style="border:0" readonly<%END IF%>><br>
						<%Next

						%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="3" width="80">계정과목</td>
					<td>수지항목</td>
					<td>연결계정과목</td>
				</tr>
				<tr bgcolor="#FFFFFF"  align="center">
					<td><input type="text" name="sANM" value="<%=CHKIIF(isNULL(sarap_nm),"","["&iarap_cd&"]"&sarap_nm)%>" style="border:0;width=200px" readonly ><input type="hidden" name="iaidx" value="<%=iarap_cd%>" class="text"></td>
					<td><input type="text" name="sACCNM" value="<%=CHKIIF(isNULL(sacc_nm),"","["&sacc_use_cd&"]"&sacc_nm)%>" style="border:0;width=200px" readonly><input type="hidden" name="sACC" value="<%=sacc_cd%>"  class="text"></td>
				</tr>
				<%IF blnMod = 1 THEN%>
				<tr bgcolor="#FFFFFF">
					<td colspan="2"><input type="button" class="button" value="수지항목 수정" onClick="jsGetARAP();"></td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center" >
					<td width="80" rowspan="2" style="padding:5px">부서별<br>자금구분</td>
					<td width="321" style="padding:5px" > 부서</td>
					<td width="201" style="padding:5px"> 금액</td>
					<td width="160" style="padding:5px"> %</td>
				</tr>
				<tr>
					<td colspan="3" bgcolor="#FFFFFF" valign="top">
					<div id="divPM">
					<%dim arrPV, arrPT
					IF isArray(arrPart) THEN %>
						<table border="0" cellpadding="3" cellspacing="0" class="a" width="737">
					<%	For intPart = 0 To UBound(arrPart,2)
						IF intPart > 0 THEN
								arrPV = arrPV&","
								arrPT =arrPT&","
							END IF
							arrPV = arrPV&arrPart(1,intPart)
							arrPT = arrPT&arrPart(3,intPart)
					%>
					<tr>
						<td width="150" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(4,intPart)%></td>
						<td width="150" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(3,intPart)%> </td>
						<td width="201" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><input type="text" name="mPM" id="mPM" size="20" value="<%=formatnumber(arrPart(2,intPart),0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="auto_amount(this.form,this);jsSetMoney('m',<%=intPart%>,2);"  onKeypress="num_check()"  > 원</td>
						<td width="160" style="border-bottom:1px solid #BABABA;" align="center"><input type="text" name="iPM"  size="4" value="<%IF mpayrequestprice <> 0 AND arrPart(2,intPart)<> 0 THEN%><%=formatnumber((arrPart(2,intPart)/mpayrequestprice)*100,1)%><%END IF%>"  style="text-align:right;<%IF blnMod = 0  THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('i',<%=intPart%>,2);">%</td>
					</tr>
					<%	Next %>
					</table>
					<%END IF%>
					</div><br>
					<input type="hidden" name="iP" id="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" id="sP" value="<%=arrPT%>">
					<input type="hidden" name="mP" id="mP" value="">
					<%IF blnMod = 1 THEN%><br>&nbsp;	<input type="button" value="부서 등록/수정" onClick="jsSetPartMoney(2,'<%=sacc_use_cd%>','<%=sACC_GRP_CD%>');" class="button"><br><BR><%END IF%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%IF isArray(arrReturn) THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">반려리스트</td>
					<td bgcolor="#FFFFFF">
						<%For intRA = 0 To UBound(arrReturn,2)%>
						<%=arrReturn(0,intRA)%>차 검토 반려&nbsp;<%=arrReturn(1,intRA)%>&nbsp;<%=formatdate(arrReturn(2,intRA),"0000-00-00")%><br>
						<%Next%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%END IF%>
		<%IF blnMod = 1 THEN%>
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td align="left"><input type="button" value="내용저장" class="button" onClick="jsPayReqUpdate('<%=sMode%>');">
						<input type="button" value="삭제" class="button" onClick="jsPayReqUpdate('D');" style="color:red;">
						</td>
						<td align="right"><input id="btnSm" type="button" value="결제요청" class="button" onClick="jsPayEappSubmit('<%=sMode%>','<%=mreportprice%>','<%=iarap_cd%>');"></td>
					</tR>
				</table>
			</td>
		</tr>
		<%END IF%>
		<%IF ipayrequeststate >=7 THEN '승인 또 결제확인된 상태일때만 내용보여준다%>
			<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="3" width="80">경영지원팀<br>관리항목</td>
					<td>결제예정일</td>
					<td>결제(입금)일</td>
					<td>해당년월(손익)</td>
					<td>서류제출여부</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
					<td><%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%></td>
					<td><%IF dpayrealdate <> "" THEN%><%=formatdate(dpayrealdate,"0000-00-00")%><%END IF%></td>
					<td><%IF syyyymm <> "" THEN%><%=year(syyyymm)%> 년
						<%=month(syyyymm)%> 월<%END IF%>
						</td>
					<td> <%IF blnTakeDoc THEN%>Y<%ELSE%>N<%END IF%>
						</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td colspan="5">*COMMENT<br>
					<%=pcomment%><Br>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
		<%END IF%>
		</form>
		</table>
	</td>
</tr>
<%IF ipayRequestIdx > 0 THEN%>
<tr>
	<td style="padding-top:20px;">
		<!-- #include virtual="/admin/approval/eapp/incComment.asp" -->
	</td>
</tr>
<%END IF%>
<Tr>
	<td height="50">&nbsp;</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
