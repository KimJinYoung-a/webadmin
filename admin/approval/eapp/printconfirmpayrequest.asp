<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 결제요청서 결재처리 
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
<%
Dim clsPay,clsMem,clseapp, clscomm,clsPM
Dim ireportidx,ipayrequestIdx,iarap_Cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate ,iauthstate
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_use_cd,sacc_nm,sedmsName,sedmscode 
Dim spartname ,ilastApprovalid,sscmLink, iauthposition
Dim chkPayRequest , spayrequestTitle
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate,pcomment ,susername
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrProc,arrPM, arrPart
Dim intA, intC, intF, intR, intRA, intP , iAuthCount, intPart
Dim blnMod 
Dim mSumPrice, mSumRealPrice   
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2), pmistate(2)	,pmstatecd(2)					
Dim scust_cd, scustnm
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	 
Dim igbn, chkID
Dim ipaytype, sCurrencyType, sCurrencyPrice		

ireportidx 		=  requestCheckvar(Request("iridx"),10)
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
iauthstate		= requestCheckvar(Request("ias"),10)
igbn					= requestCheckvar(Request("igbn"),1)
blnMod = 0 
chkID = 0
'결재 기본 폼 정보 가져오기
set clsPay = new CPayRequest
	clsPay.Freportidx = ireportidx  
	clsPay.FpayrequestIdx = ipayrequestIdx  
	
	clsPay.fnGetPayRequestReceiveData 
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
	sedmsName        = clsPay.FedmsName      
	sedmscode        = clsPay.Fedmscode 
	ilastApprovalid	 = clsPay.FlastApprovalid     

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
  pcomment		= clsPay.FpayComment
  susername		 = clsPay.Fusername
  spartname		 = clsPay.Fpartname
  spayrequestTitle	= clsPay.FpayRequestTitle 							   
 	scust_cd						= clsPay.Fcust_cd
 	scustnm						=clsPay.Fcust_nm 
 	ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	
	'//기결제리스트
	arrProc			= clsPay.fnGetProcPayRequestList	
 
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

'결재라인, 코멘트, 파일 리스트 가져오기
set clseapp = new CEApproval	
	clseapp.Freportidx 		= ireportidx  
	clseapp.FpayrequestIdx = ipayrequestIdx 	
 	 
 	arrAuth			= clseapp.fnGetAuthLineList  '결재정보
	arrComm			= clseapp.fnGetCommentList	'코멘트
	arrFile			= clseapp.fnGetAttachFileList  '첨부파일 
	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing  
 
 
 '-------------------------------------------
 '-- 결재라인 변수값 지정
 '1.결재라인 DB 저장값 변수에 저장
 '2.상태값 (ipayrequeststate = 1)  : 모두 승인대기이거나 최종승인자만 승인완료한 상태 
 ' 변수 값 변경가능 (담당자 db에서 기본 사용자 또는 현재 로그인 사용자와 동일할때는 로그인 사용자로 값 지정)
 ' 최종승인자가 승인완료 상태일때는 변경불가능 ->  arrPM(2,intP) = 1 : 최종승인자이고, pmstatecd(0) = 0 : 승인대기 상태일때 변경가능
 '------------------------------------------- 

IF isArray(arrAuth) THEN '1.결재라인 값 변수에 저장.  
		For intA = 0 To UBound(arrAuth,2)
			pmuserid(intA)  = arrAuth(2,intA) 
			pmusername(intA)= arrAuth(7,intA)
			pmjobname(intA) = arrAuth(10,intA)
			pmstate(intA)	= fnGetPayAuthState(arrAuth(3,intA), intA+1)
			pmstatecd(intA)= arrAuth(3,intA)
			pmdate(intA)	= arrAuth(6,intA)  
		Next 
END IF	 
  
IF ipayrequeststate = 1  THEN  '2.재무회계팀 결제요청서 처리자정보 가져와서 변수값 변경 
Set clsPM	= new CPayManager
	clsPM.Fuserid = session("ssBctId")
	arrPM	= clsPM.fnGetPayManager 
Set clsPM 	= nothing  
 IF isArray(arrPM) THEN 
			For intP = 0 To UBound(arrPM,2)
			IF pmstatecd(intP) = 0 THEN 
			pmuserid(intP)  = trim(arrPM(1,intP))
			pmusername(intP)= arrPM(3,intP)
			pmjobname(intP) = arrPM(6,intP) 
			pmstate(intP)	= fnGetPayAuthState(0,arrPM(2,intP))
			pmdate(intP)	= "&nbsp;"
			END IF
			Next  
	END IF 
END IF 
 
'--------------------------------------------------	  
 IF iauthstate = 0 THEN 
 	iauthposition = 1
 ELSE 
 	iauthposition = 2
 END IF
 
'문서 수정가능여부   
 IF   ipayrequeststate = 1 OR  ipayrequeststate = 7  THEN  
  IF pmuserid(iauthstate)   = session("ssBctId") THEN
 	blnMod = 1
 	clsPay.FadminId = session("ssBctId")
 	clsPay.FauthPosition = iauthposition
	clsPay.fnCheckPayRequestView  '//결재내용 확인여부 체크  
	END IF
 END IF
 set clsPay = nothing 
%> 
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script type="text/javascript" src="eapp.js"></script>
<link rel="stylesheet" href="eapp.css" type="text/css"> 
</head> 
<body topmargin="0" leftmargin="0"> 
<table width="840" height="100%" cellpadding="0" cellspacing="0"  align="center" border="0">   
<tr>
	<td width="160" style="padding-right:10px;" ><!-- 왼쪽 메뉴-->
		<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" >
			<Tr>
				<td valign="top">
					<table width="100%" border="0" cellpadding="0" cellspacing="0" >
					<tr><td><img src="/images/top_logo.gif"></td></tr>
					<tr><td class="btdsmall">문서코드:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=sedmscode%></font></td></tr>
					<tr><td class="btdsmall">팀/부서:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=spartname%></td></tr>
					<tr><td class="btdsmall">작성자:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=susername%></td></tr>
					<tr><td class="btdsmall">작성일:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=formatdate(dregdate,"0000-00-00")%></td></tr> 
					<%IF isArray(arrAuth) THEN  %>
					<tr>
						<td style="padding-bottom:5px;">
								<table border=1 cellspacing=0 cellpadding=3   width="100%"> 
							<tr>
								<td  class="btdsmall">최종승인자</td>
							</tr>
							<tr>
								<td><%=pmusername(0)%>&nbsp;<%=pmjobname(0)%></td>
							</tr>
							<tr>
								<td><%=pmstate(0)%></td>
							</tr>
							<tr>
								<td><%=pmdate(0)%></td>
							</tr>
							
						</table>
						</td>
					</tr>
						<tr>
						<td style="padding-bottom:10px;">
								<table border=1 cellspacing=0 cellpadding=3 class="a" width="100%"> 
							<tr>
								<td  class="btdsmall">재무회계담당</td>
							</tr>
							<tr>
								<td><%=pmusername(1)%>&nbsp;<%=pmjobname(1)%></td>
							</tr>
							<tr>
								<td><%=pmstate(1)%></td>
							</tr>
							<tr>
								<td><%=pmdate(1)%></td>
							</tr>
							
						</table>
						</td>
					</tr>
					<% 
								END IF
							%>
				</table>
			</td>
		</tr>
			<tr>
				<td valign="bottom">
					<table border=0 cellspacing=0 cellpadding=0   width="100%">
					<tr>
						<td style="padding-bottom:20px;">(주)텐바이텐</td>
					</tr>
					<tr>
						<td  style="padding-bottom:5px;"> 03082<br>
					 	서울시 종로구 대학로 57<br>
					 	홍익대학교 대학로캠퍼스 교육동 14층
					</td>
					<tr> 
							<td style="padding-bottom:5px;">
					 	TEL: 02-554-2033<br>
					 	FAX: 02-2179-9245
					</td>
				</tr>
				<tr>
					<td style="padding-bottom:5px;">
					 	E-mail: <br>
					 	customer@10x10.co.kr<br>
					 	Website:<br>
					 	www.10x10.co.kr<br> 
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
	</td><!-- /왼쪽 메뉴-->
	<td  valign="top"><!-- 상세내용-->
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr> 
				<td height="50"  >
					<table border="0" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td valign="bottom" class="btd20"> idx <%=ipayrequestidx%>  </td>  
							<td align="right" valign="top"><img src="/images/10x10-logo400px.jpg"></td>
					</tr> 
					<tr>
						<td colspan="2" class="btd20" valign="top">결제요청서(<%=sarap_nm%>)</td>
					</table>
				</td>
			</tr>
			<%IF ireportidx > 0 THEN%>
			<tr>
				<td><br>[정보]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
						<tr  align="center">
							<td class="td01">idx</td>
							<td class="td01">품의서명</td>
							<td class="td01">품의금액</td>
							<td class="td01">scm문서번호</td>
						</tr>
						<tr  align="center">
							<td><%=ireportidx%></td>
							<td><%=sreportname%></td>
							<td><%=formatnumber(mreportprice,0)%></td>
							<td><%=iscmlinkno%></td>
						</tr>
					</table>
				</td>
			</tr> 
				<% Dim totPrice 
		totPrice = 0
		IF isArray(arrProc) THEN%>
		<tr>
			<td style="padding-top:15px;"><br>[기결제내용]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01">
				<tr align="center"> 
					<td class="td01">결제요청서 IDX</td>
					<td class="td01">결제(입금)일</td>
					<td class="td01">결제금액</td> 
					<td class="td01">결제상태</td>  
					<td class="td01">자금용도</td>  
				</tr>
				<%For intP = 0 To UBound(arrProc,2)
					totPrice = totPrice + arrProc(2,intP)
				%>
				<tr align="center">	
					<td><%=arrProc(0,intP)%></td>
					<td><%IF arrProc(3,intP) <> "" THEN%><%=formatdate(arrProc(3,intP),"0000-00-00")%><%END IF%></td>
					<td><%IF arrProc(2,intP) <> "" THEN%><%=formatnumber(arrProc(2,intP),0)%><%END IF%></td> 
					<td><%=fnGetPayRequestState(arrProc(4,intP))%></td> 
					<td><%=arrProc(6,intP)%></td>
				</tr>
				<%Next%> 
				</table>
			</td>
		</tr>
		<%END IF%>	
			<%END IF%>
		<input type="hidden" name="hidTP" value="<%=totPrice%>"> 
		<tr>
				<td style="padding-top:15px;"><br>[내용]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
					<tr  align="center" >
						<td class="td01">결제요청일</td>
						<td class="td01">결제요청금액</td>
						<td class="td01">결제방법</td>
						<td class="td01">비고</td> 
					</tr> 
					<tr align="center" >	
					<td class="td02"><%IF dpayrequestdate <> "" THEN%><%=formatdate(dpayrequestdate,"0000-00-00")%><%END IF%></td>
					<td class="td02"><%=formatnumber(mpayrequestprice,0)%></td>
				 	<td class="td02"><%=fnGetPayType(ipaytype)%>&nbsp;</td>
				 	<td class="td02"><span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							외화금액: <%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%> 
							</span>&nbsp;</td>
				</tr> 
				<tr  align="center">	
					<td class="td01" colspan="4">자금용도</td>
				</tr>
				<tr  align="center">	
					<td class="td02" colspan="4"><%=spayrequesttitle%>&nbsp;</td>
				</tr>
				<tr  align="center">	
					<td class="td01">거래처</td>
					<td class="td01">은행명</td>
					<td class="td01">계좌번호</td>
					<td class="td01">예금주명</td>
				</tr> 
				<tr align="center">	
					<td><%=scustnm%></td>
					<td><%=iinBank%></td> 
					<td><%=saccountno%></td>
					<td><%IF saccountholder <> "" THEN%><%=saccountholder%><%END IF%></td> 
				</tr> 
					</table>
				</td>
		</tr>
		<tr>
			<td style="padding-top:15px;"><br>[증빙서류]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0" >
				<tr align="center">
					<td class="td01">서류종류</td>
					<td class="td01">품목 </td> 
					<td class="td01">국세청승인번호</td> 
					<td class="td01">발행일	</td>
				</tr>	
				<tr align="center">
					<td class="td02">
						<%	Dim strDoc
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
							%><%=strDoc%>
					</td>
					<td class="td02"><%=sItemName%>&nbsp;</td> 
					<td class="td02"><%=setaxkey%>&nbsp;</td>
					<td class="td02"><%=dissuedate%>&nbsp;</td> 
				</tr> 
				<tr align="center">		
					<td class="td01">과세구분 </td>
					<td class="td01">총금액 </td>
					<td class="td01">공급가</td>
					<td class="td01">부가세 </td> 
				</tr>
			 <tr align="center">
			 	<td class="td02"> 
						<%
							Dim strVat
							IF sVatKind ="0" THEN 
								strVat = "과세(부가세 10%) "
							ELSEIF sVatKind ="2" THEN 
								strVat = "면세"
							ELSEIF sVatKind ="3" THEN 
								strVat = "영세" 
							END IF
						%> <%=strVat%>
						</td>
					<td class="td02"><%=formatnumber(mTotPrice,0)%> 원</td>
					<td class="td02"><%=formatnumber(mSupplyPrice,0)%> 원</td>
					<td class="td02"><%=formatnumber(mVatPrice,0)%> 원</td> 
				</tr>		 
				</table>
			</td>
		</tr>
		<tr>
			<td style="padding-top:15px;"><br>[첨부서류]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0" >
				<tr   align="center"> 
					<td class="td01" width="50%">첨부파일</td>
					<td class="td01" width="50%">관련링크</td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<%IF isArray(arrFile) THEN%>
					<td align="center" valign="top" > 
						<div id="dFile"> 
						<%  Dim arrFName,arrF, sFName, intF2,intF3, iCount
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2)
							IF arrFile(2,intF) = 0 THEN Exit For
								arrF = split(arrFile(1,intF),"/")  
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0) 
						%>
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a> 
						<input type="hidden" name="sFile" value="<%=arrFile(1,intF)%>"></div>
						<%Next
						END IF
						%> 
						</div>
					</td>
					<td> 
						<% iCount = 0
						IF isArray(arrFile) THEN
						For intF2 = intF To UBound(arrFile,2)%>
						 <%=arrFile(1,intF2)%> <br>
						<% iCount = iCount + 1
						Next
						END IF  
						%> 
					</td>
					<%ELSE%>
					<td colspan="2" align="center" >첨부 파일이 없습니다.</td>
					<%END IF%>
				</tr>
				</table>
			</td>
		</tr>
		<%IF iarap_cd <> "0" THEN%>
		<tr>
					<td style="padding-top:15px;"><br>[계정과목]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr align="center"> 
							<td class="td01">수지항목</td>
							<td class="td01">연결계정과목</td> 
						</tr>
						<tr   align="center"> 
							<td>[<%=iarap_cd%>] <%=sarap_nm%></td>
							<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
						</tr>	
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:15px;"><br>[부서별 자금구분]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr   align="center" > 
							<td class="td01"> 부서</td>
							<td class="td01">금액</td>
							<td class="td01">반영율</td>
						</tr> 
							<%dim arrPV, arrPT
							IF isArray(arrPart) THEN  
							 	For intP = 0 To UBound(arrPart,2) 	
									IF intP > 0 THEN
										arrPV = arrPV&"," 
										arrPT =arrPT&"," 
									END IF	
									arrPV = arrPV&arrPart(1,intP)
									arrPT = arrPT&arrPart(3,intP)
							%>   
							<tr>
								<td  class="td02"  align="center"> 
									<%=arrPart(4,intP)%> > <%=arrPart(3,intP)%>
								</td>
								<td  class="td02" align="center"><%=formatnumber(arrPart(2,intP),0)%> 원</td>
								<td  class="td02" align="center"><%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=(arrPart(2,intP)/mreportprice)*100%><%END IF%> %</td>
							</tr> 
							<%	Next %> 
							<%END IF%>  
						</table>
					</td>
				</tr>
				<%END IF%>
				<%IF isArray(arrReturn) THEN%>
				<tr>
					<td style="padding-top:15px;"><br>[반려리스트]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr>
							<td align="center" class="td01">반려리스트</td> 
						</tr>
						<tr>	
							<td>
								<%For intRA = 0 To UBound(arrReturn,2)%>
								<%=arrReturn(0,intRA)%>차 검토 반려&nbsp;<%=arrReturn(1,intRA)%>&nbsp;<%=formatdate(arrReturn(2,intRA),"0000-00-00")%><br>
								<%Next%>
							</td>
						</tr>
						</table>
					</td>		
				</tr>
				<%END IF%>	  
		<tr>
			<td style="padding-top:15px;"><br>[경영지원팀 관리항목]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0" >
				<tr align="center"> 
					<td class="td01">결제예정일</td> 
					<td class="td01">결제(입금)일</td>
					<td class="td01">해당년월(손익)</td>
					<td class="td01">서류제출여부</td>
				</tr>
				<tr  align="center">
					<td> <%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%></td> 
					<td><%IF dpayrealdate <> "" THEN%><%=formatdate(dpayrealdate,"0000-00-00")%><%END IF%></td>
					<td><%Dim intY, intM%>
					  <%=intY%> 년 <%=intM%> 월
						</td>
					<td> 
						 <%IF blnTakeDoc THEN%>Y<%ELSE%>N<%END IF%> 
						</td>
				</tr> 
				<tr >
					<td colspan="5"  style="border-top:1px solid #bbbbbb;"><font color="#868080">COMMENT</font> <br> 
					<%=pcomment%><br> 
					</td>
				</tr>	
				</table>
			</td>
		</tr>  
		</table><br> 
	</td>
</tr> 
</table> 
<script language="javascript">
<!--
  	document.body.onload=function(){window.print();} 
//-->
</script> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
