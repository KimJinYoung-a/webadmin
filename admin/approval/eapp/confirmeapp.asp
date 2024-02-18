<%@ language="VBScript" %>
<% option explicit %>
 
<%
'###########################################################
' Description : 받은 전자결재 결재처리
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
'변수선언
Dim clseapp,clsMem
Dim ireportidx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate,sreferid
Dim sadminid,dregdate,sarap_nm,iacc_cd,sacc_nm,sacc_use_cd,sedmsName,sedmscode
Dim spartname ,ilastApprovalid,sjob_name,sscmLink,spart_name,susername, ipart_sn
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn, arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim blnAdd
Dim sReferName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast, iNextAuthState,blnMod,iRectAuthState, iRectPartSn
Dim sscmsubmitlink
Dim sRectAuthName, sNextAuthName
Dim iRectMenu ,ipayrequestidx,blnpayEapp
Dim sCurrencyPrice,ipaytype,sCurrencyType,sACC_GRP_CD
Dim AuthID_L,AuthState_L,AuthName_L,AuthJobsn_L,AuthJobName_L,AuthConfirmTime_L,AuthSMS_L
Dim AuthID_F,AuthState_F,AuthName_F,AuthJobsn_F,AuthJobName_F,AuthConfirmTime_F,AuthSMS_F
Dim intNo, arrAID, arrATxt, arrAJSn, sRectAuthType
Dim hidAI_H, hidPS_H, AuthState_H, AuthName_H, AuthJonName_H, AuthConfirmTime_H, AuthSMS_H, hidAJ_H
dim intNo_H, intNo_L '현재 위치

'변수값 받기
iRectMenu = requestCheckvar(Request("iRM"),10)
ireportidx =  requestCheckvar(Request("iridx"),10)
ipayrequestidx = 0

'결재 기본 폼 정보 가져오기
set clseapp = new CEApproval
	clseapp.Freportidx = ireportidx
	clseapp.fnGetEAppData

	iarap_cd				 = clseapp.Farap_cd
	sreportName      = clseapp.FreportName
	mreportPrice     = clseapp.FreportPrice
	iscmlinkno       = clseapp.Fscmlinkno
	sbigo            = clseapp.Fbigo
	tContents  			 = clseapp.Freportcontents
	ireportstate     = clseapp.Freportstate
	sreferid         = clseapp.Freferid
	sadminid         = clseapp.Fadminid
	dregdate         = clseapp.Fregdate
	sarap_nm		     = clseapp.Farap_nm
	iacc_cd					 = clseapp.Facc_cd
	sacc_use_cd			 = clseapp.Facc_use_cd
	sacc_nm		       = clseapp.Facc_nm
	sedmsName        = clseapp.FedmsName
	sedmscode        = clseapp.Fedmscode
	ilastApprovalid	 = clseapp.FlastApprovalid
	sscmLink		 			= clseapp.FscmLink
	sscmsubmitlink	 = clseapp.Fscmsubmitlink
	sjob_name		 			= clseapp.Fjob_name
 	ipart_sn			=clseapp.Fpart_sn
	spart_name		= clseapp.Fpart_name
	susername		= clseapp.Fusername
	blnpayEapp			= clseapp.FispayEapp
	ipaytype				= clseapp.Fpaytype
	sCurrencyType		= clseapp.FCurrencyType
  sCurrencyPrice	= clseapp.FCurrencyPrice
  sACC_GRP_CD			= clseapp.FACC_GRP_CD

	arrAuth			= clseapp.fnGetAuthLineList
	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList
	arrReturn		= clseapp.fnGetAuthLineReturnList
	arrPart			= clseapp.fnGetPartMoneyList

	clseapp.FadminId = session("ssBctId")
	clseapp.fnCheckView    '//결재내용 확인여부 체크
set clseapp = nothing

'refer명 가져오기
set clsMem = new CTenByTenMember
 	if sreferid <> "" then
 	clsMem.Fuserid	= sreferid
	arrRefer		= clsMem.fnGetInIDOutName
	end if
 set clsMem = nothing


'결재리스트-----------------------------------------
blnMod = 0  		'문서 수정 가능여부
blnLast = 0 		'최종결재자 등록여부
 
iRectPosition = 0	'현재결재위치
iNextPosition = 0	'다음결재위치
sNextAuthId = ""	'다음결재자아이디
iNextAuthState = 0	'다음결재상태
iRectAuthState = 0	'현재결재상태
iRectPartSn =ipart_sn '현재부서

'참조 리스트--------------------------------------
  sReferName = ""
 IF isArray(arrRefer) THEN
 	For intR =0 To Ubound(arrRefer,2)
 		IF intR = 0 THEN
 			sreferid	= arrRefer(0,intR)
 			sReferName = arrRefer(1,intR) & arrRefer(5,intR)
 		ELSE
 			sreferid	=sreferid&","& arrRefer(0,intR)
 			sReferName = sReferName &","&arrRefer(1,intR) & arrRefer(5,intR)
 		END IF
	Next
 END IF
 '-------------------------------------------------
  
 
%>
<html>
<head>
	<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->  
<script type="text/javascript" src="eapp.js?t=<%=left(now(),10)%>" charset="euc-kr"/></script>
<script type="text/javascript">
	//내용수정
	function jsEappUpdate(){
			var totPM = 0;
			var mRequestPay = document.all.mRP.value.replace(/\,/g,"");

					if(document.frm.iAIdx.value >0 && document.frm.hidPE.value=="True"){
						if(jsChkBlank(document.all.iP.value) ){
							alert("부서를 등록해주세요");
							return;
						}

						if(jsChkBlank(mRequestPay) ){
							alert("부서를 등록해주세요");
							return;
						}

						if(typeof(document.all.mPM) !="undefined"){
						  	if(typeof(document.all.mPM.length)!="undefined"){
						  		for(i=0;i<document.all.mPM.length;i++){
									totPM = totPM + parseInt(document.all.mPM[i].value.replace(/\,/g,""));
									}
								}else{
									totPM = document.all.mPM.value.replace(/\,/g,"");
								}

							if (parseInt(mRequestPay) != parseInt(totPM)){
								alert("자금구분 금액이 품의금액과 다릅니다. 재설정해주세요");
								return;
							}
						}
					}
		if(confirm("내용을 수정하시겠습니까?")){
				document.all.mRP.value = document.all.mRP.value.replace(/\,/g,"");
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

		document.frm.hidM.value = "CU";
		document.frm.submit();
	}
	}
</script>

</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4" >
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">
<tr>
	<td>
		<form name="frm" method="post" action="proceapp.asp">
		<input type="hidden" name="hidM" value="C"> 
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
		<input type="hidden" name="hidRS" value="1">
		<input type="hidden" name="hidAS" value="0"> 
		<input type="hidden" name="hidAid" value="<%=sadminid%>">
		<input type="hidden" name="iLAID" value="<%=ilastApprovalid%>">
		<input type="hidden" name="sSSL" value="<%=sscmsubmitlink%>"> 
		<input type="hidden" name="hidUN" value="<%=susername%>">
		<input type="hidden" name="iRM" value="<%=iRectMenu%>">
		<input type="hidden" name="iAIdx" value="<%=iarap_cd%>">
		<input type="hidden" name="hidPE" value="<%=blnPayEapp%>"> 
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
		<Tr>
			<td align="right"  style="border-bottom:1px dashed #cccccc;"><input type="button" value="프린트" class="button" onClick="jsPopModPrint(<%=ireportidx%>);"></td>
		</tr>
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b><%=sedmsname%><%IF iarap_cd <> "0" THEN%>_<%=sarap_nm%><%END IF%></b></td>
					<td align="right"><input type='button' class='button' value='전결규정보기' onClick='popDecision();'></td>
					<td align="right" width="100"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="60" align="center">문서코드</td>
					<td bgcolor="#FFFFFF" width="200"><%=sedmscode%></td>
					<td rowspan="6" bgcolor="#FFFFFF" valign="top"  ><!--결재자 리스트-->
						<table  width="100%"   cellpadding="0" cellspacing="1" class="a" border="0">
						<tr align="center">
							<%  
									intNo = 0
									intNo_H = 0
									intNo_L = 0
								IF isArray(arrAuth) THEN
									For intA = 0 to UBound(arrAuth,2)
									 intNo = intNo+1
									'--결재위치 확인
										IF arrAuth(2,intA) = session("ssBctId") THEN
											sRectAuthId = arrAuth(2,intA)		'현재 결재아이디
											sRectAuthName = arrAuth(7,intA)	'현재 결재이름
											iRectPosition= arrAuth(1,intA)	'현재 결재 위치
											iRectAuthState	= arrAuth(3,intA) '현재 결재 상태
											iRectPartSn	= arrAuth(9,intA) '현재부서
											sRectAuthType = arrAuth(4,intA) '현재결재타입(D-결재선, A-합의, L-최종승인)
										 	 
												IF intA+1 <= UBound(arrAuth,2) THEN
													iNextPosition = arrAuth(1,intA+1)
													sNextAuthId	  = arrAuth(2,intA+1)
													sNextAuthName = arrAuth(7,intA+1)
													iNextAuthState = arrAuth(3,intA+1)
													iRectPartSn	= arrAuth(9,intA+1)
												ELSE
													iNextPosition = iRectPosition+1
												END IF
									 
										END IF 
										'--------------------------------------------------
										'문서 수정가능여부
										 IF (iRectAuthState = 0 OR  iRectAuthState = 3  ) AND sRectAuthId = session("ssBctId") THEN
										 	blnMod = 1
										 END IF
  
										'--결재선 내역
										IF arrAuth(4,intA)="A"  THEN ''합의
											hidAI_H         = arrAuth(2,intA)
											hidPS_H         = arrAuth(9,intA)
											AuthState_H     = arrAuth(3,intA)
											AuthName_H      = arrAuth(7,intA)
											hidAJ_H         = arrAuth(8,intA)
											AuthJonName_H   = arrAuth(10,intA)
											AuthConfirmTime_H = arrAuth(6,intA)
											AuthSMS_H       = arrAuth(11,intA)
											intNo_H 				= intNo
										ELSEIF arrAuth(4,intA)="L" THEN   '최종결재자 
								    	 	AuthID_L       	= arrAuth(2,intA) 
											AuthState_L     = arrAuth(3,intA)
											AuthName_L      = arrAuth(7,intA)
											AuthJobsn_L     = arrAuth(8,intA)
											AuthJobName_L   = arrAuth(10,intA)
											AuthConfirmTime_L= arrAuth(6,intA)
											AuthSMS_L       = arrAuth(11,intA)
											intNo_L 				= intNo
										ELSEIF arrAuth(4,intA)="F" THEN		'최종합의자
						    	 			AuthID_F       	= arrAuth(2,intA) 
											AuthState_F     = arrAuth(3,intA)
											AuthName_F      = arrAuth(7,intA)
											AuthJobsn_F     = arrAuth(8,intA)
											AuthJobName_F   = arrAuth(10,intA)
											AuthConfirmTime_F= arrAuth(6,intA)
											AuthSMS_F       = arrAuth(11,intA)
										ELSE
											if arrAID = "" THEN
												arrAID 		= arrAuth(2,intA)
												arrAJSn 	= arrAuth(8,intA)
												arrATxt 	= arrAuth(7,intA)&" "&arrAuth(10,intA)
											else
												arrAID 		= arrAID& ","&arrAuth(2,intA)
												arrAJSn 	= arrAJSn& ","&arrAuth(8,intA)
												arrATxt 	= arrATxt& ","&arrAuth(7,intA)+" "+arrAuth(10,intA)
											end if	 
							 
							%>
									<td valign="top" width="180"  height="100%">
										<div id="dAP<%=intNo%>">
										<table width="100%"  cellpadding="5" cellspacing="0" class="a"  height="100%" border="0" >
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20"><%=intNo%>차 검토</td></tr>
											<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>
											<tr><td align="Center"><%=arrAuth(7,intA)%>&nbsp;<%=arrAuth(10,intA)%></td></tr>
											<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">	<%IF blnmod =  1 and intNo = iNextPosition THEN%><input type='checkbox' value='1' name='chkSms' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr>
										</table>
										</div>
									</td>	
									<%								    	 			
								        END IF
								    Next   
								   	if arrAID = "" THEN    
								   %>
								   	<td valign="top"  width="180"  height="100%">
										<div id="dAP1">
										<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0"  height="100%">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">&nbsp;</td></tr>
										<tr><td height="100" valign="bottom"></td></tr>
										</table>
										</div>
									</td>
								   <%end if%>
								  <td valign="top"  width="180"  height="100%">
							    	<div id="dAP_H">
							    	<table width="100%" cellpadding="5" cellspacing="0" class="a"  height="100%" border=0>
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>"  height="20">합의</td></tr>
											<% if (hidAI_H<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_H)%></td></tr>
											<tr><td align="Center"><%=AuthName_H%>&nbsp;<%=AuthJonName_H%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_H) THEN %><%=formatdate(AuthConfirmTime_H,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center"><%IF blnmod =1 and intNo_H = iNextPosition THEN%>	<input type='checkbox' value='1' name='chkSms_H' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr> 
							   		 <% else %>
											<tr><td align="Center">&nbsp;</td></tr>
											<tr><td align="Center"></td></tr>
											<% end if %>
										</table>
							    	</div>
						    	</td>
						    	 <td valign="top"  width="180"  height="100%">
							    	<div id="dAP0">
							    	<table width="100%" cellpadding="5" cellspacing="0" class="a"  height="100%" border=0>
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">최종<%=chkIIF(AuthID_F<>"","합의","승인")%></td></tr>
										<% if (AuthID_L<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_L)%></td></tr>
											<tr><td align="Center"><%=AuthName_L%>&nbsp;<%=AuthJobName_L%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_L) THEN %><%=formatdate(AuthConfirmTime_L,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center"><%IF blnmod = 1 and intNo_L = iNextPosition  THEN%><input type='checkbox' value='1' name='chkSms' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr> 
										<% elseif (AuthID_F<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_F)%></td></tr>
											<tr><td align="Center"><%=AuthName_F%>&nbsp;<%=AuthJobName_F%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_F) THEN %><%=formatdate(AuthConfirmTime_F,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center"><%IF arrAID="" AND hidAI_H="" and blnmod = 1 THEN%><input type='checkbox' value='1' name='chkSms' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr> 
										<% else %>
											<tr><td align="Center">&nbsp;</td></tr>
											<tr><td align="Center"><%=sjob_name%></td></tr>
										<% end if %>
										</table>
							    	</div>
						    	</td>
						    	
								<%
									ELSE
								%>
								  	<td valign="top"  height="100%">
										<div id="dAP1">
										<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0"  height="100%">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">&nbsp;</td></tr>
										<tr><td height="100" valign="bottom"></td></tr>
										</table>
										</div>
									</td>
									<td valign="top"  width="180"  height="100%">
									    <div id="dAP_H">
									    <table width="100%" cellpadding="5" cellspacing="0" class="a"  height="100%">
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">합의</td></tr>
											<tr><td align="Center">&nbsp;</td></tr>
											<tr><td align="Center"></td></tr>
											</table>
									    </div>
								    </td>
									<td valign="top"  width="180"  height="100%">
										<div id="dAP0">
										<table width="100%" cellpadding="5" cellspacing="0" class="a"  height="100%">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">최종승인</td></tr>
										<tr><td align="Center">&nbsp;</td></tr>
										<tr><td align="Center"><%=sjob_name%></td></tr>
										</table>
										</div>
									</td>
								<%
									END IF 
								%>
								</tr>  
						</table>
					</div> 
					<input type="hidden" name="hidATxt" id="hidATxt" value="<%=arrATxt%>">
					<input type="hidden" name="hidAJ" id="hidAJ" value="<%=arrAJSn%>">
					<input type="hidden" name="hidALI" id="hidALI" value="<%=AuthID_L%>"><!--최종결재자아이디-->
					<input type="hidden" name="hidALTxt" id="hidALTxt" value="<%=AuthName_L%>&nbsp;<%=AuthJobName_L%>">
					<input type="hidden" name="hidALJ" id="hidALJ" value="<%=AuthJobsn_L%>">
					<input type="hidden" name="hidAHI" id="hidAHI" value="<%=AuthID_F%>"><!--최종합의자아이디-->
					<input type="hidden" name="hidAHTxt" id="hidAHTxt" value="<%=AuthName_F%>&nbsp;<%=AuthJobName_F%>">
					<input type="hidden" name="hidAHJ" id="hidAHJ" value="<%=AuthJobsn_F%>">
					<input type="hidden" name="hidRfI" id="hidRfI" value="<%=sreferId%>"><!--참조아이디-->  
					<input type="hidden" name="blnL" id="blnL" value="<%=blnLast%>"><!--최종승인자 등록여부-->
					<input type="hidden" name="hidAI_H" id="hidAI_H" value="<%=hidAI_H%>"><!--합의자 아이디-->
					<input type="hidden" name="hidATxt_H" id="hidATxt_H" value="<%=AuthName_H%>&nbsp;<%=AuthJonName_H%>"> 
					<input type="hidden" name="hidAN" value="<%=sRectAuthName%>">		 
					<input type="hidden" name="hidAI" value="<%=sNextAuthId%>">
					<input type="hidden" name="hidPS" value="<%=iRectPartSn%>">
					<input type="hidden" name="iAP" value="<%=iNextPosition%>">
					<input type="hidden" name="iRAP" value="<%=iRectPosition%>">
					<input type="hidden" name="iRAT" value="<%=sRectAuthType%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>"  align="center">팀/부서</td>
					<td bgcolor="#FFFFFF"><%=spart_name%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성자</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성일</td>
					<td bgcolor="#FFFFFF"><%=formatdate(dregdate,"0000-00-00")%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
					<td bgcolor="#FFFFFF"><%=fnGetReportState(ireportstate)%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">참조</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sRfN" id="sRfN" value="<%=sReferName%>" size="20" style="border:0;" class="input" readonly></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="60" rowspan="2" valign="top" align="center">품의내용</td>
					<td>품의서 IDX</td>
					<td>품의서명</td>
					<td>품의금액</td>
					<td>결제타입</td>
					<td>SCM 문서번호</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
					<td><%=ireportidx%></td>
					<td><%=sreportname%><input type="hidden" name="sRN" value="<%=sreportname%>"></td>
					<td><%=formatnumber(mreportprice,0)%><input type="hidden" name="mRP" size="20" maxlength="20"  value="<%=formatnumber(mreportprice,0)%>"> </td>
					<td>
						 <%=fnGetPayType(ipaytype)%>
						<div  id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;">
							 <%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%>
						</div>
					</td>
					<td  align="Center"><input type="hidden" name="iSL" value="<%=iscmlinkno%>" ><%IF iscmlinkno> 0 THEN%><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');"><%=iscmlinkno%> <%IF sscmLink <> "" THEN%>>>상세보기<%END IF%></a><%END IF%></td>
				 </tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" width="60" rowspan="3"  align="center">내용</td>
					<td bgcolor="#FFFFFF"  height="200" valign="top">
					<%=tContents%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2" width="60">첨부서류</td>
					<td>첨부파일</td>
					<td>관련링크</td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<td align="center" valign="top">
						<div id="dFile">
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2)
							IF arrFile(2,intF) = 0 THEN Exit For
								arrF = split(arrFile(1,intF),"/")
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0)
						%>
						 <a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>  <Br>
						<%Next
						ELSE
						%>
						&nbsp;
						<%
						END IF
						%>
						</div>
					</td>
					<td><%
						IF isArray(arrFile) THEN
						 For intF2 = intF To UBound(arrFile,2)%>
						<a href="javascript:jsFileLink('<%=arrFile(1,intF2)%>')"><%=arrFile(1,intF2)%></a><br>
						<% Next
					END IF
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
					<td rowspan="2" width="60">계정과목</td>
					<td>수지항목</td>
					<td>연결계정과목</td>
				</tr>
				<tr bgcolor="#FFFFFF"  align="center">
					<td>[<%=iarap_cd%>]<%=sarap_nm%></td>
					<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
				</tr>
				</table>
			</td>
		</tr>
		<%IF   blnPayEapp THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="60" rowspan="2" style="padding:5px">부서별<br>자금구분</td>
					<td width="300" style="padding:5px" > 부서</td>
					<td width="205" style="padding:5px"> 금액</td>
					<td width="205" style="padding:5px"> %</td>
				</tr>
				<tr>
					<td colspan="3" bgcolor="#FFFFFF" valign="top">
					<div id="divPM">
					<%dim arrPV, arrPT
					IF isArray(arrPart) THEN %>
					<table border="0" cellpadding="3" cellspacing="0" class="a" width="760">
					<%	For intP = 0 To UBound(arrPart,2)
							IF intP > 0 THEN
								arrPV = arrPV&","
								arrPT =arrPT&","
							END IF
							arrPV = arrPV&arrPart(1,intP)
							arrPT = arrPT&arrPart(3,intP)
					%>
						<tr>
							<td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(4,intP)%></td>
							<td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(3,intP)%> </td>
							<td width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><input type="hidden" name="mPM" size="20" value="<%=formatnumber(arrPart(2,intP),0)%>"><%=formatnumber(arrPart(2,intP),0)%> 원</td>
							<td width="200" style="border-bottom:1px solid #BABABA;" align="center"><%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=formatnumber((arrPart(2,intP)/mreportprice)*100,1)%><%END IF%> %</td>
						</tr>
					<%	Next %>
					</table>
					<%END IF%>
					</div><br>
					<input type="hidden" name="iP" id="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" id="sP" value="<%=arrPT%>">
					<input type="hidden" name="mP" id="mP" value="">
					<%IF blnMod = 1 THEN%><br>&nbsp;	<input type="button" value="부서 등록/수정" onClick="jsSetPartMoney(1,'<%=sacc_use_cd%>','<%=sACC_GRP_CD%>');" class="button"><Br><Br><%END IF%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%END IF%> 
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
		<%IF blnMod =1  THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0">
					<tr>
						<td  align="left">
							<input type="button" value="내용수정" class="button" onClick="jsEappUpdate();">
						</td>
						<td align="right">
						<input type="button" value="보류" class="button" onClick="jsEappConfirm(3);">
						&nbsp;<input type="button" value="반려" class="button" onClick="jsEappConfirm(5);">
						<% if (sRectAuthType="A") then %>
						&nbsp;<input type="button" value="합의승인" class="button" onClick="jsEappConfirm(1);">
						<% else %>
						&nbsp;<input type="button" value="승인" class="button" onClick="jsEappConfirm(1);"  >
						<% end if %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%END IF%>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td style="padding-top:20px;">
		<!-- #include virtual="/admin/approval/eapp/incComment.asp" -->
	</td>
</tr>
<tr>
	<td height="50"></td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
</body>
</html>
