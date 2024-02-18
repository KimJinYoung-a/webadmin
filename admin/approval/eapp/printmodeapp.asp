<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 전자결재 수정
' History : 2011.03.14 정윤정 생성
'			2017.05.16 한용민 수정
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim clseapp,clsMem
Dim ireportidx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,ireportstate,sreferid
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_use_cd,sacc_nm,sedmsName,sedmscode ,sscmsubmitLink
Dim ipart_sn,ilastApprovalid,sjob_name,sscmLink,spart_name, susername 
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim sReferName,sEappName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast,iNextAuthState,blnMod  
							
ireportidx =  requestCheckvar(Request("iridx"),10) 
 	 
'결재 기본 폼 정보 가져오기
set clseapp = new CEApproval
	clseapp.Freportidx = ireportidx 
	clseapp.fnGetEAppData
	   
	iarap_cd				 = clseapp.Farap_cd
	sreportName      = clseapp.FreportName       
	mreportPrice     = clseapp.FreportPrice      
	iscmlinkno       = clseapp.Fscmlinkno        
	sbigo            = clseapp.Fbigo             
	tContents  			= clseapp.Freportcontents   
	ireportstate     = clseapp.Freportstate      
	sreferid         = clseapp.Freferid          
	sadminid         = clseapp.Fadminid          
	dregdate         = clseapp.Fregdate          
	sarap_nm         = clseapp.Farap_nm        	
	sacc_cd         = clseapp.Facc_cd   
	sacc_use_cd			 = clseapp.Facc_use_cd        	
	sacc_nm          = clseapp.Facc_nm          	
	sedmsName        = clseapp.FedmsName         
	sedmscode        = clseapp.Fedmscode         
	ilastApprovalid  = clseapp.FlastApprovalid   
	sscmLink				  = clseapp.FscmLink					
	sscmsubmitLink	= clseapp.FscmsubmitLink		
	sjob_name			  = clseapp.Fjob_name					
	ipart_sn				  = clseapp.Fpart_sn					
	spart_name			  = clseapp.Fpart_name				
	susername				= clseapp.Fusername					
   
	arrAuth			= clseapp.fnGetAuthLineList
	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList 
	arrReturn		= clseapp.fnGetAuthLineReturnList 
 	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing 
 
'부서명 가져오기
set clsMem = new CTenByTenMember 
 	if sreferid <> "" then
 	clsMem.Fuserid	= sreferid
	arrRefer		= clsMem.fnGetInIDOutName
	end if
 set clsMem = nothing
  
 IF iarap_cd  <> "0" THEN 
 	sEappName = sedmsname&"_"&sarap_nm
 ELSE
 	sEappName = sedmsname 
 END IF	 

'결재리스트-----------------------------------------
blnMod = 0  		'문서 수정 가능여부
blnLast = 0 		'최종결재여부
iRectPosition = 0	'현재결재위치 
iNextPosition = 1	'다음결재위치
sNextAuthId = ""	'다음결재자아이디
iNextAuthState = 0	'다음결재상태
sRectAuthId = sadminid	 '현재결재 아이디 = 결재등록자

IF isArray(arrAuth) THEN  
	 	sNextAuthId	 = arrAuth(2,0)
	 	iNextAuthState = arrAuth(3,0)   
END IF
 
'--------------------------------------------------	   

 '문서 수정가능여부
 IF(iReportState = 0  OR  iReportState = 5 ) AND sRectAuthId = session("ssBctId") THEN
 	blnMod = 1
 END IF	
  
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
<!-- #include virtual="/lib/db/dbclose.asp" -->    
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script type="text/javascript" src="eapp.js"></script>
<link rel="stylesheet" href="eapp.css" type="text/css"> 
</head> 
<body topmargin="0" leftmargin="0"> 
<table width="840" height="100%" cellpadding="0" cellspacing="0" class="a" align="center" border="0">   
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
					<tr><td style="padding-bottom:10px;"><%=spart_name%></td></tr>
					<tr><td class="btdsmall">작성자:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=susername%></td></tr>
					<tr><td class="btdsmall">작성일:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=formatdate(dregdate,"0000-00-00")%></td></tr> 
					<%IF sReferName <> "" THEN%>
					<tr><td class="btdsmall">참조</td></tr>
					<tr><td style="font-size:14px;padding-bottom:10px;"><%=sReferName%>&nbsp;</td></tr>
					<%END IF%>
						<%IF isArray(arrAuth) THEN
									For intA = 0 To UBound(arrAuth,2) 
									 
							%>
					<tr>
						<td style="padding-bottom:5px;">
								<table border=1 cellspacing=0 cellpadding=3 class="a" width="100%"> 
							<tr>
								<td  class="btdsmall"><%IF arrAuth(4,intA) ="A" THEN%>합의<%ELSEIF arrAuth(4,intA) ="L" THEN%>최종승인<%ELSE%><%=intA+1%>차 검토<%END IF%></td>
							</tr>
							<tr>
								<td><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td>
							</tr>
							<tr>
								<td><%=fnGetAuthState(arrAuth(3,intA))%></td>
							</tr>
							<tr>
								<td><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td>
							</tr> 
						</table>
						</td>
					</tr>
					<%
									Next
								END IF
							%>
				</table>
			</td>
		</tr>
			<tr>
				<td valign="bottom">
					<table border=0 cellspacing=0 cellpadding=0   width="100%">
					<tr>
						<td style="padding-bottom:10px;">(주)텐바이텐</td>
					</tr>
					<tr>
						<td  style="padding-bottom:5px;"> 03082<br>
					 	서울시 종로구 대학로 57<br>
					 	홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐
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
						<td valign="bottom" class="btd20"> idx <%=ireportidx%>  </td>  
						<td align="right" valign="top"><!--<img src="/images/10x10-logo400px.jpg">--></td>
					</tr> 
					<tr>
						<td colspan="2" class="btd20" valign="top"><%=sEappName%></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td ><br>[정보]<br>
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
				<tr>
				<td style="padding-top:15px;"><br>[내용]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
						<tr>
							<td style="border-bottom:1px solid #bbbbbb;"><%=tContents%></td>
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
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount
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
								<td  class="td02" align="center"> 
									<%=arrPart(4,intP)%>
							  	>
									<%=arrPart(3,intP)%>
								</td>
								<td  class="td02" align="center"><%=formatnumber(arrPart(2,intP),0)%> 원</td>
								<td  class="td02" align="center"><%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=(arrPart(2,intP)/mreportprice)*100%><%END IF%>%</td>
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
				</table><Br>
			</td>
		</tr>
		</table>
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
