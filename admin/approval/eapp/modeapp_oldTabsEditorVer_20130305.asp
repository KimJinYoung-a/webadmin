<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 전자결재 수정
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
Dim clseapp,clsMem
Dim ireportidx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,ireportstate,sreferid
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_nm,sacc_use_cd,sedmsName,sedmscode ,sscmsubmitLink
Dim ipart_sn,ilastApprovalid,sjob_name,sscmLink,spart_name, susername 
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim sReferName,sEappName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast,iNextAuthState,blnMod  ,iLastposition
Dim blnpayEapp,		mpayrequestprice					
ireportidx =  requestCheckvar(Request("iridx"),10) 
Dim iRectMenu 
Dim ipayrequestidx 
Dim sCurrencyPrice,ipaytype,sCurrencyType
Dim sACC_GRP_CD

	ipayrequestidx = 0
 iRectMenu = requestCheckvar(Request("iRM"),10) 
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
	sacc_nm          = clseapp.Facc_nm   
	sacc_use_cd			 = clseapp.Facc_use_cd       	
	sedmsName        = clseapp.FedmsName         
	sedmscode        = clseapp.Fedmscode         
	ilastApprovalid  = clseapp.FlastApprovalid   
	sscmLink				  = clseapp.FscmLink					
	sscmsubmitLink	= clseapp.FscmsubmitLink		
	sjob_name			  = clseapp.Fjob_name					
	ipart_sn				  = clseapp.Fpart_sn					
	spart_name			  = clseapp.Fpart_name				
	susername				= clseapp.Fusername				
	blnpayEapp			= clseapp.FispayEapp
	mpayrequestprice = clseapp.Fpayrequestprice	
	ipaytype				= clseapp.Fpaytype	
	sCurrencyType		= clseapp.FCurrencyType	
  sCurrencyPrice	= clseapp.FCurrencyPrice	
  sACC_GRP_CD			= clseapp.FACC_GRP_CD
   	
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
 
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->  
<script type="text/javascript" src="eapp.js"></script>
</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4">
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">  

<tr>
	<td>
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<form name="frm" method="post" action="proceapp.asp">  
		<input type="hidden" name="hidM" value="U">
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
		<input type="hidden" name="hidRS" value="<%=ireportstate%>">
		<input type="hidden" name="iAIdx" value="<%=iarap_cd%>">  
		<input type="hidden" name="iAP" value="<%=iNextPosition%>"> 
		<input type="hidden" name="hidAid" value="<%=sadminid%>">
		<input type="hidden" name="hidRfI" id="hidRfI" value="<%=sreferId%>">
		<input type="hidden" name="hidAI" id="hidAI" value="<%=sNextAuthId%>">
		<input type="hidden" name="hidPS" value="<%=session("ssAdminPsn")%>">
		<input type="hidden" name="iLAID" value="<%=ilastApprovalid%>">   
		<input type="hidden" name="hidUN" value="<%=susername%>">
		<input type="hidden" name="hidAN" value="">
		<input type="hidden" name="iRM" value="<%=iRectMenu%>">
		<input type="hidden" name="hidPE" value="<%=blnPayEapp%>">
		<Tr>
			<td align="right" style="border-bottom:1px dashed #cccccc;"><input type="button" value="프린트" class="button" onClick="jsPopModPrint(<%=ireportidx%>);"></td>
		</tr> 
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b><%=sEappName%></b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="60" align="center">문서코드</td>
					<td bgcolor="#FFFFFF"><%=sedmscode%></td>
					<td rowspan="6" bgcolor="#FFFFFF" valign="top" width="500"><!--결재자 리스트-->
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a">
						<tr align="center">
							<%IF isArray(arrAuth) THEN
								For intA = 0 To UBound(arrAuth,2) 
									IF arrAuth(4,intA) THEN
										blnLast = 1  
										iLastposition=arrAuth(1,intA)
							%>
								<td valign="top" width="150">
								<div id="dAP<%=intA+1%>">
								<table width="100%" cellpadding="5" cellspacing="0" class="a" width="100%">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>	
								</table>
								</div>
								</td>
							<%			
										Exit For
									END IF	 
								%>
							<td valign="top" width="150">
								<div id="dAP<%=intA+1%>">
									<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>"><%=intA+1%>차 검토</td></tr>
									<%IF  iNextPosition = arrAuth(1,intA) and blnMod = 1  THEN  %>
									<tr><td align="Center"><input type="text" name="sASD" style="border:0;text-align:center;" value="<%=fnGetAuthState(arrAuth(3,intA))%>"></td></tr>
									<tr><td align="Center"><input type="text" name="sALN" id="sALN" value="<%=arrAuth(7,intA)&" "&arrAuth(10,intA)%>" style="border:0;text-align:center;" readonly size="20"><input type="hidden" name="hidAJ" id="hidAJ"  value="<%=arrAuth(10,intA)%>"></td></tr>
									<tr><td align="Center"><input type="text" name="sADD" value="<%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%>" style="border:0;text-align:center;"></td></tr>
									<tr><td align="Center"><input type="button" class="button" value="결재자 등록" onClick="jsRegID(1);"><br>
										<input type="checkbox" value="1" name="chkSms" <%IF arrAuth(11,intA) THEN%> checked<%END IF%>> SMS전송
										</td></tr>
									<%ELSE%>
									<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>
									<tr><td align="Center"><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td></tr>
									<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td></tr>
									<tr><td align="Center">&nbsp;</td></tr>
									<%END IF%>
									</table> 
								</div>
							</td>
							<% Next 
							 ELSE		
							%>
							<td valign="top" width="150">
							<div id="dAP1">
									<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">1차 검토</td></tr>
									<tr><td align="Center"><input type="text" name="sASD" style="border:0;" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sALN" id="sALN" value="" style="border:0;" readonly size="20"><input type="hidden" name="hidAJ" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sADD" value="" style="border:0;text-align:center;"></td></tr>
									<tr><td align="Center"><input type="button" class="button" value="결재자 등록" onClick="jsRegID(1);"><br>
										<input type="checkbox" value="1" name="chkSms" checked> SMS전송</td></tr>
									</table> 
							</div>
							</td>
							<%END IF%>
							<td valign="top">
								<table width="100%" cellpadding="5" cellspacing="0" class="a" width="100%">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>	
								</table>
							</td>
							<td valign="top"  width="150">
								<div id="dAP0">
									<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">최종승인자</td></tr> 
								<%IF blnLast=1 THEN%>
									<%IF  iNextPosition = arrAuth(1,intA) and blnMod = 1  THEN  %>
									<tr><td align="Center"><input type="text" name="sASD" style="border:0;text-align:center;" value="<%=fnGetAuthState(arrAuth(3,intA))%>"></td></tr>
									<tr><td align="Center"><input type="text" name="sALN" id="sALN" value="<%=arrAuth(7,intA)&" "&arrAuth(10,intA)%>" style="border:0;text-align:center;" readonly size="20"><input type="hidden" name="hidAJ" id="hidAJ" value="<%=arrAuth(10,intA)%>"></td></tr>
									<tr><td align="Center"><input type="text" name="sADD" value="<%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%>" style="border:0;text-align:center;"></td></tr>
									<tr><td align="Center"><input type="button" class="button" value="결재자 등록" onClick="jsRegID(1);document.frm.blnL.value=1;"><br>
										<input type="checkbox" value="1" name="chkSms" <%IF arrAuth(11,intA) THEN%> checked<%END IF%>> SMS전송</td></tr>
									<%ELSE%>
									<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>	
									<tr><td align="Center"><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td></tr>	
									<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td></tr>	
									<tr><td align="Center">&nbsp;</td></tr>	
									<%END IF%>	
								<%ELSE%>		
									<tr><td align="Center">&nbsp;</td></tr>	
									<tr><td align="Center"><%=sjob_name%></td></tr>
									<tr><td align="Center">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>		
								<%END IF%>
									</table>
								</div>
							</td> 
						</tr>  
						<input type="hidden" name="blnL" value="<%=blnLast%>">		
						</table>
					</td> 
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">팀/부서</td>
					<td bgcolor="#FFFFFF"><%=spart_name%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성자</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">작성일</td>
					<td bgcolor="#FFFFFF"><%IF ireportstate > 0 THEN%><%=formatdate(dregdate,"0000-00-00")%><%ELSE%><%=date()%><%END IF%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
					<td bgcolor="#FFFFFF"><%=fnGetReportState(ireportstate)%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">참조</td>
					<td bgcolor="#FFFFFF"><%IF blnMod = 1 THEN%><input type="button" class="button" value="참조 등록" onClick="jsRegID(2);"><%END IF%><input type="text" name="sRfN" id="sRfN" value="<%=sReferName%>" size="20" style="border:0;" readonly></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="60" rowspan="2" valign="top" align="center">품의내용</td>
					<td>IDX</td>
					<td>품의서명</td>
					<td>품의금액</td>
					<td>결제타입</td> 
					<td>SCM 문서번호</td> 
				</tr>
				<tr bgcolor="#FFFFFF" align="center"> 
					<td  align="Center"><%=ireportidx%></td>
					<td><input type="text" name="sRN" size="40" maxlength="60" value="<%=sreportname%>" <%IF blnMod = 0  THEN%>style="border:0" readonly<%END IF%>></td>
					<td><input type="text" name="mRP" size="15" maxlength="20"  value="<%=formatnumber(mreportprice,0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0" readonly<%ELSE%>"<%END IF%> <%IF not blnPayEapp THEN%>disabled class="text_ro"<%END IF%>  onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" ></td>
					<td>
						<%  
						IF blnmod= 0 THEN 
							%><%=fnGetPayType(ipaytype)%>
					 <%ELSE%>	
							<select name="selPT" onChange="jsChFC();" class="select" <%IF not blnPayEapp THEN%>disabled<%END IF%>>
								<%sboptPayType ipaytype%> 
							</select>
						<%END IF%> 
						<div  id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							<%IF blnMod=0 THEN%><%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%>
							<%ELSE%><%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"> 
							<%END IF%>
						</div>
					</td>
					<td  align="Center"><input type="hidden" name="iSL" value="<%=iscmlinkno%>" ><%IF iscmlinkno> 0 THEN%><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');"><%=iscmlinkno%> <%IF sscmLink <> "" THEN%> >>상세보기<%END IF%></a><%END IF%></td>
				 </tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" width="60" rowspan="3" align="center">내용</td>
					<td bgcolor="#FFFFFF" height="200"> 
					<%IF blnMod = 1  THEN%>
					<!--#Include Virtual = "/admin/approval/eapp/incEditor.asp" -->	
					<%ELSE%>
					<%=tContents%><div style="display:none;"><textarea name="editor" readonly><%=tContents%></textarea></div>
					<%END IF%>
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
						<%IF blnMod = 1   THEN %><input type="button" value="파일첨부" class="button" onClick="jsAttachFile('');"><%END IF%>
						<div id="dFile"> 
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2)
							IF arrFile(2,intF) = 0 THEN Exit For
								arrF = split(arrFile(1,intF),"/")  
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0) 
						%>
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>&nbsp;<%IF blnMod = 1 THEN %><input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"> <%END IF%>
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
						<input type="text" name="sL" size="60" maxlength="120" value="<%=arrFile(1,intF2)%>" <%IF blnMod = 0 THEN%>style="border:0;cursor:hand;" readonly  onClick="jsFileLink('<%=arrFile(1,intF2)%>');"<%END IF%>><br>
						<% iCount = iCount + 1
						Next
						END IF 
						For intF3= iCount To 4
						%>
						<input type="text" name="sL" size="60" maxlength="120" <%IF blnMod = 0  THEN%>style="border:0" readonly<%END IF%>><br>
						<%Next 
						%> 
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%IF iarap_cd <> "0" THEN%>
		<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="60">계정과목</td> 
							<td>수지항목</td>
							<td>연결계정과목</td> 
						</tr>
						<tr bgcolor="#FFFFFF"  align="center"> 
							<td>[<%=iarap_cd%>] <%=sarap_nm%></td>
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
									<td width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><input type="text" name="mPM" size="20" value="<%=formatnumber(arrPart(2,intP),0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('m',<%=intP%>,1);"> 원</td>
									<td width="200" style="border-bottom:1px solid #BABABA;" align="center"><input type="text" name="iPM"  size="4" value="<%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=formatnumber((arrPart(2,intP)/mreportprice)*100,1)%><%END IF%>"  style="text-align:right;<%IF blnMod = 0  THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('i',<%=intP%>,1);">%</td>
								</tr> 
							<%	Next %>
							</table>
							<%END IF%>
							</div>
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
		<%END IF%>
		<%IF isArray(arrReturn) THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="60">반려리스트</td> 
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
			<td align="center">
				<table border="0" cellpadding="5" cellspacing="0" width="100%">
				<tr>
					<td align="left">
					<input type="button" value="작성중으로(임시저장)" class="button" onClick="jsEappSubmit(0);">&nbsp;
					<input type="button" value="삭제" class="button" onClick="jsEappSubmit(-1);" style="color:red;">
					</td>
					<td align="right"> <input type="button" value="결재등록" class="button" onClick="jsEappSubmit(1);"></td> 
				</tr>
				</table>
			</td>
		</tr>
		<%ELSEIF sRectAuthId = session("ssBctId") and (ireportstate =7 or ireportstate = 8 ) and mreportPrice > mpayrequestprice and blnpayEapp THEN %>
		<tr>
			<td width="100%" align="right">
			<!--	<input type="button" value="품의금액 추가품의" class="button" onClick="jsPopView('/admin/approval/eapp/regAddeapp.asp?iridx=<%=ireportidx%>&iLp=<%=iLastposition%>');">-->
				<input type="button" value="결재등록" class="button" onClick="jsPopView('/admin/approval/eapp/regpayrequest.asp?iridx=<%=ireportidx%>');">
			</td>
		</tr>
		<%END IF%>
		</form>
		</table>
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
