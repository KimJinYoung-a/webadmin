<%@ language="VBScript" %>
<% option explicit %> 
<% 
'###########################################################
' Description : 전자결재 수정
' History : 2011.03.14 정윤정  생성
'           2013.03.05 허진원 - 이노디터로 변경
'						2016.06	정윤정- 다음에디터로 변경
'################################################################## 
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
Dim icid1,icid2,icid3,icid4,idepartment_id,sdepartmentnamefull
Dim isAgreeNeed, isAgreeNeedTarget
dim iedmsidx
Dim addFileName, addFileNamePh

ipayrequestidx = 0
iRectMenu = requestCheckvar(Request("iRM"),10)
'get default form 
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
	idepartment_id	  = clseapp.Fdepartment_id
	sdepartmentnamefull= clseapp.Fdepartmentnamefull
	susername				= clseapp.Fusername
	blnpayEapp			= clseapp.FispayEapp
	mpayrequestprice = clseapp.Fpayrequestprice
	ipaytype				= clseapp.Fpaytype
	sCurrencyType		= clseapp.FCurrencyType
  sCurrencyPrice	= clseapp.FCurrencyPrice
  sACC_GRP_CD			= clseapp.FACC_GRP_CD
  icid1				= clseapp.Fcid1
  icid2				= clseapp.Fcid2
  icid3				= clseapp.Fcid3
  icid4				= clseapp.Fcid4
	isAgreeNeed			= clseapp.FisAgreeNeed
	isAgreeNeedTarget	= clseapp.FisAgreeNeedTarget
	iedmsidx 	= clseapp.FedmsIdx
	arrAuth			= clseapp.fnGetAuthLineList
	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList
	arrReturn		= clseapp.fnGetAuthLineReturnList
 	arrPart			= clseapp.fnGetPartMoneyList
	addFileName		= getEdmsFileName(clseapp.FedmsCode, clseapp.FedmsName, clseapp.FedmsFile, addFileNamePh)
set clseapp = nothing

'get partname
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
 

'권한관리 -----------------------------------------
blnMod = 0  		'문서 수정 가능여부
blnLast = 0 		'최종결재여부
iRectPosition = 0	'현재결재위치
iNextPosition = 1	'다음결재위치
sNextAuthId = ""	'다음결재자아이디
iNextAuthState = 0	'다음결재상태
sRectAuthId = sadminid	 '현재결재 아이디 = 결재등록자 
'--------------------------------------------------

 'write auth---------------------------------- 
 IF   (iReportState = 0  OR  iReportState = 5 ) AND sRectAuthId = session("ssBctId")    THEN
 	blnMod = 1
 END IF
 

 'refer list -------------------------------------
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
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>   
<%IF blnMod = 1  THEN%>  
<!-- daumeditor head -------------------------> 
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="euc-kr"/>    
<script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="euc-kr"></script> 
<script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="euc-kr"></script> 
<!-- daumeditor  --> 
<script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frm',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",
        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                } 
            }
        }
    }; 
   
</script> 
<!-- //daumeditor head ------------------------->
<% end if %>
<script type="text/javascript" src="eapp.js?t=<%=left(now(),10)%>"  charset="euc-kr"/></script>  
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
		<input type="hidden" name="iaidx" value="<%=iarap_cd%>"> 
		<input type="hidden" name="sACC" value="<%=sacc_cd%>"> 
		<input type="hidden" name="hidAid" value="<%=sadminid%>">  
		<input type="hidden" name="hidPS" value="<%=session("ssAdminPsn")%>">
		<input type="hidden" name="iLAID" value="<%=ilastApprovalid%>"> 
		<input type="hidden" name="hidJN" value="<%=sjob_name%>"> 
		<input type="hidden" name="hidUN" value="<%=susername%>"> 
		<input type="hidden" name="iRM" value="<%=iRectMenu%>">
		<input type="hidden" name="hidPE" value="<%=blnPayEapp%>">
		<input type="hidden" name="hidcid1" value="<%=icid1%>">
		<input type="hidden" name="hidcid2" value="<%=icid2%>">
		<input type="hidden" name="hidcid3" value="<%=icid3%>">
		<input type="hidden" name="hidcid4" value="<%=icid4%>">
	 
		<Tr>
			<td align="right" style="border-bottom:1px dashed #cccccc;"><input type="button" value="프린트" class="button" onClick="jsPopModPrint(<%=ireportidx%>);"></td>
		</tr>
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b><%=sEappName%></b></td> 
					<td align="right" width="100"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				<tr>
					<td colspan="2" align="right" >
							<!--%IF blnmod = 1 THEN%--><input type="button" class="button" style="color:blue;" value="결재선등록" onClick="jsRegID(1);"><!--%END IF%-->
						<input type='button' class='button' value='전결규정보기' onClick='popDecision();'>
					</td>
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
						<div id="dAP">
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a">
								<tr align="center"> 
									<%
									Dim AuthID_L,AuthState_L,AuthName_L,AuthJobsn_L,AuthJobName_L,AuthConfirmTime_L,AuthSMS_L
									Dim AuthID_F,AuthState_F,AuthName_F,AuthJobsn_F,AuthJobName_F,AuthConfirmTime_F,AuthSMS_F
									Dim intNo, arrAID, arrATxt, arrAJSn 
									Dim hidAI_H, hidPS_H, AuthState_H, AuthName_H, AuthJonName_H, AuthConfirmTime_H, AuthSMS_H, hidAJ_H
									Dim isRect
									intNo = 0
									arrAID =""
									arrATxt=""
									arrAJSn=""
									isRect = 1
									IF isArray(arrAuth) THEN 
								    For intA = 0 to UBound(arrAuth,2)  
								    		if  arrAuth(3,intA) = 1 then
								    			isRect = isRect  + 1
								    	  end if
								        IF  arrAuth(4,intA) ="A"  THEN ''합의 
								            hidAI_H         = arrAuth(2,intA)
														hidPS_H         = arrAuth(9,intA)
														AuthState_H     = arrAuth(3,intA)
														AuthName_H      = arrAuth(7,intA)
														hidAJ_H         = arrAuth(8,intA)
														AuthJonName_H   = arrAuth(10,intA)
								            AuthConfirmTime_H = arrAuth(6,intA)
								            AuthSMS_H       = arrAuth(11,intA)
								        ELSEIF arrAuth(4,intA)="L" THEN  
								    	 			AuthID_L       	= arrAuth(2,intA) 
														AuthState_L     = arrAuth(3,intA)
														AuthName_L      = arrAuth(7,intA)
														AuthJobsn_L     = arrAuth(8,intA)
														AuthJobName_L   = arrAuth(10,intA)
								            AuthConfirmTime_L= arrAuth(6,intA)
								            AuthSMS_L       = arrAuth(11,intA)
								        ELSEIF arrAuth(4,intA)="F" THEN  
						    	 			AuthID_F       	= arrAuth(2,intA) 
											AuthState_F     = arrAuth(3,intA)
											AuthName_F      = arrAuth(7,intA)
											AuthJobsn_F     = arrAuth(8,intA)
											AuthJobName_F   = arrAuth(10,intA)
								            AuthConfirmTime_F= arrAuth(6,intA)
								            AuthSMS_F       = arrAuth(11,intA)
								    	 ELSE 
								    	 		intNo = intNo+1
								    	 		if arrAID = "" THEN  
								    	 			arrAID 		= arrAuth(2,intA)
														arrAJSn 	= arrAuth(8,intA) 
														arrATxt 	= arrAuth(7,intA)&" "&arrAuth(10,intA)
								    	 		else	
									    	 		arrAID 		= arrAID& ","&arrAuth(2,intA)
									    	 		arrAJSn 	= arrAJSn& ","&arrAuth(8,intA) 
									    	 		arrATxt 	= arrATxt& ","&arrAuth(7,intA)&" "&arrAuth(10,intA)
									    	 	end if	  
									%>
									<td valign="top" height="100%" width="180">
										<div id="dAP<%=intNo%>">
										<table width="100%"  cellpadding="5" cellspacing="0" class="a"  height="100%" border="0">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20"><%=intNo%>차 검토</td></tr>
											<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>
											<tr><td align="Center"><%=arrAuth(7,intA)%>&nbsp;<%=arrAuth(10,intA)%></td></tr>
											<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">	<%IF isRect =  intNo THEN%><input type='checkbox' value='1' name='chkSms' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr>
										</table>
										</div>
									</td>	
									<%								    	 			
								        end if
								    Next   
								   	if arrAID = "" THEN    
								   %>
								   	<td valign="top"  height="100%">
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
											<tr><td align="Center"><%IF arrAID="" and blnmod =1 THEN%>	<input type='checkbox' value='1' name='chkSms_H' checked> SMS|¼?%ELSE%>&nbsp;<%END IF%></td></tr> 
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
											<tr><td align="Center"><%IF arrAID="" AND hidAI_H="" and blnmod = 1 THEN%><input type='checkbox' value='1' name='chkSms' checked> SMS전송<%ELSE%>&nbsp;<%END IF%></td></tr> 
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
								  ELSE  %>
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
				<input type="hidden" name="hidAI" id="hidAI" value="<%=arrAID%>"><!--결재선아이디(결재순서->)-->
					<input type="hidden" name="hidATxt" id="hidATxt" value="<%=arrATxt%>">
					<input type="hidden" name="hidAJ" id="hidAJ" value="<%=arrAJSn%>">
					<input type="hidden" name="hidALI" id="hidALI" value="<%=AuthID_L%>"><!--최종결재자아이디-->
					<input type="hidden" name="hidALTxt" id="hidALTxt" value="<%=AuthName_L%>&nbsp;<%=AuthJobName_L%>">
					<input type="hidden" name="hidALJ" id="hidALJ" value="<%=AuthJobsn_L%>">
					<input type="hidden" name="hidAHI" id="hidAHI" value="<%=AuthID_F%>"><!--최종합의자아이디-->
					<input type="hidden" name="hidAHN" id="hidAHTxt" value="<%=AuthName_F%>&nbsp;<%=AuthJobName_F%>">
					<input type="hidden" name="hidAHJ" id="hidAHJ" value="<%=AuthJobsn_F%>">
					<input type="hidden" name="hidRfI" id="hidRfI" value="<%=sreferId%>"><!--참조아이디-->  
					<input type="hidden" name="blnL" id="blnL" value="<%=blnLast%>"><!--최종승인자 등록여부-->
					<input type="hidden" name="hidAI_H" id="hidAI_H" value="<%=hidAI_H%>"><!--합의자 아이디-->
					<input type="hidden" name="hidATxt_H" id="hidATxt_H" value="<%=AuthName_H%>&nbsp;<%=AuthJonName_H%>">
					<input type="hidden" name="hidPS_H" id="hidPS_H" value="<%=hidPS_H%>">
					<input type="hidden" name="tmpisAgreeNeed" value="<%=isAgreeNeed%>">
					<input type="hidden" name="tmpisAgreeNeedTarget" value="<%=isAgreeNeedTarget%>">	
					</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">팀/부서</td>
					<td bgcolor="#FFFFFF"><%=sdepartmentnamefull%></td>
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
					<td bgcolor="#FFFFFF"><input type="text" name="sRfN" id="sRfN" value="<%=sReferName%>" size="35" style="border:0;" readonly class="input"></td>
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
					<td><input type="text" name="sRN" class="text" size="40" maxlength="60" value="<%=sreportname%>" <%IF blnMod = 0  THEN%>style="border:0" readonly<%END IF%>></td>
					<td><input type="text" name="mRP" class="text" size="15" maxlength="20"  value="<%=formatnumber(mreportprice,0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0" readonly<%ELSE%>"<%END IF%>  onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" ></td>
					<td>
						<%
						IF blnmod= 0 THEN
							%><%=fnGetPayType(ipaytype)%><input type="hidden" name="selPT" value="<%=ipaytype%>">
					 <%ELSE%>
							<select name="selPT" onChange="jsChFC();" class="select" <%IF not blnPayEapp THEN%>disabled<%END IF%>>
								<%sboptPayType ipaytype%>
							</select>
						<%END IF%>
						<div  id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;">
							<%IF blnMod=0 THEN%><%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%>
							<input type="hidden" name="selCT" value="<%=sCurrencyType%>">
							<input type="hidden" name="sCP" value="<%=sCurrencyPrice%>">
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
					 <textarea name="editor" id="editor" style="width: 100%; height: 490px;"><%=tContents%></textarea>    
					 <script type="text/javascript"> 
					    EditorCreator.convert(document.getElementById("editor"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                            EditorJSLoader.ready(function (Editor) {
                                new Editor(config);
                                Editor.modify({
                                    content: '<%=tContents%>'
                                });
                            });
                        });  
					    </script>
                	<!-- daumeditor   -->
					<%ELSE%>
					 <%=tContents%> 
					<%END IF%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<% if addFileName<>"" then %>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" width="60" align="center">관련서류</td>
					<td bgcolor="#FFFFFF"><span onclick="jsEdmsDownload('<%=uploadImgUrl%>','<%=addFileName%>','<%=addFileNamePh%>');" style="cursor:pointer;" title="관련서류 양식 다운로드">▼ <%=addFileName%></span></td>
				</tr>
				</table>
			</td>
		</tr>
		<% end if %>
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
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>&nbsp;<%IF blnMod = 1 THEN %><input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"><%END IF%>
							<input type="hidden" name="sFileP[]"   value="<%= arrFile(1,intF)%>"></div>
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
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2" width="60">계정과목</td>
					<td>수지항목</td>
					<td>연결계정과목</td>
				</tr>
				<tr bgcolor="#FFFFFF"  align="center">
				    <!--
					<td>[<%=iarap_cd%>] <%=sarap_nm%></td>
					<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
					-->
					<td><input type="text" name="sANM" value="<%=CHKIIF(isNULL(sarap_nm),"","["&iarap_cd&"]"&sarap_nm)%>" style="border:0;width=200px" readonly ></td>
					<td><input type="text" name="sACCNM" value="<%=CHKIIF(isNULL(sacc_nm),"","["&sacc_use_cd&"]"&sacc_nm)%>" style="border:0;width=200px" readonly></td>
				</tr>
				<%IF blnMod = 1 THEN%><tr bgcolor="#FFFFFF">
					<td colspan="3"><input type="button" class="button" value="수지항목 수정" onClick="jsGetARAP();"></td>
				</tr><%END IF%>
				</table>
			</td>
		</tr>  
		<%IF blnPayEapp THEN%> 
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
									<td width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><input type="text" name="mPM" class="text" size="20" value="<%=formatnumber(arrPart(2,intP),0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0;width:90%;" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('m',<%=intP%>,1);"> 원</td>
									<td width="200" style="border-bottom:1px solid #BABABA;" align="center"><input type="text" class="text"  name="iPM"  size="4" value="<%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=formatnumber((arrPart(2,intP)/mreportprice)*100,1)%><%END IF%>"  style="text-align:right;<%IF blnMod = 0  THEN%>border:0;width:90%;" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('i',<%=intP%>,1);">%</td>
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
		<%if iedmsidx <>33 then %>
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
	<td height="100">&nbsp;</td>
</tr>
</table>

 <!-- #include virtual="/lib/db/dbclose.asp" -->
</body>
</html>
 