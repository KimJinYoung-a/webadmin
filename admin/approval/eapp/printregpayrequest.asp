<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
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
<!-- #include virtual="/lib/function.asp"-->
<%
Dim clsPay,clsMem,clseapp, clscomm,clsPM
Dim ireportidx,ipayrequestIdx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate ,scust_cd
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_use_cd,sacc_nm,sedmsName,sedmscode ,pcomment
Dim spartname ,ilastApprovalid,sscmLink,susername
Dim chkPayRequest 
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate 
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrProc,arrPM,arrPart
Dim intA, intC, intF, intR, intRA, intP , iAuthCount, intPM, intPart 
Dim blnMod 
Dim mSumPrice, mSumRealPrice 
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2), pmistate(2)
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	,scustnm,spayrequesttitle
Dim arrFName,arrF, sFName, intF2,intF3, iCount
Dim ipaytype, sCurrencyType, sCurrencyPrice					
ireportidx 		=  requestCheckvar(Request("iridx"),10)
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
 
 IF ipayrequestIdx = "" THEN ipayrequestIdx = 0 '��ϵ� ��û���� ���� ��� ǰ�Ǽ� ���� ���� �����ͼ� default�� �ѷ��ش�. 
 	
'���� �⺻ �� ���� ��������
set clsPay = new CPayRequest
	clsPay.Freportidx = ireportidx  
	clsPay.FpayrequestIdx = ipayrequestIdx 
	IF ipayrequestIdx <= 0 THEN '�űԵ���ϋ��� üũ
		chkPayRequest = clsPay.fnCheckPayRequest
		 
		IF chkPayRequest = 0 THEN 
			set clsPay = nothing 
	%>
	<!-- #include virtual="/lib/db/dbclose.asp" --> 
	<%		Alert_return "������û�� ����� �Ұ����մϴ�. �����͸� Ȯ�����ּ���" 
			response.end
		END IF	
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
  pcomment					= clsPay.FpayComment																													   
	scust_cd						= clsPay.Fcust_cd
	scustnm						=clsPay.Fcust_nm
	spayrequesttitle	=clsPay.FpayRequestTitle																									   
 	susername					=clsPay.Fusername						 
	spartname				  =clsPay.Fpartname
	ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	IF ipayrequestIdx = 0   THEN  
		ipayrequestIdx = -1 'ǰ�Ǽ� ������ ������û�� ������ üũ�� ���� (0=ǰ�Ǽ�)
		clsPay.FpayrequestIdx = ipayrequestIdx
	END IF	
	'//���������Ʈ
	arrProc			= clsPay.fnGetProcPayRequestList	 
		IF ipayrequestIdx > 0 THEN
	'//��������
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
 

	
'�������, �ڸ�Ʈ, ���� ����Ʈ ��������
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
 
 
'�������, �ڸ�Ʈ, ���� ����Ʈ ��������
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
 
 '���缱 ����Ʈ ����
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
		'�繫ȸ���� ������û�� ó��������
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
	
'�μ��� ��������
IF isNull(susername) or susername ="" THEN susername = session("ssBctCname")
IF isNull(spartname) or spartname ="" THEN 
set clsMem = new CTenByTenMember
	clsMem.Fpart_sn = session("ssAdminPsn")
	clsMem.fnGetPartName 
	spartname = clsMem.Fpart_name 
 set clsMem = nothing 
END IF 
%>
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script type="text/javascript" src="eapp.js"></script>
<link rel="stylesheet" href="eapp.css" type="text/css"> 
</head> 
<body topmargin="0" leftmargin="0"> 
<table width="840" height="100%" cellpadding="0" cellspacing="0" class="a" align="center" border="0">   
<tr>
	<td width="160" style="padding-right:10px;" ><!-- ���� �޴�-->
		<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0" >
			<Tr>
				<td valign="top">
					<table width="100%" border="0" cellpadding="0" cellspacing="0" >
					<tr><td><img src="/images/top_logo.gif"></td></tr>
					<tr><td class="btdsmall">�����ڵ�:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=sedmscode%></font></td></tr>
					<tr><td class="btdsmall">��/�μ�:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=spartname%></td></tr>
					<tr><td class="btdsmall">�ۼ���:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=susername%></td></tr>
					<tr><td class="btdsmall">�ۼ���:</td></tr>
					<tr><td style="padding-bottom:10px;"><%=formatdate(dregdate,"0000-00-00")%></td></tr> 
					<%IF isArray(arrAuth) THEN  %>
					<tr>
						<td style="padding-bottom:5px;">
								<table border=1 cellspacing=0 cellpadding=3   width="100%"> 
							<tr>
								<td  class="btdsmall">����������</td>
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
								<td  class="btdsmall">�繫ȸ����</td>
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
						<td style="padding-bottom:20px;">(��)�ٹ�����</td>
					</tr>
					<tr>
						<td  style="padding-bottom:5px;"> 03082<br>
					 	����� ���α� ���з� 57<br>
					 	ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����
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
	</td><!-- /���� �޴�-->
	<td  valign="top"><!-- �󼼳���-->
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr> 
				<td height="50"  >
					<table border="0" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td valign="bottom" class="btd20"> idx <%=ipayrequestidx%>  </td>  
							<td align="right" valign="top"><img src="/images/10x10-logo400px.jpg"></td>
					</tr> 
					<tr>
						<td colspan="2" class="btd20" valign="top">������û��(<%=sarap_nm%>)</td>
					</table>
				</td>
			</tr>
			<%IF ireportidx > 0 THEN%>
			<tr>
				<td><br>[����]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
						<tr  align="center">
							<td class="td01">idx</td>
							<td class="td01">ǰ�Ǽ���</td>
							<td class="td01">ǰ�Ǳݾ�</td>
							<td class="td01">scm������ȣ</td>
						</tr>
						<tr  align="center">
							<td><%=ireportidx%>&nbsp;</td>
							<td><%=sreportname%>&nbsp;</td>
							<td><%=formatnumber(mreportprice,0)%>&nbsp;</td>
							<td><%=iscmlinkno%>&nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<%END IF%>
				<% Dim totPrice 
		totPrice = 0
		IF isArray(arrProc) THEN%>
		<tr>
			<td style="padding-top:15px;"><br>[���������]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01">
				<tr align="center"> 
					<td class="td01">������û�� IDX</td>
					<td class="td01">����(�Ա�)��</td>
					<td class="td01">�����ݾ�</td> 
					<td class="td01">��������</td>  
					<td class="td01">�ڱݿ뵵</td>  
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
		<input type="hidden" name="hidTP" value="<%=totPrice%>"> 
		<tr>
				<td style="padding-top:15px;"><br>[����]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
					<tr  align="center" >
						<td class="td01">������û��</td>
						<td class="td01">������û�ݾ�</td>
						<td class="td01">�������</td>
						<td class="td01">���</td> 
					</tr> 
					<tr align="center" >	
					<td class="td02"><%IF dpayrequestdate <> "" THEN%><%=formatdate(dpayrequestdate,"0000-00-00")%><%END IF%></td>
					<td class="td02"><%=formatnumber(mpayrequestprice,0)%></td>
				 	<td class="td02"><%=fnGetPayType(ipaytype)%></td>
				 	<td class="td02"><span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							��ȭ�ݾ�: <%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%> 
							</span>&nbsp;</td>
				</tr> 
				<tr  align="center">	
					<td class="td01" colspan="4">�ڱݿ뵵</td>
				</tr>
				<tr  align="center">	
					<td class="td02" colspan="4"><%=spayrequesttitle%>&nbsp;</td>
				</tr>
				<tr  align="center">	
					<td class="td01">�ŷ�ó</td>
					<td class="td01">�����</td>
					<td class="td01">���¹�ȣ</td>
					<td class="td01">�����ָ�</td>
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
			<td style="padding-top:15px;"><br>[��������]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0" >
				<tr align="center">
					<td class="td01">��������</td>
					<td class="td01">ǰ�� </td> 
					<td class="td01">����û���ι�ȣ</td> 
					<td class="td01">������	</td>
				</tr>	
				<tr align="center">
					<td class="td02">
						<%	Dim strDoc
							IF ipaydockind ="1" THEN 
								strDoc = "���ݰ�꼭-����"
							ELSEIF ipaydockind ="2" THEN 
								strDoc = "���ݰ�꼭-����"
							ELSEIF ipaydockind ="3" THEN 
								strDoc = "���ݿ�����-�ҵ������"
							ELSEIF ipaydockind ="4" THEN 
								strDoc = "���ݿ�����-����������"
							ELSEIF ipaydockind ="5" THEN 
								strDoc = "��Ÿ������"
							ELSEIF ipaydockind ="8" THEN 
								strDoc = "��꼭 ���� ����"	
							ELSE
								strDoc = "��������"	
							END IF
							%><%=strDoc%>
					</td>
					<td class="td02"><%=sItemName%>&nbsp;</td> 
					<td class="td02"><%=setaxkey%>&nbsp;</td>
					<td class="td02"><%=dissuedate%>&nbsp;</td> 
				</tr> 
				<tr align="center">		
					<td class="td01">�������� </td>
					<td class="td01">�ѱݾ� </td>
					<td class="td01">���ް�</td>
					<td class="td01">�ΰ��� </td> 
				</tr>
			 <tr align="center">
			 	<td class="td02">
						<%
							Dim strVat
							IF sVatKind ="0" THEN 
								strVat = "����(�ΰ��� 10%) "
							ELSEIF sVatKind ="2" THEN 
								strVat = "�鼼"
							ELSEIF sVatKind ="3" THEN 
								strVat = "����" 
							END IF
						%> <%=strVat%>
						</td>
					<td class="td02"><%=formatnumber(mTotPrice,0)%></td>
					<td class="td02"><%=formatnumber(mSupplyPrice,0)%></td>
					<td class="td02"><%=formatnumber(mVatPrice,0)%></td> 
				</tr>	 
				</table>
			</td>
		</tr>
		<tr>
			<td style="padding-top:15px;"><br>[÷�μ���]<br>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0" >
				<tr   align="center"> 
					<td class="td01" width="50%">÷������</td>
					<td class="td01" width="50%">���ø�ũ</td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<%IF isArray(arrFile) THEN%>
					<td align="center" valign="top" > 
						<div id="dFile"> 
						<%  
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
					<td colspan="2" align="center" >÷�� ������ �����ϴ�.</td>
					<%END IF%>
				</tr>
				</table>
			</td>
		</tr>
		<%IF iarap_cd <> "0" THEN%>
		<tr>
					<td style="padding-top:15px;"><br>[��������]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr align="center"> 
							<td class="td01">�����׸�</td>
							<td class="td01">�����������</td> 
						</tr>
						<tr   align="center"> 
							<td>[<%=iarap_cd%>] <%=sarap_nm%></td>
							<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
						</tr>	
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:15px;"><br>[�μ��� �ڱݱ���]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr   align="center" > 
							<td class="td01"> �μ�</td>
							<td class="td01">�ݾ�</td>
							<td class="td01">�ݿ���</td>
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
									<%=arrPart(4,intP)%> > <%=arrPart(3,intP)%>
								</td>
								<td  class="td02" align="center"><%=formatnumber(arrPart(2,intP),0)%> ��</td>
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
					<td style="padding-top:15px;"><br>[�ݷ�����Ʈ]<br>
						<table width="100%" align="left" cellpadding="5" cellspacing="0" class="tbl01" border="0">
						<tr>
							<td align="center" class="td01">�ݷ�����Ʈ</td> 
						</tr>
						<tr>	
							<td>
								<%For intRA = 0 To UBound(arrReturn,2)%>
								<%=arrReturn(0,intRA)%>�� ���� �ݷ�&nbsp;<%=arrReturn(1,intRA)%>&nbsp;<%=formatdate(arrReturn(2,intRA),"0000-00-00")%><br>
								<%Next%>
							</td>
						</tr>
						</table>
					</td>		
				</tr>
				<%END IF%>	
				<%IF ipayrequeststate >=7 THEN '���� �� ����Ȯ�ε� �����϶��� ���뺸���ش�%>
				<tr>
				<td style="padding-top:15px;"><br>[�濵������ �����׸�]<br>
					<table width="100%" border="0" cellpadding="5" cellspacing="0"  class="tbl01">
						<tr  align="center">
							<td class="td01">����������</td> 
							<td class="td01">����(�Ա�)��</td>
							<td class="td01">�ش���(����)</td>
							<td class="td01">�������⿩��</td> 
						</tr>
						<tr  align="center">
							<td class="td02"><%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%>&nbsp;</td>
							<td class="td02"><%IF dpayrealdate <> "" THEN%><%=formatdate(dpayrealdate,"0000-00-00")%><%END IF%>&nbsp;</td>
							<td class="td02"><%IF syyyymm <> "" THEN%><%=year(syyyymm)%> ��  <%=month(syyyymm)%> ��<%END IF%>&nbsp;</td>
							<td class="td02"> <%IF blnTakeDoc THEN%>Y<%ELSE%>N<%END IF%>
						</td>
				</tr> 
				<tr bgcolor="#FFFFFF">
					<td colspan="5"  class="td01">*COMMENT<br>
					<%=pcomment%><Br>
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
<!-- #include virtual="/lib/db/dbclose.asp" --> 
