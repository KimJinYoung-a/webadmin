<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ������û�� ����ó�� 
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
<%
Dim clsPay,clsMem,clseapp, clscomm,clsPM
Dim ireportidx,ipayrequestIdx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate 
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_nm,sedmsName,sedmscode 
Dim spartname ,ilastApprovalid,sscmLink, iauthposition,iBizNo
Dim chkPayRequest ,spayrequestTitle  
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate,pcomment,susername 
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrProc,arrPM, arrPart
Dim intA, intC, intF, intR, intRA, intP , iAuthCount, intPart 
Dim mSumPrice, mSumRealPrice  
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2), pmistate(2) ,pmstatecd(2)	 
Dim scust_cd, scustnm,chkID
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	
Dim ipaytype, sCurrencyType, sCurrencyPrice,serpLinkType
Dim sRectAuthId
ireportidx 		=  requestCheckvar(Request("iridx"),10)
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
	sRectAuthId =  session("ssBctId")
 chkID = 0
 
'���� �⺻ �� ���� ��������
set clsPay = new CPayRequest
	clsPay.Freportidx = ireportidx  
	clsPay.FpayrequestIdx = ipayrequestIdx  
	
	clsPay.fnGetPayRequestReceiveData 
	iarap_cd		 = clsPay.Farap_cd		
	sreportName      = clsPay.FreportName    
	mreportPrice     = clsPay.FreportPrice   
	iscmlinkno       = clsPay.Fscmlinkno     
	sbigo            = clsPay.Fbigo      
	ireportstate     = clsPay.Freportstate  
	sadminid         = clsPay.Fadminid        
	sarap_nm     		= clsPay.Farap_nm   
	sacc_cd		 			= clsPay.Facc_cd
	sacc_nm       	= clsPay.Facc_nm        
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
  spayrequestTitle	= clsPay.FpayRequestTitle 
  susername		 = clsPay.Fusername
  spartname		 = clsPay.Fpartname
  spayrequestTitle	= clsPay.FpayRequestTitle 							   
 	scust_cd						= clsPay.Fcust_cd
 	scustnm						=clsPay.Fcust_nm 
 	iBizNo						= clsPay.FBiz_no
 	ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	serpLinkType			= clsPay.FerpLinkType
	
	'//���������Ʈ
	arrProc			= clsPay.fnGetProcPayRequestList	
	
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
set clsPay = nothing 	
	 

'�������, �ڸ�Ʈ, ���� ����Ʈ ��������
set clseapp = new CEApproval	
	clseapp.Freportidx 		= ireportidx  
	clseapp.FpayrequestIdx = ipayrequestIdx 	
 	 
 	arrAuth			= clseapp.fnGetAuthLineList  '��������
	arrComm			= clseapp.fnGetCommentList	'�ڸ�Ʈ
	arrFile			= clseapp.fnGetAttachFileList  '÷������ 
	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing  

IF isArray(arrAuth) THEN '�������� 
	For intA = 0 To UBound(arrAuth,2)
		pmuserid(intA)  = arrAuth(2,intA) 
		pmusername(intA)= arrAuth(7,intA)
		pmjobname(intA) = arrAuth(10,intA)
		pmstate(intA)	= fnGetPayAuthState(arrAuth(3,intA), intA+1)
		pmdate(intA)	= arrAuth(6,intA)  
		pmstatecd(intA)= arrAuth(3,intA)
	Next 
END IF

 '������û���϶�(ipayrequeststate = 0 or 1)�� ������û��ó���� ���� ���� 
 '����ó�� �Ŀ��� ���� �Ұ���(�� �������� �Ŀ��� �繫ȸ��ó�� ����ڸ� ���氡��)
IF  ipayrequeststate = 1 THEN
'�繫ȸ���� ������û�� ó��������
Set clsPM	= new CPayManager
	arrPM	= clsPM.fnGetPayManager 
Set clsPM 	= nothing  

IF isArray(arrPM) THEN 
		For intP = 0 To ubound(arrPM,2) 
		IF arrPM(2,intP) = 1 THEN  '������� �����������϶�
			IF pmstatecd(0) = 0 THEN '���°� ���δ���϶� ����ó��
			pmuserid(0)  = arrPM(1,intP)	 	
			pmusername(0)= arrPM(3,intP)
			pmjobname(0) = arrPM(6,intP)
			pmstate(0) = fnGetPayAuthState(0,1)
			END IF
		ELSEIF  arrPM(2,intP)	 = 2 THEN   
			IF Cstr(trim(arrPM(1,intP)))	= Cstr(session("ssBctId")) THEN  
				pmuserid(1)  = arrPM(1,intP)	 	
				pmusername(1)= arrPM(3,intP)
				pmjobname(1) = arrPM(6,intP)
				pmstate(1) = fnGetPayAuthState(0,2) 
				chkID = 1
			END IF
		END IF
		Next  
		 
		IF 	chkID =0 THEN 
				pmuserid(1)  = arrPM(1, ubound(arrPM,2))	 	
				pmusername(1)= arrPM(3, ubound(arrPM,2))
				pmjobname(1) = arrPM(6, ubound(arrPM,2))
				pmstate(1) = fnGetPayAuthState(0,2)
		END IF
	END IF 
END IF
		
'�μ��� ��������
set clsMem = new CTenByTenMember
	clsMem.Fpart_sn = session("ssAdminPsn")
	clsMem.fnGetPartName 
 	spartname = clsMem.Fpart_name 
 set clsMem = nothing 
  
 
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
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0" ondblclick="javascript:jsPopView('viewpayrequest.asp?iridx=<%=ireportidx%>&ipridx=<%=ipayrequestIdx%>');"> 
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b>������û��(<%=sarap_nm%>)</b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">������û��idx</td>
					<td bgcolor="#FFFFFF" ><%=ipayrequestidx%></td>
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="300">
					<!------������ ����Ʈ------------------------------------------------------------>
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0">
						<tr>
							<td  valign="top" width="150"> 
								<table width="100%"  cellpadding="5" cellspacing="0" class="a"> 
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>����������</td> </tr>
									<tr align="Center"><td><%=pmstate(0)%></td> </tr>
									<tr align="Center"> <td><%=pmusername(0)%>&nbsp;<%=pmjobname(0)%></td> </tr>
									<tr align="Center"><td><%=pmdate(0)%></td></tr>
								</table> 
							</td> 
							<td  valign="top">
								<table width="100%"  cellpadding="5" cellspacing="0" class="a"> 
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>�繫ȸ����</td> </tr>
									<tr align="Center"><td><font color="gray"><%=pmstate(1)%></font></td> </tr>
									<tr align="Center"> <td><%=pmusername(1)%>&nbsp;<%=pmjobname(1)%></td> </tr>
									<tr align="Center"><td><%=pmdate(1)%></td></tr>
								</table>
							</td> 
						</tr> 
						</table>
						<!------//������ ����Ʈ------------------------------------------------------------>
					</td> 
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">��/�μ�</td>
					<td bgcolor="#FFFFFF"><%=spartname%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=dregdate%></td>
				</tr> 
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(ipayrequeststate)%></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="80" rowspan="2"  align="center" >����ǰ�Ǽ�</td>
					<td>ǰ�Ǽ� IDX</td>
					<td>ǰ�Ǽ���</td>
					<td>ǰ�Ǳݾ�(��)</td>  
					<td>���</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center"> 
					<td><%=ireportidx%></td>
					<td><%=sreportname%></td>
					<td><%=formatnumber(mreportprice,0)%></td>  
					<td><a href="javascript:jsPopView('/admin/approval/eapp/confirmeapp.asp?iridx=<%=ireportidx%>')">�󼼺���>></a></td>
				</tr>
				</table>	
			</td>
		</tr>
		<%IF isArray(arrProc) THEN%>
		<tr>
			<td> 
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr align="center">
					<td  bgcolor="<%= adminColor("tabletop") %>" width="80" rowspan="<%=UBound(arrProc,2)+2%>">���������</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">������û�� IDX</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">����(�Ա�)��</td>
					<td  bgcolor="<%= adminColor("tabletop") %>">�����ݾ�(��)</td> 
					<td  bgcolor="<%= adminColor("tabletop") %>">��������</td> 
					<td  bgcolor="<%= adminColor("tabletop") %>">���</td>
				</tr>
				<%For intP = 0 To UBound(arrProc,2)%>
				<tr align="center">	
					<td bgcolor="#FFFFFF"><%=arrProc(0,intP)%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(3,intP) <> "" THEN%><%=formatdate(arrProc(3,intP),"0000-00-00")%><%END IF%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(2,intP) <> "" THEN%><%=formatnumber(arrProc(2,intP),0)%><%END IF%></td> 
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(arrProc(4,intP))%></td> 
					<td bgcolor="#FFFFFF"><a href="javascript:jsPopView('confirmpayrequest.asp?iridx=<%=ireportidx%>&ipridx=<%=arrProc(0,intP)%>&ias=<%=arrProc(5,intP)%>')">�󼼺��� >></a></td>
				</tr>
				<%Next%>
				</table>
			</td>
		</tr>
		<%END IF%>	
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="80" rowspan="5">������û����</td>
					<td>������û��</td>
					<td>������û�ݾ�(��)</td>
					<td>�������</td> 
					<td width="250" >���</td>  
				</tr>
				<tr align="center"  bgcolor="#FFFFFF">	
					<td><%IF dpayrequestdate <> "" THEN%><%=formatdate(dpayrequestdate,"0000-00-00")%><%END IF%></td>
					<td><%IF mpayrequestprice <> "" THEN%><%=formatnumber(mpayrequestprice,0)%><%END IF%></td>
	 				<td><%=fnGetPayType(ipaytype)%></td>
					<td>	  
							<span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							��ȭ�ݾ�: <%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%> 
							</span>
					</td>
				</tr> 	
				<tr>	
					<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڱݿ뵵</td>  
					<td colspan="3" bgcolor="#FFFFFF"><%=sPayrequesttitle%></td>
				</tr>   
				<tr  align="center"  bgcolor="<%= adminColor("tabletop") %>">	
					<td>�ŷ�ó</td>
					<td>�����</td> 
					<td>�����ָ�</td>
					<td>���¹�ȣ</td>
				</tr> 
				<tr align="center"  bgcolor="#FFFFFF">	
					<td><%=scustnm%></td>
					<td><%=iinBank%></td> 
					<td><%IF saccountholder <> "" THEN%><%=saccountholder%><%END IF%></td>
					<td><%IF saccountno <> "" THEN%><%=saccountno%><%END IF%></td>
				</tr> 
				</table>
			</td>
		</tr>
	<tr>
			<td> 
				<table width="100%" align="left" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td rowspan="8" width="80"  align="center" style="padding:5px;">��������</td>   
					<td colspan="2" bgcolor="#FFFFFF"  style="padding:5px;"> 
						 	<!-- #include virtual="/admin/approval/eapp/incDocInfo.asp" -->   
						</td>
				</tr>	
				<tr>
					<td width="90" bgcolor="<%= adminColor("tabletop") %>" align="center"  style="padding:5px;">��������</td>
					<td bgcolor="#FFFFFF"  style="padding:5px;" width="620"> 
						<% 
					 	Dim strDoc
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
					 	%> <%=strDoc%> 
					</td>
				</tr> 
				<tr>
					<td bgcolor="#FFFFFF" colspan="2">
						<div id="dView1" style="display:<%IF not (ipaydockind = "1" or ipaydockind ="2") or setaxkey = "" THEN%>none<%END IF%>;">
						<table border="0" cellpadding="5" cellspacing="0" class="a" width="100%"> 
						<tr>
							<td  width="90" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> �������� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> 
									<% 
								Dim strVat
									IF sVatKind ="0" THEN 
										strVat = "����(�ΰ��� 10%) "
									ELSEIF sVatKind ="2" THEN 
										strVat = "�鼼"
									ELSEIF sVatKind ="3" THEN 
										strVat = "����" 
									END IF
								%><%=strVat%>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ������ </td>
							<td bgcolor="#FFFFFF"  colspan="3" style="border-bottom:1px solid #BABABA;"><%=dissuedate%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ǰ�� </td>
							<td colspan="5" bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"><%=sItemName%> </td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> �ѱݾ� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> <%=formatnumber(mTotPrice,0)%> ��</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ���ް� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=formatnumber(mSupplyPrice,0)%> ��</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> �ΰ��� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"><%=formatnumber(mVatPrice,0)%> ��</td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> ����û���ι�ȣ </td>
								<td bgcolor="#FFFFFF" style="border-right:1px solid #BABABA;"><%=setaxkey%></td>
								<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> ��� </td>
								<td bgcolor="#FFFFFF"  colspan="3"><%=sDocBigo%></td>
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
					<td rowspan="2" width="80">÷�μ���</td>
					<td>÷������</td>
					<td>���ø�ũ</td>
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
						<a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a> 
						</div>
						<%Next
						END IF
						%>
						</div><Br>
					</td>
					<td> 
						<% iCount = 0
						IF isArray(arrFile) THEN
						For intF2 = intF To UBound(arrFile,2)%>
						  <a href="javascript:jsFileLink('<%=arrFile(1,intF2)%>')"><%=arrFile(1,intF2)%></a><br>
						<% iCount = iCount + 1
						Next
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
					<td rowspan="3" width="80">��������</td> 
							<td>�����׸�</td>
							<td>�����������</td> 
						</tr>
						<tr bgcolor="#FFFFFF"  align="center"> 
							<td><input type="text" name="sANM" value="<%=sarap_nm%>"  style="border:0" readonly ><input type="hidden" name="iaidx" value="<%=iarap_cd%>"></td>
							<td><input type="text" name="sACCNM" value="<%=sacc_nm%>"  style="border:0" readonly ><input type="hidden" name="sACC" value="<%=sacc_cd%>"></td>
						</tr>	 
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center" >
					<td width="60" rowspan="2" style="padding:5px">�μ���<br>�ڱݱ���</td>
					<td width="300" style="padding:5px" > �μ�</td> 
					<td width="205" style="padding:5px"> �ݾ�</td>
					<td width="205" style="padding:5px"> %</td>
				</tr>
				<tr>
					<td colspan="3" bgcolor="#FFFFFF" valign="top">	 
					<div id="divPM">
					<%dim arrPV, arrPT
					IF isArray(arrPart) THEN %>
						<table border="0" cellpadding="3" cellspacing="0" class="a" width="760">  
					<%	For intPart = 0 To UBound(arrPart,2)
							if intPart > 0 then 
								arrPV = arrPV&"," 
								arrPT = arrPT&","
							end if	
							
							arrPV = arrPV&arrPart(1,intPart)
							arrPT = arrPT&arrPart(3,intPart) 
					%>  
					<tr>
						<td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(4,intPart)%></td>
						<td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(3,intPart)%> </td>
						<td width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><%=formatnumber(arrPart(2,intPart),0)%> ��</td>
						<td width="200" style="border-bottom:1px solid #BABABA;" align="center"><%IF mpayrequestprice <> 0 AND arrPart(2,intPart)<> 0 THEN%><%=formatnumber((arrPart(2,intPart)/mpayrequestprice)*100,1)%><%END IF%>%</td>
					</tr> 
					<%	Next %>
					</table>
					<%END IF%>
					</div><br>
					<input type="hidden" name="iP" id="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" id="sP" value="<%=arrPT%>">
					<input type="hidden" name="mP" id="mP" value=""> 
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
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">�ݷ�����Ʈ</td> 
					<td bgcolor="#FFFFFF">
						<%For intRA = 0 To UBound(arrReturn,2)%>
						<%=arrReturn(0,intRA)%>�� ���� �ݷ�&nbsp;<%=arrReturn(1,intRA)%>&nbsp;<%=formatdate(arrReturn(2,intRA),"0000-00-00")%><br>
						<%Next%>
					</td>
				</tr>
				</table>
			</td>		
		</tr>
		<%END IF%>	
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="3" width="80">�濵������<br>�����׸�</td>
					<td>����������</td> 
					<td>����(�Ա�)��</td>
					<td>�ش���(����)</td>
					<td>�������⿩��</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
					<td><input type="text" name="dPD" size="10" value="<%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%>" size="10" style="border:0" readonly></td>
					<td><%IF dpayrealdate <> "" THEN%><%=formatdate(dpayrealdate,"0000-00-00")%><%END IF%></td>
					<td><%IF syyyymm <> "" THEN%><%=year(syyyymm)%> �� 
						<%=month(syyyymm)%> ��<%END IF%>
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
<tr>
	<td style="padding-top:20px;"> 
		<!-- #include virtual="/admin/approval/eapp/incComment.asp" --> 
	</td>
</tr>
<tr>
	<td height="50">&nbsp;</td>
</tr>
</table> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
