<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ���ڰ��� ����
' History : 2011.03.14 ������  ����
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
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_nm,sacc_use_cd,sedmsName,sedmscode ,sscmsubmitLink
Dim ipart_sn,ilastApprovalid,sjob_name,sscmLink,spart_name, susername 
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim sReferName,sEappName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast,iNextAuthState,blnMod  
Dim blnpayEapp,		mpayrequestprice	
Dim iLastposition				
ireportidx =  requestCheckvar(Request("iridx"),10) 
iLastposition =  requestCheckvar(Request("iLP"),10) 
'���� �⺻ �� ���� ��������
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
	sarap_nm         = clseapp.Farap_nm        	
	sacc_cd          = clseapp.Facc_cd          	
	sacc_nm          = clseapp.Facc_nm   
	sacc_use_cd			 = clseapp.Facc_use_cd       	
	sedmsName        = clseapp.FedmsName         
	sedmscode        = clseapp.Fedmscode         
	ilastApprovalid  = clseapp.FlastApprovalid   
	sscmLink				 = clseapp.FscmLink					
	sscmsubmitLink	 = clseapp.FscmsubmitLink		
	sjob_name			   = clseapp.Fjob_name					
	ipart_sn				 = clseapp.Fpart_sn					
	spart_name			 = clseapp.Fpart_name				
	susername				 = clseapp.Fusername				
	blnpayEapp			 = clseapp.FispayEapp
	mpayrequestprice = clseapp.Fpayrequestprice	
   
	arrAuth			= clseapp.fnGetAuthLineList
	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList 
	arrReturn		= clseapp.fnGetAuthLineReturnList 
 	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing 
 
'�μ��� ��������
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

'���縮��Ʈ-----------------------------------------
blnMod = 0  		'���� ���� ���ɿ���
blnLast = 0 		'�������翩��
iRectPosition = 0	'���������ġ 
iNextPosition = iLastPosition '����������ġ
sNextAuthId = ""	'���������ھ��̵�
iNextAuthState = 0	'�����������
sRectAuthId = sadminid	 '������� ���̵� = ��������

IF isArray(arrAuth) THEN  
	 	sNextAuthId	 = arrAuth(2,(iNextPosition-1))
	 	iNextAuthState = arrAuth(3,(iNextPosition-1))   
END IF
 
'--------------------------------------------------	   

 '���� �������ɿ���
 IF  sRectAuthId = session("ssBctId") THEN
 	blnMod = 1
 END IF	
  
 '���� ����Ʈ--------------------------------------
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
<script language="javascript">
	function jsEappAddSubmit(){
		if(jsChkBlank(document.frm.hidAI.value) ){
				alert("�����ڸ� ������ּ���");
				return;
			}
			
			if(jsChkBlank(document.frm.mRP.value) ){
				alert("ǰ�Ǳݾ��� �Է����ּ���");
				return;
			}
			
			if(confirm("�������Ͻðڽ��ϱ�?")){
			document.frm.hidRS.value = 1;
			document.frm.submit();
		}
	}
</script>
</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4">
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">  
<form name="frmCD" method="post" action="proceapp.asp">
 <input type="hidden" name="hidM" value="CD">
 <input type="hidden" name="iCidx" value="">
 <input type="hidden" name="iRidx" value="<%=ireportidx%>"> 
 <input type="hidden" name="ipridx" value="0">  
 <input type="hidden" name="hidRU" value="modeapp.asp?iRS=<%=ireportstate%>&iridx=<%=ireportidx%>">
 </form>
<tr>
	<td>
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0">
		<form name="frm" method="post" action="proceapp.asp">  
		<input type="hidden" name="hidM" value="A">
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
		<input type="hidden" name="hidRS" value="<%=ireportstate%>">
		<input type="hidden" name="iAIdx" value="<%=iarap_cd%>">  
		<input type="hidden" name="iAP" value="<%=iLastPosition+1%>"> 
		<input type="hidden" name="hidAid" value="<%=sadminid%>">
		<input type="hidden" name="hidRfI" id="hidRfI" value="<%=sreferId%>">
		<input type="hidden" name="hidAI" id="hidAI" value="<%=sNextAuthId%>">
		<input type="hidden" name="hidPS" value="<%=session("ssAdminPsn")%>">
		<input type="hidden" name="iLAID" value="<%=ilastApprovalid%>">   
		<input type="hidden" name="hidUN" value="<%=susername%>"> 
		<input type="hidden" name="hidAN" value=""> 
		<Tr>
			<td align="right" style="border-bottom:1px dashed #cccccc;"><input type="button" value="����Ʈ" class="button" onClick="jsPopModPrint(<%=ireportidx%>);"></td>
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
					<td bgcolor="<%= adminColor("tabletop") %>" width="60" align="center">�����ڵ�</td>
					<td bgcolor="#FFFFFF"><%=sedmscode%></td>
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="500"><!--������ ����Ʈ-->
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a">
						<tr align="center">
							<td valign="top">
								<table width="100%" cellpadding="5" cellspacing="0" class="a" width="100%">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>	
								</table>
							</td> 
							<td valign="top">
								<table width="100%" cellpadding="5" cellspacing="0" class="a" width="100%">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>	
								</table>
							</td> 
							<%IF isArray(arrAuth) THEN
								For intA = 0 To UBound(arrAuth,2) 
									IF arrAuth(4,intA) THEN
										blnLast = 1  
							%>
								<td valign="top" width="150">
								<div id="dAP0">
								<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">����������</td></tr>  
									<%IF  CInt(iNextPosition) = Cint(arrAuth(1,intA)) and blnMod = 1  THEN  %> 
									<tr><td align="Center">&nbsp;</td></tr>
									<tr><td align="Center"><input type="text" name="sALN" id="sALN" value="<%=arrAuth(7,intA)&" "&arrAuth(10,intA)%>" style="border:0;text-align:center;" readonly size="20"><input type="hidden" name="hidAJ" id="hidAJ" value="<%=arrAuth(10,intA)%>"></td></tr>
									<tr><td align="Center">&nbsp;</td></tr>
									<tr><td align="Center"><input type="button" class="button" value="������ ���" onClick="jsRegID(1);document.frm.blnL.value=1;"><br>
										<input type="checkbox" value="1" name="chkSms" <%IF arrAuth(11,intA) THEN%> checked<%END IF%>> SMS����</td></tr>
									<%ELSE%>
									<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>	
									<tr><td align="Center"><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td></tr>	
									<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td></tr>	
									<tr><td align="Center">&nbsp;</td></tr>	
									<%END IF%>	 
									</table>
								</div>
								</td>
							<%	 	Exit For
								END IF			 
								%> 
							<% Next  
							END IF%> 
						</tr>  
						<input type="hidden" name="blnL" value="<%=blnLast%>">		
						</table>
					</td> 
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">��/�μ�</td>
					<td bgcolor="#FFFFFF"><%=spart_name%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%IF ireportstate > 0 THEN%><%=formatdate(dregdate,"0000-00-00")%><%ELSE%><%=date()%><%END IF%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sRfN" id="sRfN" value="<%=sReferName%>" size="20" style="border:0;" readonly></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="60" rowspan="2" valign="top" align="center">ǰ�ǳ���</td>
					<td>IDX</td>
					<td>ǰ�Ǽ���</td>
					<td>ǰ�Ǳݾ�</td>
					<td>SCM ������ȣ</td>
					<td>���</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center"> 
					<td  align="Center"><%=ireportidx%></td>
					<td><input type="text" name="sRN" size="40" maxlength="60" value="<%=sreportname%>"  style="border:0" readonly ></td>
					<td ><input type="text" name="mRP" size="15" maxlength="15" style="text-align:right;" value="" <%IF blnMod = 0   THEN%>style="border:0" readonly<%END IF%>></td>
					<td  align="Center"><input type="hidden" name="iSL" value="<%=iscmlinkno%>" ><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');"><%=iscmlinkno%></a></td>
					<td><input type="text" name="sB" size="20" value="<%=sbigo%>" style="border:0" readonly></td>
				</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA> 
			 <tr>
					<td bgcolor="#FFFFFF" height="100" valign="top"> *COMMENT<br>
						<%IF isArray(arrComm) THEN  
							For intC = 0 To UBound(arrComm,2)
							%>
							 <div id="dC<%=intC%>"><%=arrComm(1,intC)%> &nbsp;<%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%>
							 &nbsp;<%IF  sRectAuthId = arrComm(2,intC) THEN%><input type="button" class="button" value="x" onClick="jsCommDel('<%=arrComm(0,intC)%>');"><%END IF%></div>
						<%	Next
						END IF%><br>
						<%IF blnMod = 1 THEN %>
						<textarea id="tCmt" name="tCmt" rows="3" cols="100" ></textarea>   
						<%END IF%>   
					</td>
				</tr>
				</table>
			</td>
		</tr> 
		<%IF blnMod = 1 THEN%> 
		<tr>
			<td align="center">
				<table border="0" cellpadding="5" cellspacing="0" width="100%">
				<tr> 
					<td align="right"> <input type="button" value="������" class="button" onClick="jsEappAddSubmit(1);"></td> 
				</tr>
				</table>
			</td>
		</tr> 
		<%END IF%>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td height="50"></td>
</tr>
</table>  
</body>
</html>
