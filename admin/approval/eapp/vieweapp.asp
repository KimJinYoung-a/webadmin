<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ���� ���뺸�� - ������ view
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
Dim ireportidx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate,sreferid
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_nm,sedmsName,sedmscode,ionline,ioffline,iithinkso,ibnw,ifingers
Dim spartname ,ilastApprovalid,sjob_name,sscmLink,spart_name,susername, ipart_sn
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim sReferName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast, iNextAuthState,blnMod,iRectAuthState, iRectPartSn
Dim sEappName,blnPayEapp, mpayrequestprice
Dim hidAI_H, hidPS_H, AuthState_H, AuthName_H, AuthJonName_H, AuthConfirmTime_H, AuthSMS_H, hidAJ_H
Dim AuthID_L,AuthState_L,AuthName_L,AuthJobsn_L,AuthJobName_L,AuthConfirmTime_L,AuthSMS_L, sRectAuthType
Dim AuthID_F,AuthState_F,AuthName_F,AuthJobsn_F,AuthJobName_F,AuthConfirmTime_F,AuthSMS_F
Dim intNo, arrAID, arrATxt, arrAJSn 
Dim idepartment_id, sdepartmentnamefull
ireportidx =  requestCheckvar(Request("iridx"),10)
 intNo = 0
'���� �⺻ �� ���� ��������
set clseapp = new CEApproval
	clseapp.Freportidx = ireportidx  
	clseapp.fnGetEAppData
	
	iarap_cd		 			= clseapp.Farap_cd		
	sreportName      = clseapp.FreportName    
	mreportPrice     = clseapp.FreportPrice   
	iscmlinkno       = clseapp.Fscmlinkno     
	sbigo            = clseapp.Fbigo          
	tContents  			= clseapp.Freportcontents
	ireportstate     = clseapp.Freportstate   
	sreferid         = clseapp.Freferid       
	sadminid         = clseapp.Fadminid       
	dregdate         = clseapp.Fregdate       
	sarap_nm     		= clseapp.Farap_nm   
	sacc_cd		 			= clseapp.Facc_cd
	sacc_nm       	= clseapp.Facc_nm         
	sedmsName        = clseapp.FedmsName      
	sedmscode        = clseapp.Fedmscode 
	ilastApprovalid	 = clseapp.FlastApprovalid	
	sscmLink		 = clseapp.FscmLink  
	sjob_name		 = clseapp.Fjob_name   
  
 	'ipart_sn			=clseapp.Fpart_sn
	'spart_name		= clseapp.Fpart_name   
	idepartment_id	  = clseapp.Fdepartment_id
	sdepartmentnamefull= clseapp.Fdepartmentnamefull
	susername		= clseapp.Fusername
	blnPayEapp	= clseapp.FisPayEapp
	mpayrequestprice = clseapp.Fpayrequestprice
	
	arrAuth			= clseapp.fnGetAuthLineList
	arrComm			= clseapp.fnGetCommentList
	arrFile			= clseapp.fnGetAttachFileList 
	arrReturn		= clseapp.fnGetAuthLineReturnList  
	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing 
 
'refer�� ��������
set clsMem = new CTenByTenMember 
 	if sreferid <> "" then
 	clsMem.Fuserid	= sreferid
	arrRefer		= clsMem.fnGetInIDOutName
	end if
 set clsMem = nothing
  
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
</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4">
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">  
<tr>
	<td>
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0" ondblclick="javascript:jsPopView('vieweapp.asp?iridx=<%=ireportidx%>');"> 
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b><%=sedmsname%><%IF iarap_cd <> "0" THEN%>_<%=sarap_nm%><%END IF%></b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80">�����ڵ�</td>
					<td bgcolor="#FFFFFF" width="200"><%=sedmscode%></td> 
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="600"><!--������ ����Ʈ-->
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a">
						<tr align="center">
						<%	IF isArray(arrAuth) THEN
								For intA = 0 to UBound(arrAuth,2)
									IF arrAuth(4,intA)="A"  THEN ''����
										hidAI_H         = arrAuth(2,intA)
										hidPS_H         = arrAuth(9,intA)
										AuthState_H     = arrAuth(3,intA)
										AuthName_H      = arrAuth(7,intA)
										hidAJ_H         = arrAuth(8,intA)
										AuthJonName_H   = arrAuth(10,intA)
										AuthConfirmTime_H = arrAuth(6,intA)
										AuthSMS_H       = arrAuth(11,intA)
									ELSEIF arrAuth(4,intA)="L" THEN   '����������
										AuthID_L       	= arrAuth(2,intA)
										AuthState_L     = arrAuth(3,intA)
										AuthName_L      = arrAuth(7,intA)
										AuthJobsn_L     = arrAuth(8,intA)
										AuthJobName_L   = arrAuth(10,intA)
										AuthConfirmTime_L= arrAuth(6,intA)
								        AuthSMS_L       = arrAuth(11,intA) 
									ELSEIF arrAuth(4,intA)="F" THEN  '����������
										AuthID_F       	= arrAuth(2,intA) 
										AuthState_F     = arrAuth(3,intA)
										AuthName_F      = arrAuth(7,intA)
										AuthJobsn_F     = arrAuth(8,intA)
										AuthJobName_F   = arrAuth(10,intA)
										AuthConfirmTime_F= arrAuth(6,intA)
										AuthSMS_F       = arrAuth(11,intA)
								     ELSE  
								     	intNo = intNo  + 1
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
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20"><%=intNo%>�� ����</td></tr>
											<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>
											<tr><td align="Center"><%=arrAuth(7,intA)%>&nbsp;<%=arrAuth(10,intA)%></td></tr>
											<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">	&nbsp; </td></tr>
										</table>
										</div>
									</td>	
									<%								    	 			
								        end if 
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
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>"  height="20">����</td></tr>
											<% if (hidAI_H<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_H)%></td></tr>
											<tr><td align="Center"><%=AuthName_H%>&nbsp;<%=AuthJonName_H%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_H) THEN %><%=formatdate(AuthConfirmTime_H,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">&nbsp;</td></tr> 
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
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">����<%=chkIIF(AuthID_F="","����","����")%></td></tr>
										<% if (AuthID_L<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_L)%></td></tr>
											<tr><td align="Center"><%=AuthName_L%>&nbsp;<%=AuthJobName_L%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_L) THEN %><%=formatdate(AuthConfirmTime_L,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">&nbsp;</td></tr> 
										<% elseif (AuthID_F<>"") then %> 
											<tr><td align="Center"><%=fnGetAuthState(AuthState_F)%></td></tr>
											<tr><td align="Center"><%=AuthName_F%>&nbsp;<%=AuthJobName_F%></td></tr>
											<tr><td align="Center"><%IF not isNull(AuthConfirmTime_F) THEN %><%=formatdate(AuthConfirmTime_F,"0000-00-00")%><%ELSE%>&nbsp;<%END IF%></td></tr>
											<tr><td align="Center">&nbsp;</td></tr> 
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
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">����</td></tr>
											<tr><td align="Center">&nbsp;</td></tr>
											<tr><td align="Center"></td></tr>
											</table>
									    </div>
								    </td>
									<td valign="top"  width="180"  height="100%">
										<div id="dAP0">
										<table width="100%" cellpadding="5" cellspacing="0" class="a"  height="100%">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>" height="20">��������</td></tr>
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
					</td> 
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">��/�μ�</td>
					<td bgcolor="#FFFFFF"><%=sdepartmentnamefull%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=formatdate(dregdate,"0000-00-00")%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF"><input type="text" name="sRfN" value="<%=sReferName%>" size="30" style="border:0;" readonly></td> 
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="80" rowspan="2" valign="top" align="center">ǰ�ǳ���</td>
					<td>ǰ�Ǽ� IDX</td>
					<td>ǰ�Ǽ���</td>
					<td>ǰ�Ǳݾ�</td>
					<td>SCM ������ȣ</td>
					<td>���</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center"> 
					<td><%=ireportidx%></td>
					<td><%=sreportname%></td>
					<td><%IF mreportprice>0 THEN%><%=formatnumber(mreportprice,0)%><%END IF%></td>
					<td><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');"><%=iscmlinkno%></a></td>
					<td><%=sbigo%></td>
				</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" width="80" rowspan="3">����</td>
					<td bgcolor="#FFFFFF"  height="200">
					<%=tContents%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" height="100" valign="top">* COMMENT<br> 
						<%IF isArray(arrComm) THEN
							For intC = 0 To UBound(arrComm,2)
							%>
							 <div id="dC<%=intC%>"><%=arrComm(1,intC)%> &nbsp;<%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%> </div>
						<%	Next
						END IF%><br> 
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
						<a href="<%=arrFile(1,intF2)%>" target="_blank;"><%=arrFile(1,intF2)%></a><br> 
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
					<td rowspan="2" width="80">��������</td> 
					<td>�����׸�</td>
					<td>�����������</td> 
				</tr>
				<tr bgcolor="#FFFFFF"  align="center"> 
					<td><%=sarap_nm%>&nbsp;</td>
					<td><%=sacc_nm%></td>
				</tr>	
				</table>
			</td>
		</tr> 
		<%IF   blnPayEapp THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center" >
					<td width="80">�μ���<br>�ڱݱ���</td>
					<td align="left" bgcolor="#FFFFFF"> 
					<div id="divPM">
					<%dim arrPV, arrPT
					IF isArray(arrPart) THEN %>
						<table border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
						<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
							<td  width=200>�μ�</td>
							<td colspan=3 width=200>�ݾ�</td>
						</tr>
					<%	For intP = 0 To UBound(arrPart,2)
							if intp > 0 then 
								arrPV = arrPV&"," 
								arrPT = arrPT&","
							end if	
					%>   
					<tr>
						<td bgcolor="#eeeeee" > <%=arrPart(3,intP)%> </td>
						<td bgcolor="#FFFFFF" align="center"><%=formatnumber(arrPart(2,intP),0)%> ��</td>
						<td bgcolor="#FFFFFF" align="center"><%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=formatnumber((arrPart(2,intP)/mreportprice)*100)%><%END IF%> %</td>
					</tr> 
					<%	Next %>
					</table>
					<%END IF%>
					</div><br>
					<input type="hidden" name="iP" id="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" id="sP" value="<%=arrPT%>"> 
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
		</table>
	</td>
</tr>
<tr> 
	<td height="30" align="right"><%IF C_ADMIN_AUTH or C_PSMngPart THEN%><input type="button" value="����"  class="button" onClick="<%if blnPayEapp and mpayrequestprice > 0 THEN%>alert('����� ������û���� �ֽ��ϴ�. ���� �Ұ����մϴ�.');<%else%>jsEappDelete();<%end if%>" style="color:red;"><%END IF%></td>
</tr>
 </table>
 <Br><br>
 <form name="frm" method="post" action="proceapp.asp">
		<input type="hidden" name="hidM" value="D">
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
</form>
<script language="javascript">
	function jsEappDelete(){ 
		if(confirm("�����Ͻðڽ��ϱ�?")){
			document.frm.hidM.value = "DA";
			document.frm.submit();
		}
	}
</script>
<!-- ������ �� -->
</body>
</html>