<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ���� ���ڰ��� ����ó��
' History : 2011.03.14 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim clseapp,clsMem
Dim ireportidx,iarap_cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate,sreferid
Dim sadminid,dregdate,saccountName,scomm_name,scomm_desc,ierpCode,sedmsName,sedmscode 
Dim spartname ,ilastApprovalid,sjob_name,sscmLink,spart_name,susername, ipart_sn
Dim tContents
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn, arrPart
Dim intA, intC, intF, intR, intRA, intP
Dim sReferName
Dim sRectAuthId, iRectPosition,iNextPosition, sNextAuthId, blnLast, iNextAuthState,blnMod,iRectAuthState, iRectPartSn, iRectisLast
Dim sscmsubmitlink

ireportidx =  requestCheckvar(Request("iridx"),10)
 
'���� �⺻ �� ���� ��������
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
 
 
'���縮��Ʈ-----------------------------------------
blnMod = 0  		'���� ���� ���ɿ���
blnLast = 0 		'���������� ��Ͽ���
iRectisLast = 0		'���� ���������ڿ���
iRectPosition = 0	'���������ġ 
iNextPosition = 0	'����������ġ
sNextAuthId = ""	'���������ھ��̵�
iNextAuthState = 0	'�����������
iRectAuthState = 0
 iRectPartSn =ipart_sn
IF isArray(arrAuth) THEN
	For intA = 0 to UBound(arrAuth,2)
		IF arrAuth(2,intA) = session("ssBctId") THEN
			sRectAuthId = arrAuth(2,intA)	'���� ������̵�
			iRectPosition= arrAuth(1,intA)	'���� ���� ��ġ
			iRectAuthState	= arrAuth(3,intA) '���� ���� ���� 
			iRectPartSn	= arrAuth(9,intA)
			IF arrAuth(4,intA) THEN iRectisLast = 1
			IF iRectisLast <> 1 THEN '������ġ�� ���������ϋ� ������ġ ����.
				IF intA+1 <= UBound(arrAuth,2) THEN
				iNextPosition = arrAuth(1,intA+1)
				sNextAuthId	  = arrAuth(2,intA+1)
				iNextAuthState = arrAuth(3,intA+1) 
				iRectPartSn	= arrAuth(9,intA+1)
				ELSE
				iNextPosition = iRectPosition+1 
				END IF	
			END IF
		END IF	
	Next   
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
<table width="100%" cellpadding="5" cellspacing="1" class="a">  
<tr>
	<td>
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0"> 
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b><%=sedmsname%><%IF iarap_cd > 0 THEN%>(<%=saccountname%>)<%END IF%></b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80">�����ڵ�</td>
					<td bgcolor="#FFFFFF" width="200"><%=sedmscode%></td> 
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="600"><!--������ ����Ʈ-->
						<table width="100%" align="left" cellpadding="0" cellspacing="0" class="a" border="1">
						<tr align="center">
							<%IF isArray(arrAuth) THEN
								For intA = 0 To UBound(arrAuth,2)
									IF arrAuth(4,intA) THEN
										blnLast = 1 
										Exit For
									END IF	 
								%>
							<td valign="top" width="150">
								<div id="dAP<%=intA+1%>">
									<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>"><%=intA+1%>�� ����</td></tr>
									<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>
									<tr><td align="Center"><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td></tr>
									<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td></tr>
									<tr><td align="Center">&nbsp;</td></tr>
									</table> 
								</div>
							</td>
							<% Next 
							 ELSE		
							%>
							<td valign="top" width="150">
							<div id="dAP1">
									<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">1�� ����</td></tr>
									<tr><td align="Center"><input type="text" name="sASD" style="border:0;" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sALN" value="" style="border:0;" readonly size="20"><input type="hidden" name="hidAJ" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sADD" value="" style="border:0;text-align:center;"></td></tr>
									<tr><td align="Center"></td></tr>
									</table> 
							</div>
							</td>
							<%END IF%>
							<%IF blnLast = 0 and blnMod = 1 THEN%>
							<td valign="top" width="150">
							<div id="dAP<%=iNextPosition%>">
									<table width="100%"  cellpadding="5" cellspacing="0" class="a">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>"><%=iNextPosition%>�� ����</td></tr>
									<tr><td align="Center"><input type="text" name="sASD" style="border:0;" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sALN" value="" style="border:0;" readonly size="20"><input type="hidden" name="hidAJ" value=""></td></tr>
									<tr><td align="Center"><input type="text" name="sADD" value="" style="border:0;text-align:center;"></td></tr>
									<tr><td align="Center"></td></tr>
									</table> 
							</div>
							</td>	
							<%END IF%>	
							<td valign="top" >
								<table cellpadding="5" cellspacing="0" class="a" width="100%">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
									<tr><td align="Center">&nbsp;</td></tr>	
								</table>
							</td>
							<td valign="top"  width="150">
								<div id="dAP0">
									<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0">
									<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">����������</td></tr>
								<%IF blnLast=1 THEN%>
									<tr><td align="Center"><%=fnGetAuthState(arrAuth(3,intA))%></td></tr>	
									<tr><td align="Center"><%=arrAuth(7,intA)%> <%=arrAuth(10,intA)%></td></tr>	
									<tr><td align="Center"><%IF not isNull(arrAuth(6,intA)) THEN %><%=formatdate(arrAuth(6,intA),"0000-00-00")%><%END IF%></td></tr>	
									<tr><td align="Center"></td></tr>		
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
						</table>
					</td> 
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">��/�μ�</td>
					<td bgcolor="#FFFFFF"><%=spart_name%></td>
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
					<td bgcolor="#FFFFFF"><%=sReferName%></td> 
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
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
					<td><%=mreportprice%></td>
					<td><%=iscmlinkno%></td>
					<td><%=sbigo%></td>
				</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" width="80" rowspan="3"  align="center">����</td>
					<td bgcolor="#FFFFFF"  height="200" valign="top">
					<%=tContents%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#FFFFFF" height="100" valign="top">* COMMENT<br> 
						<%IF isArray(arrComm) THEN
							For intC = 0 To UBound(arrComm,2)
							%>
							 <div id="dC<%=intC%>"><%=arrComm(1,intC)%> &nbsp;<%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%>
							 &nbsp;<%IF Cstr(arrComm(2,intC)) = Cstr(session("ssBctId")) THEN%><input type="button" class="button" value="x" onClick="jsCommDel('<%=arrComm(0,intC)%>');"><%END IF%></div>
						<%	Next
						END IF%><br>
						<%IF blnMod = 1 THEN%>
						<textarea id="tCmt" name="tCmt" rows="3" cols="100"></textarea>
						<%END IF%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
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
		<%IF iarap_cd > 0 THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2" width="80">�����׸�</td>
					<td>�����ڵ�</td>
					<td>��������</td>
					<td>���</td>
				</tr>
				<tr bgcolor="#FFFFFF"  align="center">
					<td><%=ierpCode%></td>
					<td><%=scomm_name%></td>
					<td><%=scomm_desc%></td>
				</tr>	
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
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
						<td bgcolor="#eeeeee" >
							<%IF arrPart(4,intP) > 0 THEN 
								arrPV = arrPV&arrPart(4,intP)&"^" 
								arrPT = arrPT&arrPart(5,intP)&" > "
							%>
								<%=arrPart(5,intP)%>  > 
							<%END IF%>
							<%IF arrPart(6,intP) > 0 THEN
									arrPV = arrPV&arrPart(6,intP)&"^"
									arrPT = arrPT&arrPart(7,intP)&" > "
							%>
								<%=arrPart(7,intP)%> > 
							<%END IF 
								arrPV = arrPV&arrPart(1,intP)
								arrPT = arrPT&arrPart(3,intP) 
							%>
							<%=arrPart(3,intP)%>
						</td>
						<td bgcolor="#FFFFFF" align="center"><%=arrPart(2,intP)%> ��</td>
						<td bgcolor="#FFFFFF" align="center"><%IF mreportprice <> 0 AND arrPart(2,intP)<> 0 THEN%><%=(arrPart(2,intP)/mreportprice)*100%><%END IF%> %</td>
					</tr> 
					<%	Next %>
					</table>
					<%END IF%>
					</div><br>
					<input type="hidden" name="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" value="<%=arrPT%>"> 
					</td>
				</tr> 
				</table>
			</td>
		</tr>
		<%END IF%>
		<%IF isArray(arrReturn) THEN%>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="0" class="a" border="1" bgcolor=#BABABA>
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
</table> 
<script language="javascript">
<!--
 	//document.body.onload=function(){window.print();} 
//-->
</script> 
</body>
</html>
