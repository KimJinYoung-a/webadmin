<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 
' History : 2011.03.14 ������  ����
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
Dim iarapcd, iedmsIdx
Dim  sedmsname,sedmscode,sarap_cd,sarap_nm,sacc_cd,sacc_nm,sacc_use_cd,sACC_GRP_CD
Dim sEappName, mReportPrice
Dim spartname ,slastApprovalid,sjob_name,sscmLink,iscmlinkno
Dim tContents,blnPayEapp
Dim sCurrencyPrice,ipaytype,sCurrencyType

iarapcd =  requestCheckvar(Request("iAidx"),13)
if iarapcd = "" then iarapcd = 0 
iedmsIdx	=  requestCheckvar(Request("ieidx"),10)
 
'sacc_nm =  requestCheckvar(Request("sAN"),30)
tContents  = ReplaceRequestSpecialChar(Request("tC")) 
iscmlinkno		=  requestCheckvar(Request("iSL"),10) 
mReportPrice=  requestCheckvar(Request("mRP"),20)

'���� �⺻ �� ���� ��������
set clseapp = new CEApproval
	clseapp.Farap_cd  = iarapcd
	clseapp.Fedmsidx = iedmsIdx
	clseapp.fnGetEAppForm
	
	iedmsIdx         = clseapp.FedmsIdx        
	sedmsname        = clseapp.Fedmsname       
	sedmscode				= clseapp.Fedmscode				
	slastApprovalid  = clseapp.FlastApprovalid 
	sscmLink   			= clseapp.FscmLink   			
	sjob_name				= clseapp.Fjob_name				
	sarap_cd 				= clseapp.Farap_cd 				
	sarap_nm    		  = clseapp.Farap_nm    		
	sacc_cd    			= clseapp.Facc_cd    			
	sacc_nm				  = clseapp.Facc_nm		
	sacc_use_cd			= clseapp.Facc_use_cd  
	blnPayEapp			= clseapp.FisPayEapp
	sACC_GRP_CD			= clseapp.FACC_GRP_CD
	IF tContents ="" THEN
	tContents				= clseapp.FedmsForm
	END IF
set clseapp = nothing 

'�μ��� ��������
set clsMem = new CTenByTenMember
	clsMem.Fpart_sn = session("ssAdminPsn")
	clsMem.fnGetPartName
 
 	spartname = clsMem.Fpart_name
 set clsMem = nothing
  
 IF iarapcd > 0 THEN
 	sEappName = sedmsname&"_"&sarap_nm
 ELSE
 	sEappName = sedmsname 
 END IF	
 
%>

<%
 IF sscmLink <> "" and iscmlinkno ="" THEN
 	Call Alert_return ("���԰�ο� ������ �߻��Ͽ����ϴ�.") 
response.end
END IF 
%>

<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script language="javascript" src="eapp.js"></script>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#F4F4F4">
<table width="840" height="100%" cellpadding="0" cellspacing="0"  border="0" align="center">
<tr> 
	<td valign="top"> 
		<table width="100%" cellpadding="1" cellspacing="0" class="a"> 
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
				<form name="frm" method="post" action="proceapp.asp">   
				<input type="hidden" name="hidM" value="I">
				<input type="hidden" name="hidRS" value="0">
				<input type="hidden" name="iAIdx" value="<%=iarapcd%>">  
				<input type="hidden" name="ieIdx" value="<%=iedmsIdx%>">  
				<input type="hidden" name="iAP" value="1">
				<input type="hidden" name="hidAid" value="<%=session("ssBctId")%>">
				<input type="hidden" name="hidAI" id="hidAI" value="">
				<input type="hidden" name="hidRfI" id="hidRfI" value="">
				<input type="hidden" name="hidPS" value="<%=session("ssAdminPsn")%>">
				<input type="hidden" name="iLAID" value="<%=slastApprovalid%>">
				<input type="hidden" name="blnL" value="0">
				<input type="hidden" name="hidUN" value="<%=session("ssBctCname")%>">
				<input type="hidden" name="hidAN" value="">
				<input type="hidden" name="iRM" value="M010">
					<input type="hidden" name="hidPE" value="<%=blnPayEapp%>">
				<tr>
					<td>
						<table width="100%" cellpadding="5" cellspacing="1" class="a" >
						<tr>
							<td class="verdana-large"><b><%=sEappName%> </b></td>
							<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
						</tr>
						</table>
					</td>
				</tr>		
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tR>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="60">�����ڵ�</td>
							<td bgcolor="#FFFFFF" width="150"><%=sedmscode%></td>
							<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="500">
								<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a">
								<tr align="center">
									<td valign="top" width="150">
										<div id="dAP1">
											<table width="100%"  cellpadding="5" cellspacing="0" class="a">
											<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">1�� ����</td></tr>
											<tr><td align="Center"><input type="text" name="sASD" style="border:0;" value=""></td></tr>
											<tr><td align="Center"><input type="text" name="sALN" id="sALN" value="" style="border:0;" readonly size="20"><input type="hidden" name="hidAJ" value=""></td></tr>
											<tr><td align="Center"><input type="text" name="sADD" value="" style="border:0;text-align:center;"></td></tr>
											<tr><td align="Center"><input type="button" class="button" value="�����ڵ��" onClick="jsRegID(1);"></td></tr>
											</table> 
										</div>
										</td>
									<td valign="top">
										<table width="100%" cellpadding="5" cellspacing="0" class="a">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
										<tr><td align="Center">&nbsp;</td></tr>	
										</table>
									</td>
									<td valign="top"  width="150">
										<div id="dAP0">
										<table width="100%" cellpadding="5" cellspacing="0" class="a">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">����������</td></tr>
										<tr><td align="Center">&nbsp;</td></tr>	
										<tr><td align="Center"><%=sjob_name%></td></tr>
										</table>
										</div>
									</td> 
								</tr> 
								</table>
							</td> 
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >��/�μ�</td>
							<td bgcolor="#FFFFFF"><%=spartname%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >�ۼ���</td>
							<td bgcolor="#FFFFFF"><%= session("ssBctCname")%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >�ۼ���</td>
							<td bgcolor="#FFFFFF"><%=date()%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >����</td>
							<td bgcolor="#FFFFFF"><input type="button" class="button" value="�������" onClick="jsRegID(2);">	<input type="text" name="sRfN" id="sRfN" value="" size="30" style="border:0;" readonly></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td width="60" rowspan="4"  align="center">ǰ�ǳ���</td>
							<td>IDX</td>
							<td>ǰ�Ǽ���</td>
							<td>ǰ�Ǳݾ�</td>
							<td>����Ÿ��</td> 
							<td>SCM<br>������ȣ</td> 
						</tr>
						<tr bgcolor="#FFFFFF" align="center"> 
							<td></td>
							<td><input type="text" name="sRN" size="40" maxlength="50" value="<%=sEappName%>"></td> 
							<td><input type="text" name="mRP" size="15" maxlength="20" style="text-align:right;" value="<%=mReportPrice%>" <%IF not blnPayEapp THEN%>disabled class="text_ro"<%END IF%> onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" onblur="jsIsHundred();"></td>
							<td>
								<select name="selPT" onChange="jsChFC();" class="select" <%IF not blnPayEapp THEN%>disabled<%END IF%>>
								<%sboptPayType ipaytype%> 
							</select>
							<div  id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> <%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"> </div>
							</td>
							<td><input type="hidden" name="iSL" value="<%=iscmlinkno%>"><%=iscmlinkno%> <%IF sscmLink <> "" THEN%>><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');">>�󼼺���</a><%END IF%></td>   
						</tr>
						</table>	
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" width="60" rowspan="3">����</td>
							<td bgcolor="#FFFFFF" height="100"> 
							<!--#Include Virtual = "/admin/approval/eapp/incEditor.asp" -->	 
							</td>
						</tr> 
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="60">÷�μ���</td>
							<td>÷������</td>
							<td>���ø�ũ</td>
						</tr>
						<tr  bgcolor="#FFFFFF">
							<td align="center" valign="top">
								<input type="button" value="����÷��" class="button" onClick="jsAttachFile('');">
								<div id="dFile"></div>
							</td>
							<td><input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120">
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<%IF iarapcd > 0 THEN%>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="60">��������</td> 
							<td>�����׸�</td>
							<td>�����������</td> 
						</tr>
						<tr bgcolor="#FFFFFF"  align="center"> 
							<td>[<%=iarapcd%>] <%=sarap_nm%></td>
							<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
						</tr>	
						</table>
					</td>
				</tr>
				<%IF blnPayEapp THEN%>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
								<td width="60" rowspan="2" style="padding:5px">�μ���<br>�ڱݱ���</td>
							<td width="300"  style="padding:5px"> �μ�</td> 
							<td width="205" style="padding:5px"> �ݾ�</td>
							<td width="205" style="padding:5px"> %</td>
						</tr>
						<tr> 
							<td colspan="3" bgcolor="#FFFFFF" valign="top">	 
							<div id="divPM"></div><br>
							<input type="hidden" name="iP" id="iP" value="">
							<input type="hidden" name="sP" id="sP" value="">
							<input type="hidden" name="mP" id="mP" value="">
							&nbsp;<input type="button" value="�μ����" onClick="jsSetPartMoney(1,'<%=sacc_use_cd%>','<%=sACC_GRP_CD%>');" class="button" ><Br><Br>
							</td>
						</tr>	 
						</table>
					</td>
				</tr>
					<%END IF%>
				<%END IF%>
				<tr>
					<td align="center" width="100%">
						<table border="0" cellpadding="5" cellspacing="0" width="100%">
							<tr>
								<td align="left"><input type="button" value="�ӽ�����" class="button" onClick="jsEappSubmit(0);"></td>
								<td align="right"><input type="button" value="������" style="color:blue;" class="button" onClick="jsEappSubmit(1);"></td>
							</tr>
						</table>
					</td>	
				</tr>
				</form>
				</table>
			</td>
		</tr> 
		</table>
	</td>
</tr> 
</table>	
		<!-- #include virtual="/lib/db/dbclose.asp" --> 	
<!-- ������ �� -->
</body>
</html>	
