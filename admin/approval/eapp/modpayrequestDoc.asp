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
<!-- #include virtual="/lib/offshop_function.asp"--> 
<%
Dim clsPay,clsMem,clseapp, clscomm,clsPM
Dim ireportidx,ipayrequestIdx,iarap_Cd,sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate ,iauthstate
Dim sadminid,dregdate,sarap_nm,sacc_cd,sacc_use_cd,sacc_nm,sedmsName,sedmscode ,iBizNo
Dim spartname ,ilastApprovalid,sscmLink, iauthposition
Dim chkPayRequest , spayrequestTitle
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate,pcomment ,susername
Dim arrAuth,arrComm,arrFile,arrRefer,arrReturn,arrProc,arrPM, arrPart
Dim intA, intC, intF, intR, intRA, intP , iAuthCount, intPart
Dim blnMod 
Dim mSumPrice, mSumRealPrice 
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2), pmistate(2)	,pmstatecd(2)					
Dim scust_cd, scustnm
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	 
Dim igbn, chkID,iRectMenu
Dim ipaytype, sCurrencyType, sCurrencyPrice,serpLinkType

ireportidx 		=  requestCheckvar(Request("iridx"),10)
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
iauthstate		= requestCheckvar(Request("ias"),10)
igbn					= requestCheckvar(Request("igbn"),1)
iRectMenu =	requestCheckvar(Request("iRM"),10)
blnMod = 1
chkID = 0
'���� �⺻ �� ���� ��������
set clsPay = new CPayRequest
	clsPay.Freportidx = ireportidx  
	clsPay.FpayrequestIdx = ipayrequestIdx  
	
	clsPay.fnGetPayRequestReceiveData 
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
  pcomment		= clsPay.FpayComment
  susername		 = clsPay.Fusername
  spartname		 = clsPay.Fpartname
  spayrequestTitle	= clsPay.FpayRequestTitle 							   
 	scust_cd						= clsPay.Fcust_cd
 	scustnm						=clsPay.Fcust_nm 
 	iBizNo					=clsPay.FBiz_no
 	ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	serpLinkType			= clsPay.FerpLinkType
	
	'//���������Ʈ
	arrProc			= clsPay.fnGetProcPayRequestList	
 
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

'�������, �ڸ�Ʈ, ���� ����Ʈ ��������
set clseapp = new CEApproval	
	clseapp.Freportidx 		= ireportidx  
	clseapp.FpayrequestIdx = ipayrequestIdx 	
 	 
 	arrAuth			= clseapp.fnGetAuthLineList  '��������
	arrComm			= clseapp.fnGetCommentList	'�ڸ�Ʈ
	arrFile			= clseapp.fnGetAttachFileList  '÷������ 
	arrPart			= clseapp.fnGetPartMoneyList
set clseapp = nothing  
 
 
 '-------------------------------------------
 '-- ������� ������ ����
 '1.������� DB ���尪 ������ ����
 '2.���°� (ipayrequeststate = 1)  : ��� ���δ���̰ų� ���������ڸ� ���οϷ��� ���� 
 ' ���� �� ���氡�� (����� db���� �⺻ ����� �Ǵ� ���� �α��� ����ڿ� �����Ҷ��� �α��� ����ڷ� �� ����)
 ' ���������ڰ� ���οϷ� �����϶��� ����Ұ��� ->  arrPM(2,intP) = 1 : �����������̰�, pmstatecd(0) = 0 : ���δ�� �����϶� ���氡��
 '------------------------------------------- 

IF isArray(arrAuth) THEN '1.������� �� ������ ����.  
		For intA = 0 To UBound(arrAuth,2)
			pmuserid(intA)  = arrAuth(2,intA) 
			pmusername(intA)= arrAuth(7,intA)
			pmjobname(intA) = arrAuth(10,intA)
			pmstate(intA)	= fnGetPayAuthState(arrAuth(3,intA), intA+1)
			IF not arrAuth(12,intA) and  arrAuth(3,intA) > 0 THEN 	pmstate(intA) = 	"[�ڵ�]"&pmstate(intA)  
				
			pmstatecd(intA)= arrAuth(3,intA)
			pmdate(intA)	= arrAuth(6,intA)  
		Next 
END IF	 
  
IF ipayrequeststate = 1  THEN  '2.�繫ȸ���� ������û�� ó�������� �����ͼ� ������ ���� 
Set clsPM	= new CPayManager
	clsPM.Fuserid = session("ssBctId")
	arrPM	= clsPM.fnGetPayManager 
Set clsPM 	= nothing  
 IF isArray(arrPM) THEN 
			For intP = 0 To UBound(arrPM,2)
			IF pmstatecd(intP) = 0 THEN 
			pmuserid(intP)  = trim(arrPM(1,intP))
			pmusername(intP)= arrPM(3,intP)
			pmjobname(intP) = arrPM(6,intP) 
			pmstate(intP)	= fnGetPayAuthState(0,arrPM(2,intP))
			pmdate(intP)	= "&nbsp;"
			END IF
			Next  
	END IF 
END IF 
 
'--------------------------------------------------	  
 
 set clsPay = nothing 
  
%> 
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script type="text/javascript" src="eapp.js"></script>
<script type="text/javascript">
	 
 	//������������
 	function jsPayDocUpdate(){
 			 var ichkVal=0;
	 for(i=0;i<document.frm.rdoDK.length;i++){
	 	if(document.frm.rdoDK[i].checked){
	 		 ichkVal = document.frm.rdoDK[i].value;
	 	}
	 }
	
	if (ichkVal ==0){
		alert("���������� �������ּ���");
		return;
	}
	
 
	if(ichkVal == 1 || ichkVal == 2){ 
	 	if(jsChkBlank(document.frm.sEK.value)){
	 		alert("���ݰ�꼭 �˻���ư�� ���� �������� ������ ������ּ���");
	 		document.frm.btnB1.focus();
	 		return;
	 	}  
		if(ichkVal  !=0 && ichkVal != 8 && ichkVal != 9){
			if(document.frm.mTP.value.replace(/\,/g,"") != document.frm.mprp.value.replace(/\,/g,"")){
		 		alert("������û�ݾװ� ���������� �ѱݾ��� �ٸ��ϴ�.Ȯ�� �� �ٽ� ������ּ���") 
		   	return; 
			}
		 } 
	}
		 
		
		
 	 if(confirm("�������� ��������� �Ͻðڽ��ϱ�?")){  
 	 	document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
		document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
		document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");
 		document.frm.submit();
 	}
 	} 
	</script>
</head>
<body topmargin="0" leftmargin="0"  bgcolor="#F4F4F4">
<table width="840" cellpadding="0" cellspacing="0" class="a" align="center">   
<tr>
	<td>
		<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a"   border="0"> 
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
				<form name="frm" method="post" action="procpayrequest.asp">  
				<input type="hidden" name="hidM" value="DU">
				<input type="hidden" name="irIdx" value="<%=ireportidx%>">
				<input type="hidden" name="iprIdx" value="<%=ipayrequestidx%>"> 
				<input type="hidden" name="hidPDidx" value="<%=ipayDocIdx%>"> 
				<input type="hidden" name="sbizno" value="<%=iBizNo%>">
				<input type="hidden" name="hidRU" value="top">
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80">������û��idx</td>
					<td bgcolor="#FFFFFF" ><%=ipayrequestidx%></td>
					<td rowspan="6" bgcolor="#FFFFFF" valign="top" width="300">
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
					<td bgcolor="<%= adminColor("tabletop") %>">��/�μ�</td>
					<td bgcolor="#FFFFFF"><%=spartname%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=dregdate%></td>
				</tr> 
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
					<td bgcolor="#FFFFFF">
					<% if  ((not isNull(ipaytype)) and (ipaytype<>"2")) and (ipayrequeststate="7") then %>
					<select name="hidPRS">
                	<option value="7" <%IF ipayrequeststate="7" THEN%>selected<%END IF%>>��������</option>
                	<option value="9" <%IF ipayrequeststate="9" THEN%>selected<%END IF%>>�����Ϸ�</option>
	                </select>
	                <input type="button" value="����" class="button" style="color:blue;" onClick="jsEditPayState(this)">
					<% else %>
					<%=fnGetPayRequestState(ipayrequeststate)%>
					<% end if %> 
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%IF ireportidx > 0 THEN%>
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
					<td><a href="javascript:jsPopView('/admin/approval/eapp/confirmeapp.asp?iridx=<%=ireportidx%>');">�󼼺���>></a></td>
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
					<td bgcolor="#FFFFFF"><a href="javascript:jsPopView('confirmpayrequest.asp?iridx=<%=ireportidx%>&ipridx=<%=arrProc(0,intP)%>&ias=<%=arrProc(5,intP)%>');">�󼼺��� >></a></td>
				</tr>
				<%Next%>
				</table>
			</td>
		</tr>
		<%END IF%>	
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
					<td><%IF mpayrequestprice <> "" THEN%><%=formatnumber(mpayrequestprice,0)%><%END IF%><input type="hidden" name="mprp" value="<%=mpayrequestprice%>"> </td>
					<td> <%=fnGetPayType(ipaytype)%> </td>
					<td>	  
							<span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							��ȭ�ݾ�: <%IF blnMod=0 THEN%><%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%><%ELSE%><%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"><%END IF%>	
							</span>
					</td>
				</tr> 
				<tr>	
					<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڱݿ뵵</td>  
					<td colspan="3" bgcolor="#FFFFFF"><%=sPayrequesttitle%></td>
				</tr>  
				<tr  align="center"   bgcolor="<%= adminColor("tabletop") %>">	
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
						<%IF blnMod = 1 THEN%>
						<input type="radio" name="rdoDK" value="1" <%IF ipaydockind ="1"   THEN%>checked<%END IF%> onClick="jsSetDocDis(1);">���ݰ�꼭-����&nbsp;&nbsp;
						 <input type="radio" name="rdoDK" value="2" <%IF ipaydockind ="2"  THEN%>checked<%END IF%> onClick="jsSetDocDis(2);">���ݰ�꼭-����&nbsp;&nbsp; 
					 	<input type="radio" name="rdoDK" value="5" <%IF ipaydockind ="5"  THEN%>checked<%END IF%> onClick="jsSetDocDis(5);">��Ÿ������ [��Ÿ�ε�]
						<br>
					 	<input type="radio" name="rdoDK" value="9" <%IF ipaydockind ="9"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">�������� [������/����]&nbsp;&nbsp;
					 	<input type="radio" name="rdoDK" value="8" <%IF ipaydockind ="8"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">��꼭 ���� ����(���ޱ�ó��)
					 	<%ELSE
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
					 	<%END IF%>
					</td>
				</tr>
					<tr>
					<td bgcolor="#FFFFFF" colspan="2">
						<div id="dSel1" style="display:<%IF not (ipaydockind = "1" or ipaydockind ="2") OR blnMod =0 THEN%>none<%END IF%>;padding:5px;">
						<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%"> 
						<tr> 
							<td bgcolor="#FFFFFF" colspan="5"> 
								<input type="button" name="btnB1" class="button" style="color:blue;" value="���ݰ�꼭 �˻�" onClick="jsGetTax('<%=iBizNo%>','<%=mpayrequestprice%>');">
								<span id="spB2" style="display:<%IF ipaydockind <> "1" THEN%>none<%END IF%>;"><input type="button" name="btnB2" class="button" value="XML ���" onClick="jsNewRegXML();" readonly></span>
								<span id="spB3" style="display:<%IF ipaydockind <> "2" THEN%>none<%END IF%>;"><input type="button" name="btnB3" class="button" value="���̼��ݰ�꼭 ���" onClick="jsNewRegHand();" readonly></span></td>
						</tr>
						</table>
						</div>
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
								%><input type="text" name="sVK" value="<%=strVat%>" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%> ><input type="hidden" name="rdoVK" value="<%=sVatKind%>"> 
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ������ </td>
							<td bgcolor="#FFFFFF"  colspan="3" style="border-bottom:1px solid #BABABA;"><input type="text" name="dID" value="<%=dissuedate%>" size="10" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%>  <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><!--%IF blnMod=1 THEN%--><!--img src="/images/calicon.gif" id="imgCal" align="absmiddle" border="0" onClick="jsPopCal('dID');"  style="cursor:hand;"  disabled --><!--%END IF%--></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ǰ�� </td>
							<td colspan="5" bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"> <input type="text" name="sINm" value="<%=sItemName%>" size="40"  <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%> <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>> </td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> �ѱݾ� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mTP" value="<%=formatnumber(mTotPrice,0)%>" size="10" OnKeyUp="jsSetPrice();" class="text_ro" style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mTotPrice,0)%><%END IF%>��</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> ���ް� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mSP" value="<%=formatnumber(mSupplyPrice,0)%>" size="10" class="text_ro"  style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mSupplyPrice,0)%><%END IF%>��</td>
							<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"> �ΰ��� </td>
							<td bgcolor="#FFFFFF" style="border-bottom:1px solid #BABABA;"><%IF blnMod=1 THEN%><input type="text" name="mVP" value="<%=formatnumber(mVatPrice,0)%>" size="10" class="text_ro"  style="text-align:right" readonly <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>><%ELSE%><%=formatnumber(mVatPrice,0)%><%END IF%>��</td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> ����û���ι�ȣ </td>
								<td bgcolor="#FFFFFF" style="border-right:1px solid #BABABA;"><input type="text" name="sEK" value="<%=setaxkey%>" size="35" <%IF blnMod=0 THEN%>style="border:0" class="text"<%ELSE%>class="text_ro"  readonly<%END IF%>  <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>></td>
								<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center" style="border-right:1px solid #BABABA;"> ��� </td>
								<td bgcolor="#FFFFFF"  colspan="3"><input type="text" name="sDB" value="<%=sDocBigo%>" size="40" <%IF blnMod=0 THEN%>style="border:0" <%END IF%>   class="text" <%IF ipaydockind ="9"  THEN%>disabled<%END IF%>> </td>
						</tr> 
						</table>
			 			</div>
			 		</td>
			 	</tr> 
				</table>
			</td>
		</tr>
		<tr>
			<td align="right"><input type="button" value="�������� �������" style="color:blue;" class="button" onClick="jsPayDocUpdate();"></td>
		</tr>
	</form>
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
						</div>&nbsp;
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
							<td>[<%=iarap_cd%>] <input type="text" name="sANM" value="<%=sarap_nm%>" style="border:0" readonly ><input type="hidden" name="iaidx" value="<%=iarap_cd%>" class="text"></td>
							<td>[<%=sacc_use_cd%>] <input type="text" name="sACCNM" value="<%=sacc_nm%>" style="border:0" readonly><input type="hidden" name="sACC" value="<%=sacc_cd%>"  class="text"></td>
						</tr>	 
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
						<td width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"> <%=formatnumber(arrPart(2,intPart),0)%> ��</td>
						<td width="200" style="border-bottom:1px solid #BABABA;" align="center"> <%=formatnumber((arrPart(2,intPart)/mpayrequestprice)*100,1)%> %</td>
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
				<form name="frmFS" method="post" action="procpayrequest.asp"> 
				<input type="hidden" name="hidM" value="FS">
				<input type="hidden" name="irIdx" value="<%=ireportidx%>">
				<input type="hidden" name="iprIdx" value="<%=ipayrequestidx%>">  
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="3" width="60">�濵������<br>�����׸�</td>
					<td>����������</td> 
					<td>����(�Ա�)��</td>
					<td>�ش���(����)</td>
					<td>�������⿩��</td>
				</tr>
				<tr bgcolor="#FFFFFF" align="center"> 
					<td><%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%> </td> 
					<td><%IF dpayrealdate <> "" THEN%><%=formatdate(dpayrealdate,"0000-00-00")%><%END IF%></td>
					<td>  <%IF syyyymm <> "" THEN%><%=year(syyyymm)%><%END IF%>
						 ��
						 <%IF syyyymm <> "" THEN%><%=month(syyyymm)%><%END IF%> �� 
						</td>
					<td> 
						 <%IF blnTakeDoc THEN%>Y<%ELSE%>N<%END IF%> 
					</td>
				</tr> 
				<tr bgcolor="#FFFFFF">
					<td colspan="5">*COMMENT<br> 
				 <%=pcomment%> <Br>
					</td>
				</tr>	
				</table>
			</td>
		</tr>  
		</table>
	</form>
	</td>
</tr>
<tr>
	<td height="100">&nbsp;</td>
</tr>
</table> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
