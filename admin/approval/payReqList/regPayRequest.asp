<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"--> 
<!-- #include virtual="/lib/offshop_function.asp"--> 
<%
Dim clseapp, clsMem,clsPM, clsPay
Dim ipayrequestidx, spartname,ireportidx
Dim arrPM,intP,intPart
Dim pmuserid(2), pmusername(2), pmjobname(2),pmstate(2),pmdate(2),pmstatecd(2)	
Dim dpayrequestdate,mpayrequestprice,iinBank,saccountNo,saccountHolder,dpaydate,ioutBank,dpayrealdate,mpayrealprice,syyyymm,blnTakeDoc,ipayrequeststate ,spayrequesttitle
Dim sreportName,mreportPrice,iscmlinkno,sbigo,sreportcontents,ireportstate ,sadminid,sedmsName,sedmscode,ilastApprovalid,dregdate,pcomment,iBizNo
Dim arrComm, arrFile,arrPart, arrReturn,arrAuth  
Dim  intF,   intRA  
Dim iarap_cd, arrAccConts , intA, intAK, sarap_nm, sacc_nm, sacc_cd
dim   susername 
Dim scust_cd, scustnm,soutBankName
Dim ipayDocIdx,ipaydockind,svatkind,dissuedate,sitemname,mtotprice,msupplyprice,mvatprice,setaxkey,sDocbigo,sattachfile	
Dim sMode
Dim iauthposition,iauthstate
Dim ipaytype, sCurrencyType, sCurrencyPrice,serpLinkType,sACC_GRP_CD
Dim blnMod,sRectAuthId, arrProc
blnMod = 1
sMode ="I"
 ireportidx = 0
 ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
 IF ipayrequestIdx = "" THEN ipayrequestIdx = 0 
	sRectAuthId =  session("ssBctId")
IF ipayrequestIdx > 0 THEN
 	sMode = "U"
 '���� �⺻ �� ���� ��������
set clsPay = new CPayRequest 
	clsPay.FpayrequestIdx = ipayrequestIdx  
	
	clsPay.fnGetPayRequestReceiveData 
	iarap_cd				 = clsPay.Farap_cd		
	sreportName      = clsPay.FreportName    
	mreportPrice     = clsPay.FreportPrice   
	iscmlinkno       = clsPay.Fscmlinkno     
	sbigo            = clsPay.Fbigo      
	ireportstate     = clsPay.Freportstate  
	sadminid         = clsPay.Fadminid        
	sarap_nm     		 = clsPay.Farap_nm
	sacc_cd					 = clsPay.Facc_cd
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
 	scust_cd					= clsPay.Fcust_cd
 	scustnm						=clsPay.Fcust_nm
 	soutBankName			= clsPay.FoutBankName
 	susername					= clsPay.Fusername						 
	spartname					= clsPay.Fpartname
  spayrequestTitle	= clsPay.FpayRequestTitle
  ireportidx				= clsPay.Freportidx
  iBizNo						= clsPay.FBiz_no
  ipaytype 					= clsPay.Fpaytype
	sCurrencyType 		= clsPay.Fcurrencytype
	sCurrencyPrice		= clsPay.Fcurrencyprice
	serpLinkType			= clsPay.FerpLinkType
	sACC_GRP_CD				= clsPay.FACC_GRP_CD
		
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
	clseapp.FreportIdx = ireportidx
	clseapp.FpayrequestIdx = ipayrequestIdx 	  
  
	arrAuth			= clseapp.fnGetAuthLineList  
	arrPart			= clseapp.fnGetPartMoneyList
	arrReturn		= clseapp.fnGetAuthLineReturnList 
	arrFile			= clseapp.fnGetAttachFileList
set clseapp = nothing  	
END IF 

'���缱 ����Ʈ ����
	IF isArray(arrAuth) THEN
		For intA = 0 To UBound(arrAuth,2)
			pmuserid(intA)  = arrAuth(2,intA) 
			pmusername(intA)= arrAuth(7,intA)
			pmjobname(intA) = arrAuth(10,intA)
			pmstate(intA)	= fnGetPayAuthState(arrAuth(3,intA),intA+1)
			IF not arrAuth(12,intA) and  arrAuth(3,intA) > 0 THEN 	pmstate(intA) = 	"[�ڵ�]"&pmstate(intA)  
			pmstatecd(intA)= arrAuth(3,intA)
			pmdate(intA)	= arrAuth(6,intA)
			IF pmdate(intA) <> "" THEN pmdate(intA)	= formatdate(pmdate(intA),"0000-00-00")   
		Next 
END IF	 
   
IF ipayrequeststate <= 1  THEN  '2.�繫ȸ���� ������û�� ó�������� �����ͼ� ������ ���� 
		'�繫ȸ���� ������û�� ó��������
		Set clsPM	= new CPayManager
			clsPM.FisDef = 1
			arrPM	= clsPM.fnGetPayManager 
		Set clsPM 	= nothing 
 
		IF isArray(arrPM) THEN
			For intP = 0 To UBound(arrPM,2)
			IF pmstatecd(intP) = 0 THEN  
			pmuserid(intP)  = arrPM(1,intP)	 
			pmusername(intP)= arrPM(3,intP)
			pmjobname(intP) = arrPM(6,intP) 
			pmstate(intP)	= fnGetPayAuthState(0,intP+1)
			pmdate(intP)	= "&nbsp;"
			END IF
			Next 
		END IF
	END IF
 IF pmstatecd(0) = 0 THEN 
 	iauthposition = 1
 	iauthstate = 0
 ELSE 
 	iauthposition = 2
 	iauthstate = 1
 END IF
if susername = "" then susername =  session("ssBctCname")
if spartname = "" then
'�μ��� ��������
set clsMem = new CTenByTenMember
	clsMem.Fpart_sn = session("ssAdminPsn")
	clsMem.fnGetPartName 
 	spartname = clsMem.Fpart_name 
set clsMem = nothing 
end if

IF serpLinkType <> "" THEN blnMod = 0
%> 
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/admin/approval/eapp/eapp.js"></script>  
<script type="text/javascript">
	//������û�� ���
	
function jsPayReqSubmit(sMode){   
		if(jsChkBlank(document.frm.dprd.value) ){
			alert("������û���� �Է����ּ���");
			return;
		}
		
		if(jsChkBlank(document.frm.mprp.value) ){
			alert("������û�ݾ��� �Է����ּ���");
			document.frm.mprp.focus();
			return;
		}
		 
		var arapcd=document.frm.iaidx.value;
		
		if(jsChkBlank(arapcd) ){
			alert("�����׸��� ������ּ���");
			document.frm.btnarap.focus();
			return;
		}
		
		if(arapcd!=351){	//�����׸�-��Ÿ������ �ƴҰ��
		if(jsChkBlank(document.frm.hidcustcd.value) ){
			alert("�ŷ�ó�� �������ּ���");
			return;
		} 
	} 

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
		var mTotPrice = 0; 
	 
 		if(typeof(document.all.mPM) =="undefined"){
 			alert("�ڱݱ��� �μ��� ������ּ���");
 			return;
 		}
 
		if(typeof(document.all.mPM.length)!="undefined"){
			for(i=0;i<document.all.mPM.length;i++){
				mTotPrice = mTotPrice + parseInt(document.all.mPM[i].value.replace(/\,/g,""));
			}
		}else{
			 mTotPrice = document.all.mPM.value.replace(/\,/g,"");
		}
		 
		if(mTotPrice !=document.frm.mprp.value.replace(/\,/g,"")){
			alert(mTotPrice+"/"+document.frm.mprp.value.replace(/\,/g,"")+"�μ��� �ڱݱ��бݾװ� ������û�ݾ��� �ٸ��ϴ�.");
			return;
		}  
		 
		 $("input[name='sFileP[]']").each( function(index,elem) {
		     var a = $(elem).val();
		     if( document.frm.sFile.value==""){
		     	document.frm.sFile.value = a;
		    }else{
		     document.frm.sFile.value = document.frm.sFile.value + ","+a;
		   }
		  });
		  
		if(confirm("������û�Ͻðڽ��ϱ�?\n����û���ι�ȣ�� �߸� ����ϰų�, �����꼭 �� ��� ������ �������� ���������� �������� ������ ����Ϸᰡ ���� �ʽ��ϴ�.")){
				document.all.mprp.value = document.all.mprp.value.replace(/\,/g,"");
				if(typeof(document.all.mPM) !="undefined"){  
			  	if(typeof(document.all.mPM.length)!="undefined"){ 
			  		for(i=0;i<document.all.mPM.length;i++){ 
					 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
					 		if(document.frm.mP.value ==""){
							 			document.frm.mP.value =document.all.mPM[i].value;
							 }else{
							 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
							 }
						}
					}else{	
							document.all.mPM.value = document.all.mPM.value.replace(/\,/g,""); 
							document.frm.mP.value =	document.all.mPM.value;
					}	
				}	
		document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
		document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
		document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");
		document.frm.hidPRS.value=1;
		document.frm.hidM.value =sMode;
		document.frm.submit(); 
	}
} 
  
  
  //�������
  function jsPayReqUpdate(sMode){
  	$("input[name='sFileP[]']").each( function(index,elem) {
		     var a = $(elem).val();
		     if( document.frm.sFile.value==""){
		     	document.frm.sFile.value = a;
		    }else{
		     document.frm.sFile.value = document.frm.sFile.value + ","+a;
		   }
		  });
  		document.all.mprp.value = document.all.mprp.value.replace(/\,/g,"");
				if(typeof(document.all.mPM) !="undefined"){  
			  	if(typeof(document.all.mPM.length)!="undefined"){ 
			  		for(i=0;i<document.all.mPM.length;i++){ 
					 		document.all.mPM[i].value = document.all.mPM[i].value.replace(/\,/g,"");
					 		if(document.frm.mP.value ==""){
							 			document.frm.mP.value =document.all.mPM[i].value;
							 }else{
							 			document.frm.mP.value = document.frm.mP.value+","+document.all.mPM[i].value;
							 }
						}
					}else{	
							document.all.mPM.value = document.all.mPM.value.replace(/\,/g,""); 
							document.frm.mP.value =	document.all.mPM.value;
					}	
				}	
	
				document.frm.mTP.value= document.frm.mTP.value.replace(/\,/g,"");
				document.frm.mSP.value= document.frm.mSP.value.replace(/\,/g,"");
				document.frm.mVP.value= document.frm.mVP.value.replace(/\,/g,"");
  			document.frm.hidPRS.value=0;
  			document.frm.hidM.value =sMode; 
  			document.frm.submit();
  }
 
</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >  
<tr>
	<td> 
		<form name="frm" method="post" action="/admin/approval/eapp/procpayrequest.asp">  
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" >
		<input type="hidden" name="hidM" value="">   
		<input type="hidden" name="irIdx" value="<%=ireportidx%>">
		<input type="hidden" name="iprIdx" value="<%=ipayrequestidx%>">
		<input type="hidden" name="hidPRS" value="0">		  
		<input type="hidden" name="iAP" value="<%=iauthposition%>"><!-- �ݷ��϶��� ������ġ�� 2�� ��찡 �ִ�.-->
		<input type="hidden" name="blnL" value="1">
		<input type="hidden" name="iptt" value="2"> 
		<input type="hidden" name="hidRU" value="/admin/approval/payreqList/">
		<input type="hidden" name="hidcustcd" value="<%=scust_cd%>">
		<input type="hidden" name="hidPDidx" value="<%=ipayDocIdx%>">
		<input type="hidden" name="sbizno" value="<%=iBizNo%>">
		<tr>
			<td>
				<table width="100%" cellpadding="5" cellspacing="1" class="a">
				<tr>
					<td class="verdana-large"><b>������û��</b></td>
					<td align="right"><img src="/images/admin_logo_10x10.jpg"></td>
				</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tR>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">������û��Idx</td>
					<td bgcolor="#FFFFFF" ><%IF ipayrequestidx > 0 THEN%><%=ipayrequestidx%><%END IF%></td>
					<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="300">
						<!------������ ����Ʈ------------------------------------------------------------>
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0">
						<tr>
							<td  valign="top" width="150"> 
								<table width="100%"  cellpadding="5" cellspacing="0" class="a"> 
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>����������</td> </tr>
									<tr align="Center"><td><%=pmstate(0)%></td> </tr>
									<tr align="Center"> <td><input type="hidden" name="hidAI1" value="<%=pmuserid(0)%>">   <%=pmusername(0)%>&nbsp;<%=pmjobname(0)%></td> </tr>
									<tr align="Center"><td><%=pmdate(0)%></td></tr>
								</table> 
							</td> 
							<td  valign="top">
								<table width="100%"  cellpadding="5" cellspacing="0" class="a"> 
									<tr align="Center" bgcolor="<%= adminColor("tabletop") %>"><td>�繫ȸ����</td> </tr>
									<tr align="Center"><td><font color="gray"><%=pmstate(1)%></font></td> </tr>
									<tr align="Center"> <td><input type="hidden" name="hidAI2" value="<%=pmuserid(1)%>"> <%=pmusername(1)%>&nbsp;<%=pmjobname(1)%></td> </tr>
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
					<td bgcolor="#FFFFFF"><%= susername%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ۼ���</td>
					<td bgcolor="#FFFFFF"><%=chkIIF(dregdate="",date(),dregdate)%></td>
				</tr> 
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(ipayrequeststate)%></td>
				</tr>
				</table>
			</td>
		</tr> 
			<tr><!--����ǰ�Ǽ�-->
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
					<td><%IF mreportprice <> "" THEN%><%=formatnumber(mreportprice,0)%><%END IF%></td>
					<td><a href="javascript:jsPopView('/admin/approval/eapp/confirmeapp.asp?iridx=<%=ireportidx%>');">�󼼺���>></a></td>
				</tr>
				</table>
			</td>
		</tr><!--//����ǰ�Ǽ�-->
		<!--���������-->
		<% Dim totPrice
		totPrice = 0
		IF isArray(arrProc) THEN%>
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
				<%For intP = 0 To UBound(arrProc,2)
					totPrice = totPrice + arrProc(2,intP)
				%>
				<tr align="center">
					<td bgcolor="#FFFFFF"><%=arrProc(0,intP)%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(3,intP) <> "" THEN%><%=formatdate(arrProc(3,intP),"0000-00-00")%><%END IF%></td>
					<td bgcolor="#FFFFFF"><%IF arrProc(2,intP) <> "" THEN%><%=formatnumber(arrProc(2,intP),0)%><%END IF%></td>
					<td bgcolor="#FFFFFF"><%=fnGetPayRequestState(arrProc(4,intP))%></td>
					<td bgcolor="#FFFFFF"><a href="javascript:jsPopView('regpayrequest.asp?iridx=<%=ireportidx%>&ipridx=<%=arrProc(0,intP)%>');">�󼼺��� >></a></td>
				</tr>
				<%Next%>
				</table>
			</td>
		</tr>
		<%END IF%>	<!--//���������-->
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="80" rowspan="6">������û<br>����</td>
					<td>������û��</td>
					<td>������û�ݾ�(��)</td>
					<td>�������</td> 
					<td width="250" >���</td> 
				</tr>
				<tr align="center"  bgcolor="#FFFFFF">	 
				<td><input type="text" name="dprd" value="<%IF dpayrequestdate <> "" THEN%><%=formatdate(dpayrequestdate,"0000-00-00")%><%END IF%>" size="10" <%IF blnMod = 0 THEN%>style="border:0" readonly<%END IF%>><%IF blnMod = 1 THEN%><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('dprd');"  style="cursor:hand;"><%END IF%></td>
					<td><input type="text" name="mprp" value="<%IF mpayrequestprice <> "" THEN%><%=formatnumber(mpayrequestprice,0)%><%END IF%>" size="15" style="text-align:right;<%IF blnMod = 0 THEN%>border:0;" readonly<%ELSE%>"<%END IF%> onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" ></td>
					<td> <%  
						IF blnmod= 0 THEN 
							%><%=fnGetPayType(ipaytype)%>
					 <%ELSE%>	
							<select name="selPT" onChange="jsChFC();" class="select">
								<%sboptPayType ipaytype%> 
							</select>
						<%END IF%>
					</td>
					<td>	 
							<span id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> 
							��ȭ�ݾ�: <%IF blnMod=0 THEN%><%=sCurrencyType%>&nbsp;<%=sCurrencyPrice%><%ELSE%><%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"><%END IF%>	
							</span>
					</td>
				</tr> 
				<tr>	
					<td   align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڱݿ뵵</td>  
					<td colspan="3" bgcolor="#FFFFFF"><input type="text" class="text" name="sprt" value="<%=spayrequesttitle%>" size="60" <%IF blnMod= 0  THEN%>style="border:0" readonly<%END IF%>></td>
				</tr> 
				<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">	 
					<td>�ŷ�ó</td>
					<td>�����</td>
					<td>�����ָ�</td>
					<td>���¹�ȣ</td>
				</tr>
				
				<tr align="center" bgcolor="#FFFFFF">	
					<td><input type="text" name="scustnm" value="<%=scustnm%>" size="20"  readonly class="text_ro"> <%IF blnMod = 1 THEN%><input type="button" class="button" value="����" onClick="jsGetCust('<%=scust_cd%>')"><%END IF%></td>
					<td><input type="text" name="selIB" value="<%=iinBank%>" readonly class="text_ro"> </td>
					<td><input type="text" name="sah" value="<%IF saccountholder <> "" THEN%><%=saccountholder%><%END IF%>" size="15" readonly class="text_ro"></td>
					<td><input type="text" name="san" value="<%IF saccountno <> "" THEN%><%=saccountno%><%END IF%>" size="20" readonly class="text_ro"></td>
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
						<input type="radio" name="rdoDK" value="1" <%IF ipaydockind ="1" THEN%>checked<%END IF%> onClick="jsSetDocDis(1);">���ݰ�꼭-����
						 <input type="radio" name="rdoDK" value="2" <%IF ipaydockind ="2"  THEN%>checked<%END IF%> onClick="jsSetDocDis(2);">���ݰ�꼭-���� 
						<!--<input type="radio" name="rdoDK" value="3" <%IF ipaydockind ="3"  THEN%>checked<%END IF%> onClick="jsSetDocDis(3);">���ݿ�����-�ҵ������ 
						<input type="radio" name="rdoDK" value="4" <%IF ipaydockind ="4"  THEN%>checked<%END IF%> onClick="jsSetDocDis(4);">���ݿ�����-���������� -->
						<input type="radio" name="rdoDK" value="5" <%IF ipaydockind ="5"  THEN%>checked<%END IF%> onClick="jsSetDocDis(5);">��Ÿ������ [��Ÿ�ε�]
						<br>
					 	<input type="radio" name="rdoDK" value="9" <%IF ipaydockind ="9"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">�������� [������/����]
					 	<input type="radio" name="rdoDK" value="8" <%IF ipaydockind ="8"  THEN%>checked<%END IF%> onClick="jsSetDocDis(0);">��꼭 ���� ����(���ޱ�ó��)
					 	[������ �� ��꼭�� ���߿� �������] 
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
						%> 
						<%=strDoc%>
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
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2" width="80">÷�μ���</td>
					<td>÷������</td>
					<td>���ø�ũ</td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<td align="center" valign="top"> 
						 <%IF blnMod = 1 THEN %><input type="button" value="����÷��" class="button" onClick="jsAttachFile('');"> <%END IF%>
						<div id="dFile"> 
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2)
							IF arrFile(2,intF) = 0 THEN Exit For
								arrF = split(arrFile(1,intF),"/")  
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0) 
						%>
						<div id="dF<%=sFName%>"> <a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a> <%IF blnMod = 1 THEN %>&nbsp;<input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"> </a><%END IF%>
						<input type="hidden"  name="sFileP[]"   value="<%=arrFile(1,intF)%>"></div>
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
						<input type="text" name="sL" size="60" maxlength="120" value="<%=arrFile(1,intF2)%>"  <%IF blnMod = 0 THEN%>style="border:0;cursor:hand;" readonly onClick="jsFileLink('<%=arrFile(1,intF2)%>');"<%END IF%>><br>
						<% iCount = iCount + 1
						Next
						END IF 
						For intF3= iCount To 4
						%>
						<input type="text" name="sL" size="60" maxlength="120" <%IF blnMod = 0 THEN%>style="border:0" readonly<%END IF%>><br>
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
					<td rowspan="3" width="80">��������</td> 
					<td>�����׸�</td>
					<td>�����������</td> 
				</tr>
				<tr bgcolor="#FFFFFF"  align="center"> 
					<td><input type="text" name="sANM" value="<%=sarap_nm%>" style="border:0" readonly><input type="hidden" name="iaidx" value="<%=iarap_cd%>"></td>
					<td><input type="text" name="sACCNM" value="<%=sacc_nm%>" style="border:0" readonly><input type="hidden" name="sACC" value="<%=sacc_cd%>"></td>
				</tr>	 
				<%IF blnMod = 1 THEN%>
				<tr bgcolor="#FFFFFF">
					<td colspan="2"><input type="button" class="button" id="btnarap" value="�����׸� ���/����" onClick="jsGetARAP();"></td>
				</tr> 
					<%END IF%>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center" >
					<td width="80" rowspan="2" style="padding:5px">�μ���<br>�ڱݱ���</td>
					<td width="321" style="padding:5px" > �μ�</td> 
					<td width="201" style="padding:5px"> �ݾ�</td>
					<td width="160" style="padding:5px"> %</td>
				</tr>
				<tr>
					<td colspan="3" bgcolor="#FFFFFF" valign="top">	 
					<div id="divPM">
					<%dim arrPV, arrPT 
					IF isArray(arrPart) THEN %>
						<table border="0" cellpadding="3" cellspacing="0" class="a" width="737">  
					<%	For intPart = 0 To UBound(arrPart,2) 	
						IF intPart > 0 THEN
								arrPV = arrPV&"," 
								arrPT =arrPT&"," 
							END IF	
							arrPV = arrPV&arrPart(1,intPart)
							arrPT = arrPT&arrPart(3,intPart)  
					%>
					<tr>
						<td width="150" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(4,intPart)%></td>
						<td width="150" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><%=arrPart(3,intPart)%> </td>
						<td width="201" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;" align="center"><input type="text" name="mPM" id="mPM" size="20" value="<%=formatnumber(arrPart(2,intPart),0)%>" style="text-align:right;<%IF blnMod = 0   THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="auto_amount(this.form,this);jsSetMoney('m',<%=intPart%>,2);"  onKeypress="num_check()"  > ��</td>
						<td width="160" style="border-bottom:1px solid #BABABA;" align="center"><input type="text" name="iPM"  size="4" value="<%IF mpayrequestprice <> 0 AND arrPart(2,intPart)<> 0 THEN%><%=formatnumber((arrPart(2,intPart)/mpayrequestprice)*100,1)%><%END IF%>"  style="text-align:right;<%IF blnMod = 0  THEN%>border:0" readonly<%else%>"<%END IF%> onKeyUp="jsSetMoney('i',<%=intPart%>,2);">%</td>
					</tr>   
					<%	Next %>
					</table>
					<%END IF%>
					</div><br> 
					<input type="hidden" name="iP" id="iP" value="<%=arrPV%>">
					<input type="hidden" name="sP" id="sP" value="<%=arrPT%>">
					<input type="hidden" name="mP" id="mP" value="">
					<%IF blnMod = 1 THEN%> <br>&nbsp;	<input type="button" value="�μ� ���/����" onClick="jsSetPartMoney(2,'<%=sACC_GRP_CD%>');" class="button"><br><BR> <%END IF%>
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
			<td align="center">
		<%IF blnMod = 0 THEN%> 
			ERP ���� ���� ���´� ���� �Ұ���
			<% ELSE %>
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td align="left"><input type="button" value="��������" class="button" onClick="jsPayReqUpdate('<%=sMode%>');"></td> 
						<%IF ipayrequeststate <=1 THEN%><td align="right"><input type="button" value="������û" class="button"  onClick="jsPayReqSubmit('<%=sMode%>');"></td>	<%END IF%> 
					</tR>
				</table>
			<% end if %>
			</td>
		</tr> 
		</table> 
		</form>
	</td>
</tr> 
<%IF ipayrequeststate >=7 THEN '���� �� ����Ȯ�ε� �����϶��� ���뺸���ش�%>
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
					<td><%IF dpaydate <> "" THEN%><%=formatdate(dpaydate,"0000-00-00")%><%END IF%></td>
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
		<%END IF%>
<%IF ipayrequestidx > 0 THEN%>
<tr>
	<td style="padding-top:20px;"> 
		<!-- #include virtual="/admin/approval/eapp/incComment.asp" --> 
	</td>
</tr>
<%END IF%>
</table> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
