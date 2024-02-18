<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ �� ����Ʈ  
' History : 2011.05.30 ������  ����
'			2018.10.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/commonCls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim sMode, clsPart, arrType, sPartTypeName, clsComm,clsPartMoney
Dim iOpExpPartIdx, iPartTypeIdx, sOpExpPartName, blnUsing,arrPartsn, arrDepartmentID, intLoop, iPartsn, iDepartmentID, ijobsn, sAdminID,susername,ipart_sn,sjobname
Dim iOutBank,sOutAccNo,sOutName,ieappPartIdx,iOrderNo,scustnm,scust_cd
Dim sbizcd,sbiznm,sarap_Cd,sarap_nm, sCardCo, sCardNo, clsCard, arrC, intC, sregid
iOpExpPartIdx = requestCheckvar(Request("hidOEP"),10)
sMode ="I"
sregid = session("ssBctId")

'���а� ��������
Set clsPart = new COpExpPart
	IF iOpExpPartIdx <> "" THEN
		sMode ="U"
		clsPart.FOpExpPartIdx  = iOpExpPartIdx
		clsPart.fnGetOpExpPartData
		iPartTypeIdx 	= clsPart.FPartTypeIdx
		sPartTypeName   = clsPart.FPartTypeName
		sOpExpPartName  = clsPart.FOpExpPartName
		iOutBank	    = clsPart.FOutBank
		sOutAccNo		= clsPart.FOutAccNo
		sOutName	    = clsPart.FOutName
		sbizcd		 	= clsPart.Fbizsection_Cd
		sbiznm			= clsPart.Fbizsection_nm
		sarap_Cd    	= clsPart.Farap_cd
		sarap_nm		= clsPart.Farap_nm
		iOrderNo	  	= clsPart.FOrderNo
		blnUsing 		= clsPart.FIsUsing
		sAdminID		= clsPart.FAdminID
		susername		= clsPart.Fusername
		ipart_sn		= clsPart.Fpart_sn
		ijobsn			= clsPart.Fjob_sn
		sjobname		= clsPart.Fjobname
		scust_cd		= clsPart.Fcust_Cd
		scustnm			= clsPart.Fcust_nm
		sCardCo			= clsPart.FCardCo
		sCardNo			= clsPart.FCardNo
		arrPartsn		= clsPart.fnGetOpExppartInfoList
		arrDepartmentID = clsPart.fnGetOpExpDepartmentInfoList
		sregid = sAdminID
	END IF
	arrType = clsPart.fnGetOpExpPartTypeList
Set clsPart = nothing

dim clsMem,iregdepartmentid
'�μ��� ��������
set clsMem = new CTenByTenMember
	clsMem.Fuserid = sregid
	clsMem.fnGetDepartmentInfo
	iregdepartmentid		= clsMem.Fdepartment_id 
 set clsMem = nothing
 
	IF isArray(arrPartsn) THEN
		FOR intLoop = 0 To UBound(arrPartsn,2)
		 if intLoop = 0 then
		 	iPartsn =  arrPartsn(0,intLoop)
		 else
		 	iPartsn = iPartsn&","&arrPartsn(0,intLoop)
		end if
		NEXT
	END IF

	if isArray(arrDepartmentID) then
		for intLoop = 0 To UBound(arrDepartmentID,2)
			if intLoop = 0 then
				iDepartmentID =  arrDepartmentID(0,intLoop)
			else
				iDepartmentID = iDepartmentID&","&arrDepartmentID(0,intLoop)
			end if
		next
	end if
%>
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<%
 'ī��� ����Ʈ
 Set clsCard = new CCardCorp
 	arrC = clsCard.fnGetCardCorp
 Set clsCard = nothing
%>
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
<script type="text/javascript">
<!--
	//���
	function jsPartSubmit(){
		if(document.frm.selPT.value==0 && document.frm.sPTN.value==""){
			alert("���и��� ������ּ���");
			return;
		}

		if( document.frm.sPN.value==""){
			alert("��� ���������� �Է����ּ���");
			return;
		}

		document.frm.submit();
	}

	//���� 
	function jsChPT(iValue){
		if (iValue==0){
			document.all.divPT.style.display = "";
		}else{
			document.all.divPT.style.display = "none";
		}
	}

	//�μ� �߰�
	function jsAddPart(){
		var winPart = window.open("popAddPart.asp","popPart","width=800,height=960,scrollbars=yes,resizable=yes");
		winPart.focus();
	}

	function jsAddDepartment() {
		var winPart = window.open("popAddDepartment.asp","jsAddDepartment","width=800,height=960,scrollbars=yes,resizable=yes");
		winPart.focus();
	}

	//���úμ� ����
	function jsDelPart(iValue){
		var arrValue = document.frm.hidPsn.value.split(",");
		if(typeof(arrValue.length)=="undefined"){
			document.frm.hidPsn.value  = ""
		}else{
			if(arrValue[0] == iValue){
				document.frm.hidPsn.value  = document.frm.hidPsn.value.replace(iValue,"");
			}else{
				document.frm.hidPsn.value  = document.frm.hidPsn.value.replace(","+iValue,"");
			}
		}
		eval("document.all.dP"+iValue).outerHTML = "";
	}

	function jsDelDepartment(iValue) {
		var arrValue = document.frm.hidDPid.value.split(",");
		if (typeof(arrValue.length) == "undefined") {
			document.frm.hidDPid.value  = "";
		} else {
			if(arrValue[0] == iValue) {
				document.frm.hidDPid.value  = document.frm.hidDPid.value.replace(iValue,"");
			}else {
				document.frm.hidDPid.value  = document.frm.hidDPid.value.replace(","+iValue,"");
			}
		}
		eval("document.all.dDP"+iValue).outerHTML = "";
	}

	//����� ���
	function jsRegID(department_id,workerid){
		var winRI = window.open('/admin/member/tenbyten/popSetID.asp?fN=frm&department_id='+department_id+'&workerid='+workerid ,'popAL',"width=500,height=960,scrollbars=yes,resizable=yes");
		winRI.focus();
	}

	//�ڱݰ����μ� ����
	function jsGetPart(){
		var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popGetBizOne',"width=800,height=600,scrollbars=yes,resizable=yes");
		winP.focus();
	}

	//�ڱݰ����μ� ���
	function jsSetPart(selUP, sPNM){
		document.frm.selUP.value = selUP;
		document.frm.sPNM.value = sPNM;
	}

	//���޼����׸� ����
	function jsGetARAP(){
		var winP = window.open("/admin/expenses/account/popGetOpExpArap.asp","popARAP","width=800,height=960,scrollbars=yes,resizable=yes");
		winP.focus();
	}

	//���� �����׸� ��������
 	function jsSetARAP(dAC, sANM,sACCC,sACCNM){
 		document.frm.dAC.value = dAC;
 		document.frm.sANM.value = sANM;
 	}

	//�ŷ�ó ���� ����
	function jsGetCust(){
		var Strparm="";
		var cust_cd = "<%=scust_cd%>";
		if (cust_cd!=""){
			Strparm = "?selSTp=1&sSTx="+ cust_cd;
		}
		var winC = window.open("/admin/linkedERP/cust/popGetCust.asp"+Strparm,"popC","width=1280,height=960,scrollbars=yes,resizable=yes");
		winC.focus();
	}


	 //�ŷ�ó ����
	 function jsSetCust(custcd, custnm,banknm, accno, snm ){
	 document.frm.hidcustcd.value = custcd;
	 document.frm.scustnm.value = custnm;
	 document.frm.selOB.value = banknm;
	 document.frm.sOAN.value = accno;
	 document.frm.sON.value = snm;
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
	<td><strong>������ �� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frm" method="post" action="procPart.asp">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="hidOEP" value="<%=iOpExpPartIdx%>">
		<input type="hidden" name="hidPsn" value="<%=iPartsn%>">
		<input type="hidden" name="hidDPid" value="<%= iDepartmentID %>"> 
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF">
				<select name="selPT" onChange="jsChPT(this.value);">
				 <% sbOptPartType arrType,iPartTypeIdx%>
				 <option value="0">--�����߰�--</option>
				</select>
				<span id="divPT" style="display:<%IF isArray(arrType) THEN%>none<%END IF%>;"><input type="text" name="sPTN" size="20" maxlength="60"></span>
			</td>
		</tr>
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>" align="center">�����ó</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sPN" size="50" maxlength="60" value="<%=sOpExpPartName%>" class="text">
			</td>
		</tr>
		<!--<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">�μ�</td>
			<td bgcolor="#FFFFFF"> <input type="button" class="button" value="�߰�" onClick="jsAddPart();"><br><br>
			<div id="divPart">
		<% 	IF isArray(arrPartsn) THEN
				FOR intLoop = 0 To UBound(arrPartsn,2)
					%>
					<div id="dP<%=arrPartsn(0,intLoop)%>"><%=arrPartsn(1,intLoop)%> <a href="javascript:jsDelPart(<%=arrPartsn(0,intLoop)%>);">[X]</a></div>
					<%
				NEXT
			END IF
			%>
			</div>
			</td>
		</tr>-->
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">�μ�NEW</td>
			<td bgcolor="#FFFFFF">
				<input type="button" class="button" value="�߰�" onClick="jsAddDepartment();"><br><br>
				<div id="divDepartment">
					<%
					if isArray(arrDepartmentID) then
						FOR intLoop = 0 To UBound(arrDepartmentID,2)
					%>
					<div id="dDP<%=arrDepartmentID(0,intLoop)%>"><%=arrDepartmentID(1,intLoop)%> <a href="javascript:jsDelDepartment(<%=arrDepartmentID(0,intLoop)%>);">[X]</a></div>
					<%
						NEXT
					END IF
					%>
				</div>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">�����</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" name="hidAI" id="hidAI" value="<%=sadminid%>">
			<input type="hidden" name="hidAJ" id="hidAJ" value="<%=ijobsn%>">
			<input type="text" name="sAN" id="sAN" size="30" maxlength="32" value="<%=susername&" "&sjobname%>" readonly style="border:0;" class="text"> &nbsp;<input type="button" name="btnID" value="����� ���" onClick="jsRegID('<%=iregdepartmentid%>','<%=sadminid%>');" class="button"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">���ްŷ�ó �ڵ�</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="hidcustcd" value="<%=scust_cd%>" size="30"  class="text"> 
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">���ްŷ�ó</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="scustnm" value="<%=scustnm%>" size="30"  readonly class="text_ro"> <a href="javascript:jsGetCust();"> <img src="/images/icon_search.jpg" border="0"></a>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">���ް�������</td>
			<td bgcolor="#FFFFFF">
			���� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:
			<input type="text" name="selOB" value="<%=iOutBank%>" readonly class="text_ro">
			 <Br>
			���¹�ȣ :<input type="text" name="sOAN" value="<%=sOutAccNo%>" size="30" readonly class="text_ro"><Br>
			�����ָ� :<input type="text" name="sON" value="<%=sOutName%>" size="20" readonly class="text_ro">
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">�����׸�(�⺻)</td>
			<td bgcolor="#FFFFFF"><input type="hidden" name="dAC" value="<%=sarap_cd%>">	<input type="text" name="sANM" value="<%=sarap_nm%>" size="20" class="text_ro">
				<a href="javascript:jsGetARAP();"> <img src="/images/icon_search.jpg" border="0"></a>
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">�ڱݰ����μ�</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="selUP" value="<%=sbizcd%>" size="15"  class="text_ro"> <input type="text" name="sPNM" value="<%=sbiznm%>" class="text_ro" size="15">
				<a href="javascript:jsGetPart();"> <img src="/images/icon_search.jpg" border="0"></a>
				</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">ī�����</td>
			<td bgcolor="#FFFFFF">
			<select name="selCCo">
				<option value="">--����--</option>
			 <%IF isArray(arrC) THEN
			 		For intC = 0 To Ubound(arrC,2)
			 	%>
			 	<option value="<%=arrC(1,intC)%>" <%IF sCardCo = arrC(1,intC) THEN%>selected<%END IF%>><%=arrC(1,intC)%></option>
			 <%	Next
			 END IF%>
			</select>
			 <input type="text" name="sCNo" value="<%=sCardNo%>" size="22" maxlength="20"  class="text_ro">
			</td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">ǥ�ü���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iON" value="<%=iOrderNo%>" size="4" style="text-algin:right;"></td>
		</tr>
		<%IF sMode="U" THEN%>
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">��뿩��</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoU" value="1" <%IF blnUsing THEN%>checked<%END IF%>>��� <input type="radio" value="0"  name="rdoU"  <%IF not blnUsing THEN%>checked<%END IF%>>������</td>
		</tr>
		<%END IF%>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td align="center"><input type="button" value="���" class="button" onClick="jsPartSubmit();"></td>
</tr>
</table>
<!-- ������ �� -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
