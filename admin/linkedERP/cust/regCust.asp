<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ŷ�ó ����
' History : 2011.04.21 ������ ����
'			2016.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/custCls.asp"-->
<%
Dim clsC, arrBank, intB,sMode, sEmail, sDispSeq, sPostCd, sAddr,sCORPYN,sBRNType
Dim sCustCd, sCustNm, sARYN, sAPYN, sCoYN, sBizNo, sCeoNm, sTaxType, sBSCD, sINTP, sTelNo, sFaxNo
Dim sEmpno,sEmpNm, sPos,sDeptNM, sEmpTel, sEmpHP, sEmpEmail,sPSGB
Dim arrBankNm, arrBankCd, arrSavMN,arrAcctNo, intLoop, intDef
Dim sBank_cd,sacc_no, ssav_mn, arrEmp,arrAcct
Dim sCustNm7,sBizNo7,sPos7,sDeptNm7,sEmpTel7,sEmpHP7,sEmpEmail7,sDispSeq7,sEmpNm7
dim srectEmpno,srectBankno,srectAcctno
Dim isReadOnly : isReadOnly = (requestCheckvar(Request("rO"),10)<>"")
	sCustCd = requestCheckvar(Request("hidCcd"),13)
	srectEmpno= requestCheckvar(Request("hidEno"),10)
	srectBankno= requestCheckvar(Request("hidBno"),8)
	srectAcctno= requestCheckvar(Request("hidAno"),30)
sPSGB = "1"
sMode ="I"

set clsC = new CCust

	IF sCustCd <> "" THEN
	 	sMode = "U"
		clsC.FCustCd = sCustCd
		clsC.FRectempno = srectEmpno
		clsC.FRectBankNo = srectBankno
		clsC.FRectAcctNo = srectAcctno
		clsC.fnGetCustData
		sCORPYN =clsC.FCORPYN
		sBRNType =clsC.FCUSTBRNTYPE
		sPSGB = clsC.FPSGB
		IF isNull(sCORPYN) or sCORPYN ="" then sCORPYN ="Y"
		IF (sCORPYN="Y") then sPSGB="1" ''�����̸� ������ �����. // ERP���� ������ ���� �̸� ������ ��..

		IF sPSGB = "2" THEN	'�����϶�
	  		sCustNm7	= clsC.FCustNM
			sBizNo7		= clsC.FBizNo
			sPos7		= clsC.FPos
			sDeptNm7	= clsC.FDEPT_NM
			sEmpTel7	= clsC.FTELNO
			sEmpHP7		= clsC.FHP_NO
			sEmpEmail7	= clsC.FEMAIL
			sDispSeq7	= clsC.FDispSeq
			sEmpNm7 	= clsC.FEMP_NM
		ELSE
			sCustNM 	= clsC.FCustNM
			sBizNo  	= clsC.FBizNo
			sDispSeq	= clsC.FDispSeq
			sPos 		= clsC.FPos
			sDEPTNM 	= clsC.FDEPT_NM
			sEmpTel 	= clsC.FSTelNo
			sEmpHP 		= clsC.FHP_NO
			sEmpEmail 	= clsC.FSEmail
			sEmpNm 		= clsC.FEMP_NM
		END IF

		sARYN		= clsC.FARYN
		sAPYN	 	= clsC.FAPYN
		sCeoNM  	= clsC.FCeoNM
		sEMAIL  	= clsC.FEMAIL
		sTELNO  	= clsC.FTELNO
		sFAXNO  	= clsC.FFAXNO
		sTAXTYPE	= clsC.FTAXTYPE
		sBSCD   	= clsC.FBSCD
		sINTP   	= clsC.FINTP
		sPostCD 	= clsC.FPostCD
		sADDR   	= clsC.FADDR
		sEmpNo 		= clsC.FEMP_NO

		sBank_cd	= clsC.FBank_cd
		sacc_no 	= clsC.Facct_no
		ssav_mn		= clsC.Fsav_mn

		IF isNull(sBank_cd) THEN sBank_cd = ""
		'arrEmp	=clsC.fnGetCustSaleorList	-����Ʈ ó�� ���� ���ļ���..���� 1���� �����ֵ���
		'arrAcct =clsC.fnGetCustAcctList
	END IF

	arrBank = clsC.fnGetBankList
set clsC = nothing

IF sDispSeq = "" THEN sDispSeq = "10000"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script type="text/javascript">

//�����ȣ �ҷ�����
function jsSetPC(){
	var winPC = window.open("/lib/searchzip3.asp?target=frmC","PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	winPC.focus();
}

//�����ȣ set
	function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".sPCd").value = post1 + post2;

    eval(frmname + ".sAddr").value = addr+" "+dong;
}


function jsSubmitCust(){
	if(document.frmC.rdoRT[1].checked){
		if(jsChkBlank(document.frmC.scnm7.value)){
 		alert("�ŷ�ó����  �Է����ּ���");
 		document.frmC.scnm7.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sem7.value)){
 		alert("����ڸ���  �Է����ּ���");
 		document.frmC.sem7.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sBno71.value)){
 		alert("�ֹι�ȣ��  �Է����ּ���");
 		document.frmC.sBno71.focus();
 		return;
 		}

 			if(jsChkBlank(document.frmC.sBno72.value)){
 		alert("�ֹι�ȣ��  �Է����ּ���");
 		document.frmC.sBno72.focus();
 		return;
 		}
	}else{
	if(jsChkBlank(document.frmC.scnm.value)){
 		alert("�ŷ�ó����  �Է����ּ���");
 		document.frmC.scnm.focus();
 		return;
 		}

 		if(jsChkBlank(document.frmC.sBno.value)){
 		alert("����ڹ�ȣ��  �Է����ּ���");
 		document.frmC.sBno.focus();
 		return;
 		}
 	}

    if (document.frmC.selBC.value=="10000005"){
        alert("��ȯ���� ���� �Ұ� : KEB �ϳ��������� ����Ǿ����ϴ�.");
 		document.frmC.sBno.focus();
 		return;
    }

    if (confirm('����Ͻðڽ��ϱ�?')){
    	document.frmC.submit();
    }
}

//����� ȭ�� ����
 function jsSetRegFrm(iType){
 	if(iType=="2"){
 		document.all.dType1.style.display = "none";
 		document.all.dType2.style.display = "";
 	}else{
 		document.all.dType1.style.display = "";
 		document.all.dType2.style.display = "none";
 	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
	<td><strong>�ŷ�ó �űԵ��</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frmC" method="post" action="procCust.asp">
			<input type="hidden" name="ver" value="">
			<input type="hidden" name="hidM" value="<%=sMode%>">
			<input type="hidden" name="hidCcd" value="<%=sCustCd%>">
			<input type="hidden" name="hidENo" value="<%=sEmpno%>">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">���Ÿ�� </td>
				<td bgcolor="#FFFFFF"> <input type="radio" name="rdoRT" value="1" <%IF sPSGB="1" THEN%>checked<%END IF%> onClick="jsSetRegFrm(1);">����� | <input type="radio" name="rdoRT" value="2" <%IF sPSGB="2" THEN%>checked<%END IF%> onClick="jsSetRegFrm(2);">����</td></td>
			</tr>
		</table>
</tr>
<tr>
	<td>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" >
			<tr>
				<td>
					<div id="dType1" style="display:<%IF sPSGB="2" THEN%>none<%END IF%>;"><!-- �ŷ�ó��� ��-->
					<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" border="0" >
					<tr>
				 		<td>�⺻����</td>
					</tr>
					<tr>
						<td>
							<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
						 <tr>
								<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�ŷ�ó��</td>
								<td colspan="3"  bgcolor="#FFFFFF"  ><input type="text" name="scnm" value="<%=sCustNm%>" size="20" class="text">(�������)</td>
							</tr>
							<tr>
								<td   bgcolor="<%= adminColor("tabletop") %>" align="center">�ŷ�ó�з�</td>
								<td    bgcolor="#FFFFFF">
									<select name="selBRNT">
									<%sbOptCustType sBRNType%>
									</select>
								</td>
								<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�ŷ�ó����</td>
								<td   bgcolor="#FFFFFF">
									<input type="checkbox" name="chkAR" value="Y" <%IF sARYN="Y" THEN%>checked<%END IF%>>����
									<input type="checkbox" name="chkAP" value="Y" <%IF sAPYN="Y" THEN%>checked<%END IF%>>����
								</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">����ڱ���</td>
								<td  bgcolor="#FFFFFF"><input type="radio" name="rdoCo" value="Y" <%IF sCORPYN ="Y" or sCORPYN="" THEN%>checked<%END IF%>>����
									<input type="radio" name="rdoCo" value="N" <%IF sCORPYN ="N" THEN%>checked<%END IF%>>����</td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">����ڹ�ȣ</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sBno" value="<%=sBizNo%>" size="15" maxlength="13"  <%= CHKIIF(sCORPYN="Y" and Len(replace(sBizNo,"-",""))=10,"readonly class='text_ro'","class='text'") %> > (-���� ���ڸ�)</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ǥ��</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sceonm" value="<%=sCeoNm%>" size="10" class="text"></td>
								<td  bgcolor="<%= adminColor("tabletop") %>" align="center">��������/�����</td>
								<td  bgcolor="#FFFFFF">
									<select name="selTType">
										<option value="1" <%IF sTaxType="1" or  sTaxType="" THEN%>selected<%END IF%>>�Ϲݰ�����</option>
										<option value="2" <%IF sTaxType="2" THEN%>selected<%END IF%>>���̰�����</option>
										<option value="3" <%IF sTaxType="3" THEN%>selected<%END IF%>>�鼼������</option>
										<option value="4" <%IF sTaxType="4" THEN%>selected<%END IF%>>�񿵸�������</option>
										<option value="5" <%IF sTaxType="5" THEN%>selected<%END IF%>>��Ÿ�����</option>
										<option value="9" <%IF sTaxType="9" THEN%>selected<%END IF%>>�޾�</option>
										<option value="10" <%IF sTaxType="10" THEN%>selected<%END IF%>>���</option>
									</select>
								</td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sBS" value="<%=sBSCD%>" size="20" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sIN" value="<%=sINTP%>" size="20" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȭ��ȣ</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sTNo" value="<%=sTelNo%>" size="15" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ѽ���ȣ</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sFNo" value="<%=sFaxNo%>" size="15" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̸���</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sE" value="<%=sEmail%>" size="30" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">��¼���</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sDS" value="<%=sDispSeq%>" size="5" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ּ�</td>
								<td bgcolor="#FFFFFF" colspan="3">
									<input type="text" name="sPCd" value="<%=sPostCd%>" size="6" class="text_ro">
									<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmC','G')">
									<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmC','G')">
									<% '<input type="button" class="button" value="�˻�(��)"  onClick="jsSetPC();"> %>
									<input type="text" name="sAddr" value="<%=sAddr%>" size="60" class="text">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:10px;">���������</td>
				</tR>
				<tr>
					<td>
						<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">����</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sENm" value="<%=sEmpNm%>" size="10" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">����</td>
								<td  bgcolor="#FFFFFF" width="300"><input type="text" name="sEP" value="<%=sPos%>" size="10" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�μ�</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sDNm" value="<%=sDeptNm%>" size="20" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȭ</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sETN" value="<%=sEmpTel%>" size="15" class="text"></td>
							</tr>
							<tr>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�޴���</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sEHp" value="<%=sEmpHP%>" size="15" class="text"></td>
								<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̸���</td>
								<td  bgcolor="#FFFFFF"><input type="text" name="sEE" value="<%=sEmpEmail%>" size="30" class="text"></td>
							</tr>
						</table>
					</td>
			</tr>
		</table>
		</div>
		<div id="dType2" style="display:<%IF sPSGB="1" THEN%>none<%END IF%>;"><!-- ������� ��-->
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" border="0" >
			<tr>
			<td>�⺻����</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�ŷ�ó��</td>
					<td  bgcolor="#FFFFFF"  ><input type="text" name="scnm7" value="<%=sCustNm7%>" size="20" class="text"></td>
					<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����ڸ�</td>
					<td  bgcolor="#FFFFFF"  ><input type="text" name="sem7" value="<%=sEmpNm7%>" size="20" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ŷ�ó�з�</td>
					<td bgcolor="#FFFFFF">
						<select name="selBRNT7">
						<%sbOptCustType sBRNType%>
						</select>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="100">�ֹι�ȣ</td>
					<td  bgcolor="#FFFFFF">
						<input type="text" name="sBno71" value="<%IF sBizNo7 <> "" THEN%><%=left(sBizNo7,6)%><%END IF%>" size="10" maxlength="6" class="text">-
						<input type="password" name="sBno72" value="<%IF sBizNo7 <> "" THEN%><%=mid(sBizNo7,7,7)%><%END IF%>" size="10" maxlength="7" class="text">
						</td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
					<td  bgcolor="#FFFFFF" width="300"><input type="text" name="sEP7" value="<%=sPos7%>" size="10" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�μ�</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sDNm7" value="<%=sDeptNm7%>" size="20" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȭ</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sTNo7" value="<%=sEmpTel7%>" size="15" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�޴���</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sEHp7" value="<%=sEmpHP7%>" size="15" class="text"></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̸���</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sE7" value="<%=sEmpEmail7%>" size="30" class="text"></td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center">��¼���</td>
					<td  bgcolor="#FFFFFF"><input type="text" name="sDS7" value="<%=sDispSeq7%>" size="5" class="text"></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
			</div>
		</td>
	</tr>
	<tr>
		<td style="padding-top:10px;">�������� </td>
	</tr>
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"  bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td  bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">�����</td>
					<td  bgcolor="#FFFFFF">
						<select name="selBC">
							<option value="">--����--</option>
							<%IF isArray(arrBank) THEN
								For intB = 0 To UBound(arrBank,2)
							%>
							<option value="<%=arrBank(0,intB)%>" <%IF Cstr(sBank_cd) = Cstr(arrBank(0,intB)) THEN%>selected<%END IF%>><%=arrBank(1,intB)%></option>
							<%
								Next
								END IF%>
						</select>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="80"> ���¹�ȣ:</td>
					<td   bgcolor="#FFFFFF"> <input type="text" name="sAN" value="<%=sAcc_no%>" size="20" class="text"></td>
					<td  bgcolor="<%= adminColor("tabletop") %>" align="center" width="80">������</td>
					<Td  bgcolor="#FFFFFF"><input type="text" name="sSN" value="<%=sSav_mn%>" size="10" class="text"></td>
				</tr>
			</table>
		</td>
		</tr>
		<tr>
			<td align="center" colspan="3">
			<% if Not isReadOnly then %>
				<input type="button" class="button" value="���" onClick="jsSubmitCust();">
			<% end if%>
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
</body>
</html>
