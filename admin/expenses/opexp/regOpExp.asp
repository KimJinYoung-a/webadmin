<%@ language=vbscript %>
<% option explicit  %>
<%
'###########################################################
' Description : ������  ����
' History : 2011.05.30 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%

Dim sMode
Dim clsPart, clsAccount, arrAccount ,clsOpExp, clsPartMoney
Dim arrList, intLoop
Dim intY, dYear, intM, dMonth
Dim iPartTypeIdx, dYYYYMM,iOpExpPartIdx, iOpExpDailyIdx
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim arrUsePart,sOpExpPartName, sPartTypeName
Dim dYYYYMMDD,iarap_cd,minExp,mOutExp,sOpExpObj,sDetailCOnts,sbizsection_cd,sbizsection_nm,msupExp,mvatExp,sauthNo ,blnIntOut
Dim blnAdmin, blnWorker ,blnReg


 	sMode = "I"
	dYear = requestCheckvar(Request("selY"),10)
	IF dYear = "" THEN dYear = year(date())
	dMonth= requestCheckvar(Request("selM"),10)
	IF dMonth = "" THEN dMonth = month(date())

 	iPartTypeIdx = requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx = requestCheckvar(Request("selP"),10)

 	IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF iOpExpPartIdx = "" THEN iOpExpPartIdx = 0

 	iOpExpDailyIdx = 	requestCheckvar(Request("hidOED"),10)
 	IF iOpExpDailyIdx = "" THEN iOpExpDailyIdx = 0


 	'�����ʱⰪ ����--------------
 	blnWorker = 0 '�����
 	blnReg = 0 	'��ϱ���
 	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '���α���
 	IF blnAdmin THEN blnReg = 1 '���α��� ���� ��� ���ó�� �׻� ����

 '��� ���ó
 	IF iOpExpPartIdx > 0 THEN
Set clsPart = new COpExpPart
		clsPart.FOpExpPartidx = iOpExpPartIdx
		clsPart.fnGetOpExpPartName
		sOpExpPartName =clsPart.FOpExpPartName
		sPartTypeName  =clsPart.FPartTypeName
Set clsPart = nothing
 END IF

'��� ���ϸ� ����Ʈ
IF iPartTypeIdx > 0 THEN
set clsOpExp = new OpExp
	clsOpExp.FYYYYMM 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx
	arrList = clsOpExp.fnGetOpExpDailyList
	iTotCnt = clsOpExp.FTotCnt

	clsOpExp.FadminID = session("ssBctId")
	clsOpExp.FPart_sn = session("ssAdminPsn")

IF iOpExpDailyIdx > 0 THEN
	sMode ="U"

	clsOpExp.FMode  = sMode
	blnWorker = clsOpExp.fnGetOpExpAuth

 	IF  blnWorker = 1   THEN	'������̰ų� ���α����� ���� ��� ���ó�� ����
		blnReg =1
	END IF

	IF blnReg=0 THEN
		set clsOpExp = nothing
			Call Alert_close ("���������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���")
		response.end
	END IF

	clsOpExp.FOpExpDailyIdx=iOpExpDailyIdx
	clsOpExp.fnGetOpExpDailyData
	dYYYYMMDD 		= clsOpExp.FYYYYMMDD
	iOpExpPartIdx 	= clsOpExp.FOpExpPartIdx
	iarap_cd		= clsOpExp.Farap_cd
	minExp 			= clsOpExp.FinExp
	mOutExp 		= clsOpExp.FOutExp
	sOpExpObj 		= clsOpExp.FOpExpObj
	sDetailCOnts 	= clsOpExp.FDetailCOnts
	sbizsection_cd= clsOpExp.Fbizsection_cd
	msupExp 		= clsOpExp.FsupExp
	mvatExp 		= clsOpExp.FvatExp
	sauthNo			= clsOpExp.FauthNo
	blnIntOut		= clsOpexp.Finouttype
	sbizsection_nm = clsOpExp.Fbizsection_nm
ELSE
	clsOpExp.FMode  = sMode
	blnWorker  = clsOpExp.fnGetOpExpAuth

	IF blnWorker = 1   THEN	'������̰ų� ���α����� ���� ��� ���ó�� ����
		blnReg =1
	END IF
END IF
set clsOpExp = nothing

END IF

 '�����׸� ����Ʈ
set clsAccount = new COpExpAccount
	clsAccount.FOpExpPartIdx = iOpExpPartIdx
	arrAccount = clsAccount.fnGetArapRegList
set clsAccount = nothing
%>
 <script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript" src="/js/datetime.js"></script>
<script language="javascript">
<!--
	//�˻�
	function jsSearch(){
		document.frmReg.action = "regOpExp.asp";
		document.frmReg.submit();
	}

 	//���
 	function jsAddOpExp(){
 	  if((document.frmReg.selA.value==0)){
 		alert("�����׸��� �������ּ���");
 		return;
 		}


 		if(jsChkBlank(document.frmReg.iD.value)){
 		alert("��¥�� �Է����ּ���");
 		document.frmReg.iD.focus();
 		return;
	 		}

 		if(!IsDigit(document.frmReg.iD.value)){
 		alert("��¥�� ���ڸ� �Է°����մϴ�.");
 		document.frmReg.iD.focus();
 		return;
 		}

 		if(!isValidDay("<%=dYear%>","<%=dMonth%>",document.frmReg.iD.value)){
 			alert("�������� �ʴ� ��¥�Դϴ�.");
 			document.frmReg.iD.focus();
 			return;
 		}

 		if(jsChkBlank(document.frmReg.mExp.value)){
 		alert("�ݾ��� �Է����ּ���");
 		document.frmReg.mExp.focus();
 		return;
 		}

 		if(!IsInteger(document.frmReg.mExp.value)){
 		alert("�ݾ��� ���ڸ� �Է°����մϴ�.");
 		document.frmReg.mExp.focus();
 		return;
 		}
 		document.frmReg.action ="procOpExp.asp"
 		document.frmReg.submit();
 	}

 	//����
 	function jsModOpExp(idx){
 		document.frmReg.hidOED.value= idx;
 		document.frmReg.action ="regOpExp.asp" ;
 		document.frmReg.submit();
 	}

 	//����
 	function jsDelOpExp(idx){
 		if(confirm("�����Ͻðڽ��ϱ�?")){
 			document.frmDel.hidOED.value = idx;
 			document.frmDel.submit();
 		}
 	}


	//���
	function jsReset(){
		document.frmReg.hidOED.value= 0;
		document.frmReg.action = "regOpExp.asp";
		document.frmReg.submit();
	}


    function jsSetExp(iType){
    	var sellExp = document.frmReg.mExp.value;

    	if(iType==1){ //�ǸŰ��� ���ް�,�ΰ��� �ڵ� ���ó��
	    	document.frmReg.msupExp.value =   parseInt((sellExp/1.1).toFixed(5)) ;
	    	document.frmReg.mvatExp.value = sellExp - document.frmReg.msupExp.value;
    	}else if(iType==2){ //���ް��� �ΰ��� �ڵ����ó��
    		document.frmReg.mvatExp.value = sellExp - document.frmReg.msupExp.value;
    	}else if(iType==3){ //�ΰ����� ���ް� �ڵ����ó��
    		document.frmReg.msupExp.value = sellExp - document.frmReg.mvatExp.value;
    	}
    }

  	//�ڱݰ����μ� ����
	function jsGetPart(){
			var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winP.focus();
	}

	//�ڱݰ����μ� ���
	function jsSetPart(sBcd, sBnm){
			document.frmReg.sBcd.value = sBcd;
			document.frmReg.sBnm.value = sBnm;
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" >
<form name="frmDel" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=ipartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmReg" method="get" action="procOpExp.asp">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="hidOED" value="<%=iOpExpDailyIdx%>">
		<input type="hidden" name="iCP" value="<%=iCurrpage%>">
		<input type="hidden" name="dOYM" value="<%=dYear%>-<%=dMonth%>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
			<td align="left"> 
				<select name="selY" class="select">
				<%For intY = Year(date()) To 2011 STEP -1%>
				<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>��
				 <select name="selM" class="select">
				<%For intM = 1 To 12%>
				<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
				<%Next%>
				</select>��  
				&nbsp;&nbsp;
				 �����ó:&nbsp;
				 <%=sPartTypeName%> > <%=sOpExpPartName%>
				<input type="hidden" name="selPT" value="<%=ipartTypeIdx%>">
				<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
				</td> 
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td> 
			</td>
		</tr>
		</table>
	</td>
</tr>
<%IF ( blnReg = 1 )  THEN%>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
		 	<td>��</td>
			<td>�����׸�</td>
			<td>��ü��</td>
			<td>����(�󼼳���)</td>
			<td>�ݾ�</td>
			<td>���ް���</td>
			<td>�ΰ���</td>
			<td>���ι�ȣ</td>
			<td>���μ�</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
		 	<td>
				<input type="text" name="iD" size="2" value="<%IF dYYYYMMDD<> "" THEN%><%=day(dYYYYMMDD)%><%END IF%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }">��
			</td>
			<td>
				<select name="selA" class="select">
				<option value ="0">--����--</option>
				<%
				Dim intA
				If isArray(arrAccount) THEN
					For intA = 0 To UBound(arrAccount,2)
					''IF arrAccount(2,intA)<>False then  ''2013/08/12 ������ �߰� ������ ����
					%>
					<option value="<%=arrAccount(0,intA)%>^<%=arrAccount(2,intA)%>" <%IF Cstr(arrAccount(0,intA)) = Cstr(iarap_cd) THEN%>selected<%END IF%>><%=arrAccount(1,intA)%><%=chkIIF(arrAccount(2,intA)=False,"(���޾�)","")%></option>
					<%
				    ''end if
					Next
				END IF %>
				</select>
			</td>
			<td><input type="text" name="sO" size="20" value="<%=sOpExpObj%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><input type="text" name="sDC" size="40" maxlength="200" value="<%=sDetailCOnts%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><input type="text" name="mExp" size="10" style="text-align:right;" value="<%IF not blnIntOut THEN%><%=minExp%><%ELSE%><%=moutExp%><%END IF%>" onkeyup="jsSetExp(1);" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><input type="text" name="msupExp" size="10" style="text-align:right;" value="<%=msupExp%>"  onkeyup="jsSetExp(2);" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><input type="text" name="mvatExp" size="10" style="text-align:right;" value="<%=mvatExp%>" onkeyup="jsSetExp(3);" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><input type="text" name="sAN" size="10" maxlength="30" value="<%=sauthNo%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td> <input type="hidden" name="sBcd" value="<%=sbizsection_cd%>"><input type="text" name="sBnm" size="10" value="<%=sbizsection_nm%>" class="text_ro" readonly>	<a href="javascript:jsGetPart();"><img src="/images/icon_search.jpg" border="0"></a></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td align="center">
			<%IF sMode="U" THEN%>
			<input type="button" class="button" value="����" style="width:80px;color:blue;" onClick="jsAddOpExp();">
			<input type="button" class="button" value="���" style="width:80px;" onClick="jsReset();">
			<%ELSE%>
			<input type="button" class="button" value="���" style="width:80px;color:blue;" onClick="jsAddOpExp();">
			<%END IF%>
			</td>
		</tr>
		</table>
	</td>
	</form>
</tr>
<script language="javascript">
 window.onload = function(){
 	document.frmReg.iD.focus();
 	}
</script>
<%ELSEIF blnWorker = 2 THEN%>
	<tr>
	<td> <font color="red">- �����Ͻ� ���� �����޿�� �ۼ����Դϴ�. ������ ��� �ۼ��Ϸ� �� �ۼ����ּ���</font></td>
</tr>
<%ELSE%>
<tr>
	<td> <font color="red">- �ۼ��Ϸ�Ǿ� ����� �Ұ����ϰų� ��� ������ �����ϴ�.</font></td>
</tr>
<%END IF%>
<tr>
	<td>
		<div id="divList" style="height:600px;overflow:scroll;">
		<b> [ <%=dYear%>�� <%=dMonth%>�� ��� �󼼳��� -  <%=sOpExpPartName%> ]</b>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
						<td width="50">����</td>
						<td width="50">��¥(��)</td>
						<td>�����ó</td>
						<td>�����׸�</td>
						<td>��ü��</td>
						<td>����(�󼼳���)</td>
						<td>���޾�</td>
						<td>����</td>
						<td>���ް���</td>
						<td>�ΰ���</td>
						<td>���ι�ȣ</td>
						<td>���μ�</td>
						<td width="100">ó��</td>
					</tr>
					<%   Dim totInExp, totOutExp,sumInExp,sumOutExp, iNum, sumSupExp, sumVatExp, totSupExp, totVatExp
					totInExp = 0
					totOutExp = 0
					sumInExp=0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					totSupExp=0
					totVatExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					 %>
					<tr height=30 bgcolor="<%IF Cstr(arrList(0,intLoop))= Cstr(iOpExpDailyIdx) THEN%><%=adminColor("green")%><%ELSE%>#FFFFFF<%END IF%>">
						<td align="center"><%=iNum%></td>
						<td align="center"><%=day(arrList(1,intLoop))%></td>
						<td align="center"><%=arrList(12,intLoop)%> > <%=arrList(11,intLoop)%></td>
						<td align="center"><%=arrList(3,intLoop)%></td>
						<td><%=arrList(6,intLoop)%></td>
						<td><%=arrList(7,intLoop)%></td>
						<td align="right"><%=formatnumber(arrList(4,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(5,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td>
						<td align="center"><%=arrList(10,intLoop)%></td>
						<td align="center"><%=arrList(13,intLoop)%></td>
						<td align="center">
						<% if IsNULL(arrList(20,intLoop)) then %>
							<%IF ( blnReg = 1)   THEN%>
							<input type="button" class="button" value="����" onClick="jsModOpExp(<%=arrList(0,intLoop)%>);">
							<input type="button" class="button" value="����" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
							<%END IF%>
						<% else %>
						    <%= arrList(20,intLoop) %>
						<% end if %>
						</td>
					</tr>
					<%
					  totInExp = totInExp + arrList(4,intLoop)
					  totOutExp = totOutExp + arrList(5,intLoop)
					  totSupExp = totSupExp + arrList(8,intLoop)
					  totVatExp = totVatExp + arrList(9,intLoop)

					  sumInExp = sumInExp +  arrList(4,intLoop)
					  sumOutExp = sumOutExp +  arrList(5,intLoop)
					  sumSupExp = sumSupExp +  arrList(8,intLoop)
					  sumVatExp = sumVatExp +  arrList(9,intLoop)

					  iNum = iNum + 1
				IF intLoop  < UBound(arrList,2)  THEN
				 	IF Cstr(arrList(2,intLoop)) <> Cstr(arrList(2,intLoop+1)) THEN%>
				   <tr height=30 align="center" bgcolor="#FFFFFF">
				   	<td colspan="6"><b><%=arrList(3,intLoop)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				    <td colspan="4"></td>
				</tr>
				<%	sumInExp = 0
					sumOutExp = 0
					sumSupExp = 0
					sumVatExp = 0
					iNum = 1
					END IF
				END IF
					Next  %>
					<tr  height=30 align="center" bgcolor="#FFFFFF">
				   	<td colspan="6"><b><%=arrList(3,intLoop-1)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				   	<td colspan="4"></td>
					</tr>
					<%
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">
						<td colspan="14">��ϵ� ������ �����ϴ�.</td>
					</tr>
					<%END IF%>

				 <tr  height=30 align="center" bgcolor="#DDDDFF">
				   	<td colspan="6">����</td>
				   	<td align="right"><%=formatnumber(totInExp,0)%></td>
				   	<td align="right"><%=formatnumber(totOutExp,0)%></td>
				   	<td align="right"><%=formatnumber(totSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(totVatExp,0)%></td>
				   	<td colspan="4"></td>
				</tr>
				</table>
			</td>
		</tR>
		</div>
	</td>
</tr>
</table>
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->



