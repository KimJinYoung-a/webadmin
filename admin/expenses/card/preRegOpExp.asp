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
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim sMode, selA,arap_nm
Dim clsPart, clsAccount, arrAccount ,clsOpExp, clsPartMoney
Dim arrList, intLoop
Dim intY, dYear, intM, dMonth
Dim  dYYYYMM,iPartTypeIdx,iOpExpPartIdx, iOpExpDailyIdx, dauthDate,msevExp,blndeducttype
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim arrUsePart,sOpExpPartName, sPartTypeName
Dim  iarap_cd,minExp,mOutExp,sOpExpObj,sDetailCOnts,sbizsection_cd,sbizsection_nm,msupExp,mvatExp,sauthNo ,blnIntOut
Dim blnAdmin, blnWorker ,blnReg
Dim  ipartsn,sadminid
Dim idefaultArap_cd, intA

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
 	IF not blnAdmin THEN  '����Ʈ ������ ���� ����� �����ϰ� ����ڿ� ���μ�  view ����
		ipartsn  =  session("ssAdminPsn")
 		sadminid = 	session("ssBctId")
 	END IF
 '��� ���ó
 	IF iOpExpPartIdx > 0 THEN
Set clsPart = new COpExpPart
		clsPart.FOpExpPartidx = iOpExpPartIdx
		clsPart.fnGetOpExpPartName
		sOpExpPartName =clsPart.FOpExpPartName
		sPartTypeName  =clsPart.FPartTypeName

Set clsPart = nothing
Set clsPart = new COpExpPart
		clsPart.FOpExpPartidx = iOpExpPartIdx
        clsPart.fnGetOpExpPartData
        idefaultArap_cd=clsPart.Farap_cd
Set clsPart = nothing
 END IF

'��� ���ϸ� ����Ʈ
set clsOpExp = new OpExp
	clsOpExp.FSAuthDate 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FEAuthDate 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FPartTypeIdx = iPartTypeIdx
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx
	clsOpExp.FRectPartsn = ipartsn
	clsOpExp.FRectUserid = sadminid
	arrList = clsOpExp.fnGetOpExpDailyNoSetList
	iTotCnt = clsOpExp.FTotCnt

	clsOpExp.FadminID = session("ssBctId")
	clsOpExp.FPart_sn = session("ssAdminPsn")
  blnWorker = clsOpExp.fnGetOpExpPartAuth
  IF blnWorker = 1 THEN blnReg = 1
	IF blnReg=0 THEN
		set clsOpExp = nothing
			Call Alert_close ("���������� �����ϴ�. Ȯ�� �� �ٽ� �õ����ּ���")
		response.end
	END IF
IF iOpExpDailyIdx > 0 THEN
	sMode ="U"
	clsOpExp.FOpExpDailyIdx=iOpExpDailyIdx
	clsOpExp.fnGetOpExpDailyData
	dYYYYMM 		= clsOpExp.FYYYYMM
	dauthDate		= clsOpExp.Fauthdate
	iOpExpPartIdx = clsOpExp.FOpExpPartIdx
	iarap_cd			= clsOpExp.Farap_cd
	mOutExp 			= clsOpExp.FOutExp
	sOpExpObj 		= clsOpExp.FOpExpObj
	sDetailCOnts 	= clsOpExp.FDetailCOnts
	sbizsection_cd= clsOpExp.Fbizsection_cd
	msupExp 			= clsOpExp.FsupExp
	mvatExp 			= clsOpExp.FvatExp
	msevExp				= clsOpExp.FsevExp
	sauthNo				= clsOpExp.FauthNo
	blndeducttype	= clsOpExp.Fdeducttype
	blnIntOut			= clsOpexp.Finouttype
	sbizsection_nm= clsOpExp.Fbizsection_nm

END IF
set clsOpExp = nothing

 IF isNull(blndeducttype) THEN blndeducttype = False

 '�����׸� ����Ʈ
set clsAccount = new COpExpAccount
	clsAccount.FOpExpPartIdx = iOpExpPartIdx
	arrAccount = clsAccount.fnGetArapRegList
set clsAccount = nothing

'' �⺻ ���� �׸� �Է�
if (iarap_cd="") or (iarap_cd="0")then
    if (idefaultArap_cd="625") or (idefaultArap_cd="640") then
        iarap_cd = idefaultArap_cd
    end if
end if
%>
 <script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript" src="/js/datetime.js"></script>
<script language="javascript">
<!--
	//�˻�
	function jsSearch(){
		document.frmReg.action = "preRegOpExp.asp";
		document.frmReg.submit();
	}

 	//���
 	function jsAddOpExp(){
 	  if((document.frmReg.selA.value==0)){
 		alert("�����׸��� �������ּ���");
 		return;
 		}
      <% 'if (idefaultArap_cd="928") then %> //�ָ��߱ٽĴ�
// �ּ�ó��. �繫�� ������ ��û
//      if((document.frmReg.selA.value!='<%'=idefaultArap_cd%>')){
//        alert('������ �����׸� ��� �����մϴ�.');
//        document.frmReg.selA.focus();
//        return;
//      }
      <% 'end if %>
      <% 'if (idefaultArap_cd="625")  then %> //����Ȱ��ȭ��, ���Ĵ� 
// �ּ�ó��. �繫�� ������ ��û
//      if(!(document.frmReg.selA.value==625 || document.frmReg.selA.value==927 || document.frmReg.selA.value ==940)){
//        alert('������ �����׸� ��� �����մϴ�.');
//        document.frmReg.selA.focus();
//        return;
//      }
      <% 'end if %>
 		document.frmReg.action ="procOpExp.asp"
 		document.frmReg.submit();
 	}

 	//����
 	function jsModOpExp(idx){
// 	      <% if (idefaultArap_cd="625") or (idefaultArap_cd="640") then %>
//          if((document.frmReg.selA.value!='<%=idefaultArap_cd%>')){
//            alert('������ �����׸� ��� �����մϴ�.');
//            document.frmReg.selA.focus();
//            return;
//          }
//          <% end if %>

 		document.frmReg.hidOED.value= idx;
 		document.frmReg.action ="preRegOpExp.asp" ;
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
		document.frmReg.action = "preRegOpExp.asp";
		document.frmReg.submit();
	}

	// �����׸�
	function jsGetarap_cd(iOpExpPartIdx){
		if (iOpExpPartIdx==''){
			alert('�˻�Ű�� �����ϴ�.;');
			return;
		}

		var winarap_cdP = window.open('/admin/linkedERP/Biz/poparap_cdone.asp?selP='+ iOpExpPartIdx +'&menupos=<%= menupos %>','poparap_cd','width=600, height=500, resizable=yes, scrollbars=yes');
		winarap_cdP.focus();
	}

	// �����׸� ���
	function jsSetarap_cd(arap_cd, arap_nm){
		document.frmReg.selA.value = arap_cd;
		document.frmReg.arap_nm.value = arap_nm;
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
	//��볻�� ���ϵ��
	function jsSetFile(){
			var sYear = document.frmReg.selY.options[document.frmReg.selY.selectedIndex].value;
			var sMonth = document.frmReg.selM.options[document.frmReg.selM.selectedIndex].value;
			var winF = window.open('/admin/expenses/opexp/popRegFile.asp?selY='+sYear+'&selM='+sMonth+'&selP=<%=iOpExpPartIdx%>','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winF.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" >
<form name="frmDel" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
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
		<input type="hidden" name="hidNS" value="Y">
		<input type="hidden" name="hidRU" value="preRegOpExp.asp">
		<input type="hidden" name="mO"  value="<%=moutExp%>">
		<input type="hidden" name="mSP"  value="<%=msupExp%>">
		<input type="hidden" name="mV"  value="<%=mvatExp%>">
		<input type="hidden" name="mSV" value="<%=msevExp%>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
			<td align="left">
				������:
				 <%IF sMode="U" THEN%>
				<input type="hidden" name="selY" value="<%=dYear%>">
				<input type="hidden" name="selM" value="<%=dMonth%>">
				<%=dYear%>�� <%=dMonth%>��
				<%ELSE%>
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
				<%END IF%>
				&nbsp;&nbsp;
				 �����ó:&nbsp;
				   <%=sPartTypeName%> > <%=sOpExpPartName%>
				  <input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
				<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
				</td>
				<%IF sMode="I" THEN%>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
				<%END IF%>
			</td>
		</tr>
		</table>
	</td>
</tr>
 <%IF  sMode="U"  THEN%>
<%IF ( blnReg = 1  ) THEN%>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td>�����׸�</td>
			<td>��ü��</td>
			<td>����(�󼼳���)</td>
			<td>�ݾ�</td>
			<td>���ι�ȣ</td>
			<td>��������</td>
			<td>���μ�</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>
				<%
				If isArray(arrAccount) THEN
					For intA = 0 To UBound(arrAccount,2)
						IF Cstr(arrAccount(0,intA)) = Cstr(iarap_cd) THEN
							selA=arrAccount(0,intA)
							arap_nm=chkIIF(arrAccount(2,intA),"[���]","[����]") & arrAccount(1,intA)
						end if
					Next
				END IF
				%>
				<input type="hidden" name="selA" value="<%= selA %>">
				<input type="text" name="arap_nm" size="20" value="<%= arap_nm %>" class="text_ro" readonly>
				<a href="#" onclick="jsGetarap_cd('<%= iOpExpPartIdx %>'); return false;"><img src="/images/icon_search.jpg" border="0"></a>
			</td>
			<td><%=sOpExpObj%></td>
			<td><input type="text" name="sDC" size="50" maxlength="200" value="<%=sDetailCOnts%>" onKeyDown="javascript:if (event.keyCode == 13) {jsAddOpExp(); }"></td>
			<td><%=formatnumber(moutExp,0)%></td>
			<td><%=sauthNo%></td>
			<td><input type="radio"  name="rdoD" value="1" <%IF blndeducttype THEN%>checked<%END IF%>>Y &nbsp;
				 <input type="radio"  name="rdoD" value="0" <%IF not blndeducttype THEN%>checked<%END IF%>>N</td>
			<td><input type="hidden" name="sBcd" value="<%=sbizsection_cd%>"><input type="text" name="sBnm" size="10" value="<%=sbizsection_nm%>" class="text_ro" readonly>	<a href="javascript:jsGetPart();"><img src="/images/icon_search.jpg" border="0"></a></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td align="center">
			<input type="button" class="button" value="����" style="width:80px;color:blue;" onClick="jsAddOpExp();">
			<input type="button" class="button" value="���" style="width:80px;" onClick="jsReset();">
			</td>
		</tr>
		</table>
	</td>
	</form>
</tr>
<%ELSE%>
<tr>
	<td> <font color="red">- �ۼ��Ϸ�Ǿ� ����� �Ұ����ϰų� ��� ������ �����ϴ�.</font></td>
</tr>
<%END IF%>
	<%END IF%>
<tr>
	<td>
		<div id="divList" style="height:600px;overflow:scroll;">
		<b> [ <%=dYear%>�� <%=dMonth%>�� ����ī���� �󼼳��� - <%=sPartTypeName%> > <%=sOpExpPartName%>   ]</b>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
						<td width="50">����</td>
						<td width="50">������</td>
						<td>�����ó</td>
						<td>�����׸�</td>
						<td>��ü��</td>
						<td>����(�󼼳���)</td>
						<td>����</td>
						<td>���ް���</td>
						<td>�ΰ���</td>
						<td>�����</td>
						<td>���ι�ȣ</td>
						<td>��������</td>
						<td>����/��</td>
						<td>���μ�</td>
						<td>��������</td>
						<td width="100">ó��</td>
					</tr>
				<%
					Dim  totOutExp, sumOutExp, iNum, sumSupExp, sumVatExp, sumSevExp, totSupExp, totVatExp, totSevExp
					totOutExp = 0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					sumSevExp=0
					totSupExp=0
					totVatExp=0
					totSevExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					 %>
					<tr height=30 bgcolor="<%IF Cstr(arrList(0,intLoop))= Cstr(iOpExpDailyIdx) THEN%><%=adminColor("green")%><%ELSE%><%= CHKIIF(arrList(22,intLoop)=0,"#CCCCCC","#FFFFFF") %><%END IF%>">
						<td align="center"><%=iNum%></td>
						<td align="center"><%=formatdate(arrList(2,intLoop),"0000-00-00")%></td>
						<td align="center"><%=arrList(15,intLoop)%></td>
						<td align="center"><%=arrList(5,intLoop)%></td>
						<td><%=arrList(11,intLoop)%></td>
						<td><%=arrList(12,intLoop)%></td>
						<td align="right"><%=formatnumber(arrList(6,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(7,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td>
						<td align="center"><%=arrList(10,intLoop)%></td>
						<td align="center"><%=arrList(16,intLoop)%></td>
						<td align="center"><%IF arrList(19,intLoop)=1 THEN%>����<%ELSE%>����<%END IF%></td>
						<td align="center"><%=arrList(14,intLoop)%></td>
						<td align="center"><%IF arrList(17,intLoop) THEN%><font color="red">Y</font><%ELSE%><font color="blue">N</font><%END IF%></td>
						<td align="center">
						<% if IsNULL(arrList(21,intLoop)) then %>
						<%IF blnReg = 1 THEN%>
						    <% if (arrList(22,intLoop)<>0) then %>
							<input type="button" class="button" value="����" onClick="jsModOpExp(<%=arrList(0,intLoop)%>);">
							<% end if %>
							<%IF blnAdmin THEN%>
							<% if (arrList(22,intLoop)=0) then %>
						    <!-- input type="button" class="button" value="����" onClick="jsLiveOpExp(<%=arrList(0,intLoop)%>)" -->
						    <% else %>
							<input type="button" class="button" value="����" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
							<% end if %>
							<% end if %>
						<%END IF%>
						<% else %>
						    <%= arrList(21,intLoop) %>
						<% end if %>
						</td>
					</tr>
					<%
					  iNum = iNum + 1
			 	Next
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">
						<td colspan="16">��ϵ� ������ �����ϴ�.</td>
					</tr>
					<%END IF%>

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



